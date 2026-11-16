"""
Modele de classement appris a partir des PDF deja ranges.

L'arborescence existante sert d'exemples : chaque PDF deja range indique
« un document comme ca va dans ce dossier ». Pour un nouveau PDF, on cherche
les documents rangés dont le texte (OCR) ressemble le plus (TF-IDF + cosinus,
voisins les plus proches) et on propose leur dossier. Tout est local, sans
dependance externe : pas besoin de maintenir de mots-cles.

Le modele est mis en cache (sort_model.pkl) : seuls les PDF nouveaux ou
modifies sont relus.
"""

import collections
import logging
import math
import os
import pickle
import re
import threading

import sorter

log = logging.getLogger("autoscan.model")

_YEAR_DIR = re.compile(r"^(19|20)\d{2}$")
_TOKEN = re.compile(r"[a-z0-9]{3,}")
MIN_TEXT_CHARS = 30
CACHE_VERSION = 1


def tokens(text):
    return _TOKEN.findall(sorter.normalize(text))


class FolderModel:
    def __init__(self, cache_path):
        self.cache_path = cache_path
        self.entries = {}          # chemin -> dict(mtime, size, label, uses_year, tf)
        self._postings = {}        # mot -> [(index doc, poids)]
        self._idf = {}
        self._docs = []            # chemins indexes, dans l'ordre des postings
        self._labels = []
        self._lock = threading.Lock()
        self.ready = False

    # -- persistance -------------------------------------------------------

    def load(self):
        try:
            with open(self.cache_path, "rb") as f:
                data = pickle.load(f)
            if data.get("version") == CACHE_VERSION:
                self.entries = data["entries"]
                self._rebuild()
                self.ready = bool(self.entries)
        except FileNotFoundError:
            pass
        except Exception:
            log.exception("Cache du modele illisible, il sera reconstruit")

    def save(self):
        tmp = self.cache_path + ".tmp"
        with open(tmp, "wb") as f:
            pickle.dump({"version": CACHE_VERSION, "entries": self.entries}, f)
        os.replace(tmp, self.cache_path)

    # -- construction ------------------------------------------------------

    @staticmethod
    def label_for(root, directory):
        """(dossier sans l'annee, True si le dossier est range par annee)."""
        rel = os.path.relpath(directory, root).replace("\\", "/")
        if rel == ".":
            return None, False
        parts = rel.split("/")
        uses_year = bool(_YEAR_DIR.match(parts[-1])) and len(parts) > 1
        if uses_year:
            parts = parts[:-1]
        return "/".join(parts), uses_year

    def refresh(self, root, skip=(), exclude=()):
        """Parcourt l'arborescence : lit les PDF nouveaux ou modifies, oublie
        ceux qui ont disparu. Retourne (nombre de PDF, nombre relus)."""
        skip_norm = {sorter.normalize(s) for s in skip}
        excluded = {os.path.normcase(os.path.abspath(p)) for p in exclude}
        entries, reread = {}, 0
        for dirpath, dirs, files in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            top = rel.split(os.sep)[0]
            dirs[:] = [
                d for d in dirs
                if not d.startswith((".", "_"))
                and not (rel == "." and sorter.normalize(d) in skip_norm)
                and os.path.normcase(os.path.abspath(os.path.join(dirpath, d))) not in excluded
            ]
            if rel == "." or sorter.normalize(top) in skip_norm:
                continue
            label, uses_year = self.label_for(root, dirpath)
            for name in files:
                if not name.lower().endswith(".pdf"):
                    continue
                path = os.path.join(dirpath, name)
                try:
                    st = os.stat(path)
                except OSError:
                    continue
                old = self.entries.get(path)
                if old and old["mtime"] == st.st_mtime and old["size"] == st.st_size:
                    tf = old["tf"]
                else:
                    tf = collections.Counter(tokens(sorter.extract_text(path, max_pages=2)))
                    reread += 1
                entries[path] = {
                    "mtime": st.st_mtime, "size": st.st_size,
                    "label": label, "uses_year": uses_year, "tf": tf,
                }
        with self._lock:
            self.entries = entries
            self._rebuild()
            self.ready = True
        if reread or len(entries) != len(self._docs):
            try:
                self.save()
            except OSError:
                log.exception("Cache du modele non enregistre")
        return len(entries), reread

    def add(self, path, root):
        """Ajoute un PDF qui vient d'etre range (apprentissage au fil de l'eau)."""
        label, uses_year = self.label_for(root, os.path.dirname(path))
        if label is None:
            return
        st = os.stat(path)
        tf = collections.Counter(tokens(sorter.extract_text(path, max_pages=2)))
        with self._lock:
            self.entries[path] = {
                "mtime": st.st_mtime, "size": st.st_size,
                "label": label, "uses_year": uses_year, "tf": tf,
            }
            self._rebuild()
        try:
            self.save()
        except OSError:
            log.exception("Cache du modele non enregistre")

    def _rebuild(self):
        usable = [
            (p, e) for p, e in self.entries.items()
            if sum(e["tf"].values()) >= 5 and e["label"]
        ]
        n = len(usable)
        df = collections.Counter()
        for _, e in usable:
            df.update(e["tf"].keys())
        idf = {w: math.log((1 + n) / (1 + c)) + 1 for w, c in df.items()}

        postings = collections.defaultdict(list)
        docs, labels = [], []
        for i, (path, e) in enumerate(usable):
            weights = {w: (1 + math.log(c)) * idf[w] for w, c in e["tf"].items()}
            norm = math.sqrt(sum(v * v for v in weights.values())) or 1.0
            for w, v in weights.items():
                postings[w].append((i, v / norm))
            docs.append(path)
            labels.append(e["label"])
        self._idf, self._postings = idf, dict(postings)
        self._docs, self._labels = docs, labels

    # -- prediction --------------------------------------------------------

    def predict(self, text, exclude=None, k=5, min_sim=0.40, min_share=0.6):
        """Dossier propose pour ce texte.

        Retourne un dict {status, folder, uses_year, confidence, share,
        alternatives}. statut : 'matched', 'ambiguous' (les voisins se
        partagent entre plusieurs dossiers) ou 'unmatched' (rien de
        suffisamment ressemblant, ou modele vide)."""
        result = {"status": "unmatched", "folder": None, "uses_year": False,
                  "confidence": 0.0, "share": 0.0, "alternatives": []}
        with self._lock:
            tf = collections.Counter(tokens(text))
            if not self._docs or sum(tf.values()) < 5:
                return result
            q = {w: (1 + math.log(c)) * self._idf[w] for w, c in tf.items() if w in self._idf}
            norm = math.sqrt(sum(v * v for v in q.values()))
            if not norm:
                return result
            sims = collections.defaultdict(float)
            for w, qv in q.items():
                for i, dv in self._postings.get(w, ()):
                    sims[i] += qv / norm * dv
            ranked = sorted(
                ((s, i) for i, s in sims.items() if self._docs[i] != exclude),
                reverse=True,
            )[:k]
            if not ranked:
                return result

            votes = collections.defaultdict(float)
            for s, i in ranked:
                votes[self._labels[i]] += s
            total = sum(votes.values()) or 1.0
            best = max(votes, key=votes.get)
            year_votes = [
                self.entries[self._docs[i]]["uses_year"]
                for _, i in ranked if self._labels[i] == best
            ]
            result.update(
                folder=best,
                uses_year=sum(year_votes) * 2 > len(year_votes),
                confidence=ranked[0][0],
                share=votes[best] / total,
                alternatives=[l for l in sorted(votes, key=votes.get, reverse=True) if l != best],
            )
            if result["confidence"] < min_sim:
                result["status"] = "unmatched"
            elif result["share"] < min_share:
                result["status"] = "ambiguous"
            else:
                result["status"] = "matched"
            return result
