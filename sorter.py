"""
Tri automatique des PDF scannes : classement et renommage.

Deux moteurs, tout en local (rien n'est envoye a un service externe) :
  1. rules.json (facultatif) : regles a mots-cles ecrites a la main, qui
     l'emportent quand elles correspondent -- utile pour un cas precis.
  2. Un modele appris sur les PDF deja ranges (voir folder_model.py) : le
     dossier propose est celui des documents qui ressemblent le plus au
     nouveau.
Le nom du fichier devient AAAA-MM-JJ_<Libelle>.pdf, avec la date lue dans le
texte du document (ou, a defaut, la date du scan).

En ligne de commande (simulation, ne deplace rien) :

    python sorter.py --folder "<dossier A_TRIER>" --root "<dossier Documents>"
"""

import argparse
import json
import logging
import os
import re
import shutil
import sys
import unicodedata
from datetime import date, datetime, timedelta

log = logging.getLogger("autoscan.sorter")

MONTHS = {
    # francais
    "janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11,
    "decembre": 12,
    # allemand (hors doublons deja couverts)
    "januar": 1, "februar": 2, "marz": 3, "juni": 6, "juli": 7,
    "oktober": 10, "dezember": 12,
    # anglais
    "january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7,
    "august": 8, "october": 10, "december": 12,
}

_NUMERIC_DATE = re.compile(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{4}|\d{2})(?!\d)")
_ISO_DATE = re.compile(r"(?<!\d)(\d{4})-(\d{2})-(\d{2})(?!\d)")
_TEXT_DATE = re.compile(r"(?<!\d)(\d{1,2})(?:er)?\.?\s+([a-z]+)\s+(\d{4})(?!\d)")
_YEAR_DIR = re.compile(r"^(19|20)\d{2}$")
_FORBIDDEN_IN_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# ---------------------------------------------------------------------------
# Texte
# ---------------------------------------------------------------------------

def normalize(text):
    """Minuscules, sans accents, espaces compactes : pour comparer sans
    se soucier de la casse, des accents ou de la mise en page."""
    text = unicodedata.normalize("NFKD", text.lower())
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text)


def extract_text(pdf_path, max_pages=3):
    """Texte (couche OCR) des premieres pages du PDF, ou '' si illisible."""
    from pypdf import PdfReader
    try:
        reader = PdfReader(pdf_path)
        parts = []
        for page in reader.pages[:max_pages]:
            parts.append(page.extract_text() or "")
        return "\n".join(parts)
    except Exception:
        log.exception("Lecture du texte impossible pour %s", pdf_path)
        return ""


def find_date(text, today=None):
    """Premiere date plausible du texte (la date d'emission vient en
    general avant les echeances), jamais posterieure a `today` : un document
    scanne n'est pas date du futur. Retourne un `date` ou None."""
    today = today or date.today()
    norm = normalize(text)
    found = []

    def accept(y, m, d, pos):
        try:
            candidate = date(y, m, d)
        except ValueError:
            return
        if date(2000, 1, 1) <= candidate <= today + timedelta(days=1):
            found.append((pos, candidate))

    for m in _ISO_DATE.finditer(norm):
        accept(int(m[1]), int(m[2]), int(m[3]), m.start())
    for m in _NUMERIC_DATE.finditer(norm):
        y = int(m[3])
        if y < 100:
            y += 2000
        accept(y, int(m[2]), int(m[1]), m.start())
    for m in _TEXT_DATE.finditer(norm):
        month = MONTHS.get(m[2])
        if month:
            accept(int(m[3]), month, int(m[1]), m.start())

    return min(found)[1] if found else None


# ---------------------------------------------------------------------------
# Regles
# ---------------------------------------------------------------------------

def load_rules(path):
    """Charge rules.json ; chaque regle recoit ses mots-cles normalises."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    rules = []
    for raw in data["rules"]:
        rule = dict(raw)
        rule["_keywords"] = [normalize(k) for k in raw.get("keywords", [])]
        rule["_all"] = [normalize(k) for k in raw.get("all", [])]
        rule["_none"] = [normalize(k) for k in raw.get("none", [])]
        if not rule["_keywords"] and not rule["_all"]:
            continue
        rules.append(rule)
    return rules


def _has(norm_text, keyword):
    return re.search(r"(?<!\w)" + re.escape(keyword) + r"(?!\w)", norm_text) is not None


def _is_ancestor(parent, child):
    parent = parent.replace("\\", "/").rstrip("/").split("/{")[0]
    child = child.replace("\\", "/").rstrip("/").split("/{")[0]
    return child != parent and child.startswith(parent + "/")


def classify(text, rules):
    """Retourne (statut, regle, mots-cles trouves, alternatives).

    statut : 'matched', 'ambiguous' (plusieurs destinations plausibles) ou
    'unmatched'. Quand deux regles visent un dossier et l'un de ses
    sous-dossiers, la plus profonde l'emporte ; sinon la plus de mots-cles
    trouves, a egalite de poids, departage ; a egalite complete, on ne choisit
    pas (ambigu) plutot que de mal classer."""
    norm = normalize(text)
    scored = []
    for rule in rules:
        if any(_has(norm, k) for k in rule["_none"]):
            continue
        if rule["_all"] and not all(_has(norm, k) for k in rule["_all"]):
            continue
        hits = [k for k in rule["_keywords"] if _has(norm, k)]
        if rule["_keywords"] and len(hits) < rule.get("min_hits", 1):
            continue
        hits = rule["_all"] + hits
        score = (rule.get("weight", 0), len(hits), sum(len(k) for k in hits))
        scored.append((score, rule, hits))

    if not scored:
        return "unmatched", None, [], []

    folders = [r["folder"] for _, r, _ in scored]
    scored = [
        s for s in scored
        if not any(_is_ancestor(s[1]["folder"], other) for other in folders)
    ]
    scored.sort(key=lambda s: s[0], reverse=True)
    best = scored[0]
    others = [s for s in scored[1:] if s[1]["folder"] != best[1]["folder"]]
    if others and others[0][0] == best[0]:
        return "ambiguous", best[1], best[2], [s[1]["name"] for s in [best] + others]
    return "matched", best[1], best[2], [s[1]["name"] for s in others]


# ---------------------------------------------------------------------------
# Plan et application
# ---------------------------------------------------------------------------

def _safe_name(name):
    name = _FORBIDDEN_IN_NAME.sub("_", name).strip(" .")
    return name or "scan"


def plan_for(pdf_path, model, rules, sort_root, scanned_at, text=None):
    """Decide ou irait le PDF, sans rien deplacer.

    scanned_at (datetime) sert de date de repli quand le document n'en
    contient aucune de lisible."""
    if text is None:
        text = extract_text(pdf_path)
    doc_date = find_date(text, today=scanned_at.date())
    date_source = "texte" if doc_date else "scan"
    doc_date = doc_date or scanned_at.date()
    fields = {
        "date": doc_date.isoformat(),
        "year": f"{doc_date.year:04d}",
        "month": f"{doc_date.month:02d}",
    }

    plan = {
        "src": pdf_path, "status": "unmatched", "source": None, "rule": None,
        "hits": [], "alternatives": [], "confidence": None,
        "date": fields["date"], "date_source": date_source, "dst": None,
    }

    folder = filename = None
    status, rule, hits, alternatives = classify(text, rules) if rules else ("unmatched", None, [], [])
    if status == "matched":
        fields["label"] = rule["name"]
        folder = rule["folder"].format(**fields)
        filename = rule.get("filename", "{date}_{label}").format(**fields)
        plan.update(status="matched", source="regle", rule=rule["name"], hits=hits,
                    alternatives=alternatives)
    elif status == "ambiguous":
        plan.update(status="ambiguous", source="regle", rule=rule["name"], hits=hits,
                    alternatives=alternatives)
    elif model is not None:
        guess = model.predict(text, exclude=pdf_path)
        plan.update(status=guess["status"], source="appris",
                    confidence=round(guess["confidence"], 2),
                    alternatives=guess["alternatives"][:3])
        if guess["folder"]:
            plan["rule"] = guess["folder"]
        if guess["status"] == "matched":
            fields["label"] = guess["folder"].split("/")[-1]
            folder = guess["folder"] + ("/{year}" if guess["uses_year"] else "")
            folder = folder.format(**fields)
            filename = "{date}_{label}".format(**fields)

    if folder:
        filename = _safe_name(filename)
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        plan["dst"] = os.path.join(sort_root, *re.split(r"[\\/]", folder), filename)
    return plan


def _free_path(path):
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{stem}_{n}{ext}"):
        n += 1
    return f"{stem}_{n}{ext}"


def record(journal_path, entry):
    entry = dict(entry, time=datetime.now().isoformat(timespec="seconds"))
    with open(journal_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def apply_plan(plan, simulate, journal_path=None):
    """Execute (ou simule) un plan. Ne remplace jamais un fichier existant.
    Retourne le chemin final (ou celui qui serait utilise en simulation)."""
    src, dst = plan["src"], plan["dst"]
    label = f"{os.path.basename(src)}"
    if plan["status"] == "ambiguous":
        log.info("Tri %s : AMBIGU (%s ou %s) -> laisse en place",
                 label, plan["rule"], ", ".join(plan["alternatives"]) or "?")
        final = None
    elif plan["status"] == "unmatched" or not dst:
        log.info("Tri %s : aucune regle ne correspond -> laisse en place", label)
        final = None
    else:
        final = _free_path(dst)
        tag = "SIMULATION" if simulate else "deplace"
        log.info(
            "Tri %s [%s] : -> %s (%s %s%s, date %s)",
            label, tag, final, plan["source"], plan["rule"],
            f", confiance {plan['confidence']}" if plan["confidence"] is not None else "",
            plan["date_source"],
        )
        if not simulate:
            os.makedirs(os.path.dirname(final), exist_ok=True)
            shutil.move(src, final)

    if journal_path:
        record(journal_path, {
            "mode": "simulation" if simulate else "reel",
            "status": plan["status"],
            "src": src,
            "dst": final,
            "rule": plan["rule"],
            "source": plan["source"],
            "confidence": plan["confidence"],
            "alternatives": plan["alternatives"],
            "hits": plan["hits"],
            "date": plan["date"],
            "date_source": plan["date_source"],
        })
    return final


# ---------------------------------------------------------------------------
# Classement manuel (fenetre de tri) et liste des dossiers existants
# ---------------------------------------------------------------------------

def move_manual(src, folder, filename, root, journal_path=None, plan=None):
    """Range un PDF dans `folder` (relatif a root) sous le nom `filename`,
    a la demande de l'utilisateur. Cree le dossier au besoin, ne remplace
    jamais un fichier existant. Retourne le chemin final."""
    filename = _safe_name(filename)
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    parts = [p for p in re.split(r"[\\/]", folder) if p and p not in (".", "..")]
    dst = _free_path(os.path.join(root, *parts, filename))
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.move(src, dst)
    log.info("Tri manuel : %s -> %s", os.path.basename(src), dst)
    if journal_path:
        plan = plan or {}
        record(journal_path, {
            "mode": "reel", "status": "manuel", "src": src, "dst": dst,
            "rule": plan.get("rule"), "source": "manuel",
            "proposed_status": plan.get("status"),
            "confidence": plan.get("confidence"),
            "alternatives": plan.get("alternatives", []),
            "hits": [], "date": plan.get("date"), "date_source": plan.get("date_source"),
        })
    return dst


def list_folders(root, max_depth=4):
    """Dossiers existants sous root (chemins relatifs avec '/'), sans les
    dossiers d'annee ni ceux qui commencent par '.' ou '_'. Retourne un dict
    {chemin: True si le dossier est range par annee}."""
    folders = {}

    def walk(path, rel, depth):
        try:
            names = sorted(d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d)))
        except OSError:
            return
        has_years = any(_YEAR_DIR.match(d) for d in names)
        if rel:
            folders[rel] = has_years
        if depth >= max_depth:
            return
        for d in names:
            if _YEAR_DIR.match(d) or d.startswith((".", "_")):
                continue
            walk(os.path.join(path, d), f"{rel}/{d}" if rel else d, depth + 1)

    walk(root, "", 0)
    return folders


# ---------------------------------------------------------------------------
# CLI : propose un classement sans rien deplacer
# ---------------------------------------------------------------------------

def _main():
    from folder_model import FolderModel

    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--folder", required=True, help="dossier de PDF a analyser")
    ap.add_argument("--root", required=True, help="racine de classement (les dossiers deja ranges)")
    ap.add_argument("--rules", default=os.path.join(here, "rules.json"))
    ap.add_argument("--skip", nargs="*", default=["Avocat", "2FA"],
                    help="dossiers de premier niveau exclus de l'apprentissage")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    model = FolderModel(os.path.join(here, "sort_model.pkl"))
    model.load()
    total, reread = model.refresh(args.root, skip=args.skip)
    print(f"Modele : {total} PDF ranges ({reread} relus)")
    rules = load_rules(args.rules) if os.path.exists(args.rules) else []

    for name in sorted(os.listdir(args.folder)):
        if not name.lower().endswith(".pdf"):
            continue
        path = os.path.join(args.folder, name)
        plan = plan_for(path, model, rules, args.root, datetime.fromtimestamp(os.path.getmtime(path)))
        dest = os.path.relpath(plan["dst"], args.root) if plan["dst"] else "-"
        conf = f", confiance {plan['confidence']}" if plan["confidence"] is not None else ""
        alt = f", autres: {plan['alternatives']}" if plan["alternatives"] else ""
        print(f"{name}\n    [{plan['status']}] {dest}  ({plan['source']} {plan['rule']}{conf}{alt}, "
              f"date {plan['date']}/{plan['date_source']})")
    return 0


if __name__ == "__main__":
    sys.exit(_main())
