"""
Fenetre de tri des scans : liste les PDF en attente dans OUTPUT_FOLDER, montre
un apercu, propose un dossier et un nom (voir sorter.py / folder_model.py) et
range le PDF d'un clic. Chaque classement alimente ensuite l'apprentissage.

Lancee depuis le menu de l'icone (« Trier les scans... ») dans un processus a
part, ou directement :  pythonw sort_gui.py
"""

import ctypes
import logging
import os
import queue
import re
import shutil
import sys
import threading
import tkinter as tk
from datetime import date, datetime
from logging.handlers import RotatingFileHandler
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk
from pypdf import PdfReader

import config
import sorter
from folder_model import FolderModel

log = logging.getLogger("autoscan.gui")

try:  # apercu de tout PDF si pypdfium2 est installe (facultatif)
    import pypdfium2
except ImportError:
    pypdfium2 = None

POLL_MS = 3000          # surveillance du dossier de sortie
RESULT_POLL_MS = 150    # reception des analyses en arriere-plan
SETTLE_SECONDS = 3      # un PDF tout juste ecrit n'est pas encore propose
_YEAR_DIR = re.compile(r"^(19|20)\d{2}$")


# ---------------------------------------------------------------------------
# Apercu
# ---------------------------------------------------------------------------

def page_count(path):
    try:
        return max(1, len(PdfReader(path).pages))
    except Exception:
        return 1


def render_page(path, index):
    """Image PIL de la page `index`, ou None si l'apercu est impossible.
    pypdfium2 si present ; sinon on prend l'image integree a la page, ce qui
    suffit pour les PDF d'Autoscan (une image scannee par page)."""
    if pypdfium2 is not None:
        try:
            pdf = pypdfium2.PdfDocument(path)
            try:
                return pdf[index].render(scale=1.6).to_pil().convert("RGB")
            finally:
                pdf.close()
        except Exception:
            log.exception("pypdfium2 : apercu impossible pour %s", path)
    try:
        page = PdfReader(path).pages[index]
        images = list(page.images)
        if images:
            return images[0].image.convert("RGB")
    except Exception:
        log.exception("Apercu impossible pour %s", path)
    return None


def move_to_recycle_bin(path):
    """Envoie le fichier a la Corbeille Windows (recuperable). Leve OSError en
    cas d'echec. Pas de suppression definitive."""
    from win32com.shell import shell, shellcon
    flags = (shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION
             | shellcon.FOF_NOERRORUI | shellcon.FOF_SILENT)
    result, aborted = shell.SHFileOperation(
        (0, shellcon.FO_DELETE, os.path.abspath(path), None, flags, None, None))
    if result != 0 or aborted:
        raise OSError(f"La Corbeille a refuse le fichier (code {result})")
    if os.path.exists(path):
        raise OSError("Le fichier est toujours present apres l'envoi a la Corbeille")


# ---------------------------------------------------------------------------
# Analyse en arriere-plan (modele + proposition pour chaque PDF)
# ---------------------------------------------------------------------------

class Analyzer(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)
        self.todo = queue.Queue()
        self.results = queue.Queue()
        self.model = FolderModel(config.SORT_MODEL_FILE)
        self.status = "Chargement du modele..."

    def run(self):
        try:
            self.model.load()
            total, reread = self.model.refresh(
                config.SORT_ROOT, skip=config.SORT_SKIP, exclude=[config.OUTPUT_FOLDER])
            self.status = f"Modele : {total} PDF ranges"
        except Exception:
            log.exception("Modele de tri indisponible")
            self.status = "Modele indisponible (propositions desactivees)"
        self.results.put(("status", self.status))
        while True:
            path = self.todo.get()
            try:
                rules = sorter.load_rules(config.RULES_FILE) if os.path.exists(config.RULES_FILE) else []
                scanned_at = datetime.fromtimestamp(os.path.getmtime(path))
                plan = sorter.plan_for(path, self.model, rules, config.SORT_ROOT, scanned_at)
            except Exception:
                log.exception("Analyse impossible pour %s", path)
                plan = None
            self.results.put(("plan", path, plan))


# ---------------------------------------------------------------------------
# Fenetre
# ---------------------------------------------------------------------------

class SortApp:
    def __init__(self, root):
        self.root = root
        root.title("Autoscan - Trier les scans")
        root.geometry("1180x760")
        root.minsize(900, 600)

        self.analyzer = Analyzer()
        self.files = []              # chemins, du plus recent au plus ancien
        self.plans = {}              # chemin -> plan (ou None si en cours d'analyse)
        self.current = None
        self.undo_stack = []         # (chemin final, chemin d'origine)
        self.folders = {}            # dossier relatif -> range par annee ?
        self.page = 0
        self.pages = 1
        self.preview_cache = {}
        self.photo = None
        self.name_dirty = False
        self._resize_job = None

        self._build()
        self._refresh_folders()
        self.analyzer.start()
        self._poll_results()
        self._poll_folder()

    # -- construction ------------------------------------------------------

    def _build(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=26)
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        # Colonne de gauche : liste des scans en attente
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        self.count_label = ttk.Label(left, text="", font=("Segoe UI", 10, "bold"))
        self.count_label.pack(anchor="w", pady=(0, 4))
        columns = ("name", "dest", "state")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="browse")
        for col, title, width in (("name", "Fichier", 170), ("dest", "Proposition", 190), ("state", "", 70)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, anchor="w", stretch=(col != "state"))
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Delete>", lambda e: self._delete_current())

        # Colonne de droite : apercu + formulaire
        right = ttk.Frame(paned)
        paned.add(right, weight=3)

        nav = ttk.Frame(right)
        nav.pack(fill="x")
        self.prev_btn = ttk.Button(nav, text="<", width=3, command=lambda: self._turn_page(-1))
        self.prev_btn.pack(side="left")
        self.page_label = ttk.Label(nav, text="", width=10, anchor="center")
        self.page_label.pack(side="left", padx=4)
        self.next_btn = ttk.Button(nav, text=">", width=3, command=lambda: self._turn_page(1))
        self.next_btn.pack(side="left")
        ttk.Label(nav, text="Double-clic sur l'apercu : ouvrir le PDF", foreground="gray").pack(side="right")

        self.preview = tk.Label(right, background="#e8e8e8", anchor="center")
        self.preview.pack(fill="both", expand=True, pady=6)
        self.preview.bind("<Configure>", self._on_preview_resize)
        self.preview.bind("<Double-Button-1>", lambda e: self._open_pdf())

        form = ttk.Frame(right)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Suggestions").grid(row=0, column=0, sticky="nw", pady=2)
        self.chips = ttk.Frame(form)
        self.chips.grid(row=0, column=1, columnspan=3, sticky="ew", pady=2)

        ttk.Label(form, text="Dossier").grid(row=1, column=0, sticky="w", pady=2)
        self.folder_var = tk.StringVar()
        self.folder_box = ttk.Combobox(form, textvariable=self.folder_var)
        self.folder_box.grid(row=1, column=1, sticky="ew", pady=2)
        self.folder_box.bind("<KeyRelease>", self._filter_folders)
        self.folder_box.bind("<<ComboboxSelected>>", lambda e: self._on_folder_changed())
        ttk.Button(form, text="Parcourir...", command=self._browse).grid(row=1, column=2, padx=(6, 0))
        self.folder_var.trace_add("write", lambda *a: self._on_folder_changed())

        self.year_var = tk.BooleanVar()
        ttk.Checkbutton(form, text="Sous-dossier par annee", variable=self.year_var,
                        command=self._update_target).grid(row=2, column=1, sticky="w")
        datebox = ttk.Frame(form)
        datebox.grid(row=2, column=2, columnspan=2, sticky="e")
        ttk.Label(datebox, text="Date").pack(side="left", padx=(0, 4))
        self.date_var = tk.StringVar()
        date_entry = ttk.Entry(datebox, textvariable=self.date_var, width=11)
        date_entry.pack(side="left")
        self.date_var.trace_add("write", lambda *a: self._on_date_changed())

        ttk.Label(form, text="Nom").grid(row=3, column=0, sticky="w", pady=2)
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(form, textvariable=self.name_var)
        name_entry.grid(row=3, column=1, sticky="ew", pady=2)
        name_entry.bind("<Key>", lambda e: setattr(self, "name_dirty", True))
        self.name_var.trace_add("write", lambda *a: self._update_target())
        ttk.Label(form, text=".pdf").grid(row=3, column=2, sticky="w", padx=(4, 0))

        self.target_label = ttk.Label(form, text="", foreground="#1a5fb4", wraplength=700, justify="left")
        self.target_label.grid(row=4, column=0, columnspan=4, sticky="w", pady=(4, 6))

        buttons = ttk.Frame(form)
        buttons.grid(row=5, column=0, columnspan=4, sticky="ew")
        self.sort_btn = ttk.Button(buttons, text="Classer", style="Accent.TButton", command=self._sort_current)
        self.sort_btn.pack(side="left")
        ttk.Button(buttons, text="Passer", command=self._skip).pack(side="left", padx=6)
        ttk.Button(buttons, text="Ouvrir le PDF", command=self._open_pdf).pack(side="left")
        self.delete_btn = ttk.Button(buttons, text="Supprimer (Corbeille)", command=self._delete_current)
        self.delete_btn.pack(side="left", padx=(18, 0))
        self.undo_btn = ttk.Button(buttons, text="Annuler le dernier classement", command=self._undo)
        self.undo_btn.pack(side="right")

        self.status = ttk.Label(self.root, text="", foreground="gray")
        self.status.pack(fill="x", padx=10, pady=(0, 6))

        name_entry.bind("<Return>", lambda e: self._sort_current())
        self.folder_box.bind("<Return>", lambda e: self._sort_current())
        date_entry.bind("<Return>", lambda e: self._sort_current())
        self._set_form_enabled(False)

    # -- liste des scans ---------------------------------------------------

    def _pending_pdfs(self):
        found = []
        try:
            for name in os.listdir(config.OUTPUT_FOLDER):
                path = os.path.join(config.OUTPUT_FOLDER, name)
                if name.lower().endswith(".pdf") and os.path.isfile(path):
                    mtime = os.path.getmtime(path)
                    if datetime.now().timestamp() - mtime >= SETTLE_SECONDS:
                        found.append((mtime, path))
        except OSError:
            pass
        return [p for _, p in sorted(found, reverse=True)]

    def _poll_folder(self):
        present = self._pending_pdfs()
        changed = False
        for path in present:
            if path not in self.plans:
                self.plans[path] = None
                self.files.append(path)
                self.analyzer.todo.put(path)
                changed = True
        for path in list(self.files):
            if path not in present:
                self.files.remove(path)
                self.plans.pop(path, None)
                changed = True
        if changed:
            self.files = [p for p in present if p in self.plans]
            self._reload_tree()
        self.root.after(POLL_MS, self._poll_folder)

    def _poll_results(self):
        try:
            while True:
                msg = self.analyzer.results.get_nowait()
                if msg[0] == "status":
                    self._refresh_folders()
                    self._set_status(msg[1])
                else:
                    _, path, plan = msg
                    if path in self.plans:
                        self.plans[path] = plan or {"status": "unmatched", "alternatives": [],
                                                    "rule": None, "dst": None,
                                                    "date": date.today().isoformat(), "confidence": None}
                        self._update_row(path)
                        if path == self.current:
                            self._load_form(path)
        except queue.Empty:
            pass
        self.root.after(RESULT_POLL_MS, self._poll_results)

    def _row_values(self, path):
        plan = self.plans.get(path)
        name = os.path.basename(path)
        if plan is None:
            return (name, "analyse...", "")
        folder = self._plan_folder(plan)
        label = {"matched": "propose", "ambiguous": "ambigu"}.get(plan["status"], "inconnu")
        return (name, folder or "-", label)

    def _reload_tree(self):
        keep = self.current
        self.tree.delete(*self.tree.get_children())
        for path in self.files:
            self.tree.insert("", "end", iid=path, values=self._row_values(path))
        self.count_label.config(text=f"{len(self.files)} scan(s) a trier")
        if keep in self.files:
            self.tree.selection_set(keep)
        elif self.files:
            self.tree.selection_set(self.files[0])
        else:
            self.current = None
            self._clear_form()

    def _update_row(self, path):
        if self.tree.exists(path):
            self.tree.item(path, values=self._row_values(path))

    def _on_select(self, _event=None):
        selection = self.tree.selection()
        if not selection or selection[0] == self.current:
            return
        self.current = selection[0]
        self.page = 0
        self.pages = page_count(self.current)
        self._load_form(self.current)
        self._show_preview()

    def _skip(self):
        if not self.files or self.current not in self.files:
            return
        i = self.files.index(self.current)
        self.tree.selection_set(self.files[(i + 1) % len(self.files)])

    # -- formulaire --------------------------------------------------------

    def _refresh_folders(self):
        self.folders = sorter.list_folders(config.SORT_ROOT)
        self.all_folders = sorted(self.folders, key=str.lower)
        self.folder_box["values"] = self.all_folders

    def _filter_folders(self, event):
        if event.keysym in ("Return", "Tab", "Up", "Down", "Escape"):
            return
        query = sorter.normalize(self.folder_var.get())
        self.folder_box["values"] = [
            f for f in self.all_folders if query in sorter.normalize(f)] or self.all_folders

    def _plan_folder(self, plan):
        """Dossier propose (relatif, sans annee) ou None."""
        if plan.get("dst"):
            rel = os.path.relpath(os.path.dirname(plan["dst"]), config.SORT_ROOT).replace("\\", "/")
            parts = rel.split("/")
            if len(parts) > 1 and _YEAR_DIR.match(parts[-1]):
                parts = parts[:-1]
            return "/".join(parts)
        return None

    def _suggestions(self, plan):
        suggestions = []
        first = self._plan_folder(plan)
        if first:
            suggestions.append((first, plan.get("confidence")))
        elif plan.get("source") == "appris" and plan.get("rule"):
            suggestions.append((plan["rule"], plan.get("confidence")))
        if plan.get("source") == "appris":
            suggestions += [(alt, None) for alt in plan.get("alternatives", [])]
        seen, result = set(), []
        for folder, conf in suggestions:
            if folder not in seen:
                seen.add(folder)
                result.append((folder, conf))
        return result[:4]

    def _set_form_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.delete_btn.configure(state="normal" if self.current else "disabled")
        for widget in (self.folder_box, self.sort_btn):
            widget.configure(state=state)

    def _clear_form(self):
        self._set_chips([])
        self.folder_var.set("")
        self.name_var.set("")
        self.date_var.set("")
        self.target_label.config(text="")
        self.preview.configure(image="", text="Aucun scan a trier", font=("Segoe UI", 12))
        self.page_label.config(text="")
        self._set_form_enabled(False)

    def _set_chips(self, suggestions):
        for child in self.chips.winfo_children():
            child.destroy()
        if not suggestions:
            ttk.Label(self.chips, text="aucune", foreground="gray").pack(side="left")
        for folder, conf in suggestions:
            text = folder + (f"  ({conf:.2f})" if conf is not None else "")
            ttk.Button(self.chips, text=text, command=lambda f=folder: self._pick(f)).pack(
                side="left", padx=(0, 6))

    def _pick(self, folder):
        self.folder_var.set(folder)
        self.name_dirty = False
        self._on_folder_changed()

    def _load_form(self, path):
        plan = self.plans.get(path)
        self._set_form_enabled(plan is not None)
        self.name_dirty = False
        if plan is None:
            self._set_chips([])
            self.folder_var.set("")
            self.date_var.set(date.today().isoformat())
            self.target_label.config(text="Analyse en cours...")
            return
        self.date_var.set(plan.get("date") or date.today().isoformat())
        self._set_chips(self._suggestions(plan))
        folder = self._plan_folder(plan) if plan["status"] == "matched" else ""
        self.folder_var.set(folder or "")
        if folder:
            self.year_var.set(bool(plan["dst"]) and bool(
                _YEAR_DIR.match(os.path.basename(os.path.dirname(plan["dst"])))))
        self._apply_default_name()
        self._update_target()

    def _default_name(self):
        folder = self.folder_var.get().strip().strip("/\\")
        leaf = folder.split("/")[-1] if folder else ""
        if not leaf and self.current:
            leaf = os.path.splitext(os.path.basename(self.current))[0]
        return f"{self.date_var.get().strip()}_{leaf}".strip("_")

    def _apply_default_name(self):
        if not self.name_dirty:
            self.name_var.set(self._default_name())

    def _on_folder_changed(self):
        folder = self.folder_var.get().strip().strip("/\\").replace("\\", "/")
        if folder in self.folders:
            self.year_var.set(self.folders[folder])
        else:
            self.year_var.set(False)  # nouveau dossier : a cocher a la main si voulu
        self._apply_default_name()
        self._update_target()

    def _on_date_changed(self):
        self._apply_default_name()
        self._update_target()

    def _browse(self):
        start = os.path.join(config.SORT_ROOT, *self.folder_var.get().split("/"))
        chosen = filedialog.askdirectory(
            parent=self.root, title="Dossier de destination",
            initialdir=start if os.path.isdir(start) else config.SORT_ROOT)
        if not chosen:
            return
        rel = os.path.relpath(chosen, config.SORT_ROOT).replace("\\", "/")
        if rel.startswith(".."):
            messagebox.showwarning("Dossier", "Choisis un dossier a l'interieur de :\n" + config.SORT_ROOT)
            return
        parts = rel.split("/")
        if len(parts) > 1 and _YEAR_DIR.match(parts[-1]):
            parts = parts[:-1]
        self.folder_var.set("/".join(parts))

    def _target(self):
        """(dossier relatif, nom de fichier, erreur) pour l'etat actuel du formulaire."""
        folder = self.folder_var.get().strip().strip("/\\").replace("\\", "/")
        name = self.name_var.get().strip()
        if not folder:
            return None, None, "Choisis un dossier."
        if not name:
            return None, None, "Donne un nom au fichier."
        iso = self.date_var.get().strip()
        year = None
        if self.year_var.get():
            try:
                year = datetime.strptime(iso, "%Y-%m-%d").year
            except ValueError:
                return None, None, "Date invalide (format AAAA-MM-JJ) : necessaire pour le dossier d'annee."
            folder = f"{folder}/{year:04d}"
        return folder, name, None

    def _update_target(self):
        if self.current is None:
            return
        folder, name, error = self._target()
        if error:
            self.target_label.config(text=error, foreground="#b00020")
            self.sort_btn.configure(state="disabled")
            return
        exists = os.path.isdir(os.path.join(config.SORT_ROOT, *folder.split("/")))
        shown = f"{folder}/{name}.pdf" + ("" if exists else "   (nouveau dossier)")
        self.target_label.config(text="Destination : " + shown, foreground="#1a5fb4")
        self.sort_btn.configure(state="normal")

    # -- actions -----------------------------------------------------------

    def _sort_current(self):
        path = self.current
        if path is None or self.plans.get(path) is None:
            return
        folder, name, error = self._target()
        if error:
            messagebox.showwarning("Classement", error)
            return
        try:
            final = sorter.move_manual(
                path, folder, name, config.SORT_ROOT, config.SORT_JOURNAL, self.plans.get(path))
        except OSError as e:
            messagebox.showerror("Classement impossible", f"{e}\n\nLe PDF est peut-etre ouvert dans un autre programme.")
            return
        self.undo_stack.append((final, path))
        top = folder.split("/")[0]
        if sorter.normalize(top) not in {sorter.normalize(s) for s in config.SORT_SKIP}:
            threading.Thread(target=self._learn, args=(final,), daemon=True).start()
        self._set_status(f"Classe : {os.path.relpath(final, config.SORT_ROOT)}")

        self._remove_from_list(path)

    def _remove_from_list(self, path):
        """Retire un PDF traite de la liste et selectionne le suivant."""
        i = self.files.index(path)
        self.files.remove(path)
        self.plans.pop(path, None)
        self.preview_cache = {k: v for k, v in self.preview_cache.items() if k[0] != path}
        self.current = None
        self._refresh_folders()
        self.tree.delete(path)
        self.count_label.config(text=f"{len(self.files)} scan(s) a trier")
        if self.files:
            self.tree.selection_set(self.files[min(i, len(self.files) - 1)])
        else:
            self._clear_form()

    def _delete_current(self):
        path = self.current
        if path is None or path not in self.files:
            return
        name = os.path.basename(path)
        if not messagebox.askyesno(
                "Supprimer ce scan ?",
                f"Envoyer a la Corbeille :\n\n{name}\n\n(Il reste recuperable depuis la Corbeille Windows.)",
                icon="warning", default="no", parent=self.root):
            return
        try:
            move_to_recycle_bin(path)
        except Exception as e:
            log.exception("Suppression impossible : %s", path)
            messagebox.showerror("Suppression impossible",
                                 f"{e}\n\nLe PDF est peut-etre ouvert dans un autre programme.")
            return
        log.info("Scan envoye a la Corbeille : %s", path)
        sorter.record(config.SORT_JOURNAL, {"mode": "reel", "status": "supprime", "src": path, "dst": None})
        self._set_status(f"Envoye a la Corbeille : {name}")
        self._remove_from_list(path)

    def _learn(self, final):
        try:
            self.analyzer.model.add(final, config.SORT_ROOT)
        except Exception:
            log.exception("Apprentissage impossible pour %s", final)

    def _undo(self):
        if not self.undo_stack:
            self._set_status("Rien a annuler.")
            return
        final, original = self.undo_stack.pop()
        if not os.path.exists(final):
            self._set_status("Impossible d'annuler : le fichier n'est plus la.")
            return
        back = original
        if os.path.exists(back):
            stem, ext = os.path.splitext(original)
            back = f"{stem}_restaure{ext}"
        try:
            shutil.move(final, back)
        except OSError as e:
            messagebox.showerror("Annulation impossible", str(e))
            return
        log.info("Classement annule : %s -> %s", final, back)
        sorter.record(config.SORT_JOURNAL, {"mode": "reel", "status": "annule", "src": final, "dst": back})
        self._set_status(f"Annule : {os.path.basename(back)} est de nouveau a trier")
        # le modele oubliera l'ancien emplacement au prochain rafraichissement

    def _open_pdf(self):
        if self.current and os.path.exists(self.current):
            os.startfile(self.current)

    def _set_status(self, text):
        self.status.config(text=text)

    # -- apercu ------------------------------------------------------------

    def _turn_page(self, step):
        if self.current is None:
            return
        self.page = max(0, min(self.pages - 1, self.page + step))
        self._show_preview()

    def _on_preview_resize(self, _event):
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(120, self._show_preview)

    def _show_preview(self):
        self._resize_job = None
        path = self.current
        if path is None:
            return
        self.page_label.config(text=f"{self.page + 1} / {self.pages}")
        self.prev_btn.configure(state="normal" if self.page > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.page < self.pages - 1 else "disabled")

        key = (path, self.page)
        if key not in self.preview_cache:
            if len(self.preview_cache) > 6:
                self.preview_cache.pop(next(iter(self.preview_cache)))
            self.preview_cache[key] = render_page(path, self.page)
        image = self.preview_cache[key]
        if image is None:
            self.photo = None
            self.preview.configure(
                image="", font=("Segoe UI", 11),
                text="Apercu indisponible pour ce PDF.\n(pip install pypdfium2 pour tout afficher)\n"
                     "Double-clic pour l'ouvrir.")
            return
        box_w = max(200, self.preview.winfo_width() - 8)
        box_h = max(200, self.preview.winfo_height() - 8)
        scale = min(box_w / image.width, box_h / image.height, 1.0)
        size = (max(1, int(image.width * scale)), max(1, int(image.height * scale)))
        self.photo = ImageTk.PhotoImage(image.resize(size, Image.LANCZOS))
        self.preview.configure(image=self.photo, text="")


# ---------------------------------------------------------------------------
# Lancement (une seule fenetre a la fois)
# ---------------------------------------------------------------------------

def _already_running():
    try:
        import win32api
        import win32event
        import winerror
        mutex = win32event.CreateMutex(None, False, "Autoscan.SortGui")
        _already_running.handle = mutex  # garde le mutex vivant
        return win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS
    except Exception:
        return False


def main():
    handler = RotatingFileHandler(
        os.path.join(config.SCRIPT_DIR, "sort_gui.log"), maxBytes=500_000, backupCount=2, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logging.basicConfig(level=logging.INFO, handlers=[handler])

    if _already_running():
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # texte net sur ecran HiDPI
    except Exception:
        pass
    root = tk.Tk()
    SortApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
