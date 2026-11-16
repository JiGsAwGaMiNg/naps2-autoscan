"""Reglages partages entre l'application (autoscan.py) et la fenetre de tri
(sort_gui.py), lus depuis .env. Ne depend d'aucun module lourd : on peut
l'importer sans lancer le scan ni creer de log."""

import os

from dotenv import load_dotenv

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))

OUTPUT_FOLDER = os.environ["OUTPUT_FOLDER"]

# Tri automatique (voir sorter.py). SORT_MODE=simulation : on journalise ce qui
# serait fait sans rien deplacer ; SORT_MODE=active : les PDF sont deplaces.
AUTO_SORT = os.environ.get("AUTO_SORT", "true").lower() != "false"
SORT_MODE = os.environ.get("SORT_MODE", "simulation").lower()
SORT_ROOT = os.environ.get("SORT_ROOT") or os.path.dirname(os.path.abspath(OUTPUT_FOLDER))
SORT_SKIP = [s.strip() for s in os.environ.get("SORT_SKIP", "Avocat,2FA").split(",") if s.strip()]

# Le log reste a cote du script, pas dans OUTPUT_FOLDER : ce dossier est
# souvent synchronise (Nextcloud...) et un log qui grossit y serait renvoye
# en permanence.
LOG_FILE = os.environ.get("LOG_FILE") or os.path.join(SCRIPT_DIR, "autoscan.log")
STATE_FILE = os.path.join(SCRIPT_DIR, "autoscan_state.json")
RULES_FILE = os.path.join(SCRIPT_DIR, "rules.json")
SORT_JOURNAL = os.environ.get("SORT_JOURNAL") or os.path.join(SCRIPT_DIR, "sort_journal.jsonl")
SORT_MODEL_FILE = os.environ.get("SORT_MODEL_FILE") or os.path.join(SCRIPT_DIR, "sort_model.pkl")
