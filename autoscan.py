"""
Scan automatique multi-pages pour Brother DS-640, via NAPS2 (moteur WIA natif).

Tourne en continu et totalement invisible (icone dans la barre des taches
seulement -- lance-le avec pythonw.exe, pas python.exe). Boucle sur des
tentatives de scan :
  - Des qu'une feuille est presente dans le bac, elle est scannee (en image
    brute) et ajoutee au document en cours.
  - S'il n'y a pas de feuille, l'appel echoue rapidement (~2s) et on
    reessaie.
  - Si aucune nouvelle page n'arrive pendant WAIT_SECONDS apres la
    derniere page scannee, chaque page est ré-orientee automatiquement si
    besoin (detection via Tesseract OSD), puis toutes les pages sont
    fusionnees et passees a l'OCR en un seul PDF final dans OUTPUT_FOLDER.

Couleurs de l'icone :
  bleu = en attente d'un document
  vert = document en cours de constitution (des pages sont deja scannees)
  rouge = erreur (jam, device deconnecte, etc.)

L'appel WIA direct via pywin32 (couche "Automation") s'est avere
incompatible avec le driver Brother DS-640 sur cette machine (echec
systematique et lent de Transfer(), ~10-27s). NAPS2 utilise l'API WIA 2.0
native (la meme que "Telecopie et numerisation Windows", qui fonctionne),
et echoue rapidement (~2s) quand le bac est vide.
"""

import logging
import os
import subprocess
import threading
import time
import uuid
from datetime import datetime

import pystray
import pytesseract
import win32job
from dotenv import load_dotenv
from PIL import Image, ImageDraw
from pypdf import PdfWriter

# ---------------------------------------------------------------------------
# Configuration (lue depuis .env -- voir .env.example)
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))

OUTPUT_FOLDER = os.environ["OUTPUT_FOLDER"]
WAIT_SECONDS = int(os.environ.get("WAIT_SECONDS", "15"))
POLL_INTERVAL = int(os.environ.get("POLL_INTERVAL", "2"))

NAPS2_CONSOLE = os.environ.get(
    "NAPS2_CONSOLE", os.path.join(SCRIPT_DIR, "NAPS2", "App", "NAPS2.Console.exe")
)
DEVICE_NAME = os.environ["DEVICE_NAME"]
DPI = int(os.environ.get("DPI", "200"))
COLOR_MODE = os.environ.get("COLOR_MODE", "color")  # "color", "gray" ou "bw"
PAGE_SIZE = os.environ.get("PAGE_SIZE", "a4")
DESKEW = os.environ.get("DESKEW", "true").lower() != "false"
AUTO_ROTATE = os.environ.get("AUTO_ROTATE", "true").lower() != "false"
OCR_LANG = os.environ.get("OCR_LANG", "fra+eng")
TESSERACT_CMD = os.environ.get("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")

SESSION_DIR = os.path.join(OUTPUT_FOLDER, ".autoscan_tmp")
LOG_FILE = os.path.join(OUTPUT_FOLDER, "autoscan.log")

EMPTY_FEEDER_MARKERS = (
    "aucune page dans le dispositif",  # FR
    "no pages were found in the document feeder",  # EN
)

pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("autoscan")


class ScanFailed(Exception):
    pass


# ---------------------------------------------------------------------------
# Job Object Windows : garantit que NAPS2.Console.exe est tue automatiquement
# si ce script est arrete brutalement (fermeture de fenetre, Ctrl+C, kill),
# pour eviter un sous-processus orphelin qui garde le scanner "occupe".
# ---------------------------------------------------------------------------

_job_handle = win32job.CreateJobObject(None, "")
_job_info = win32job.QueryInformationJobObject(_job_handle, win32job.JobObjectExtendedLimitInformation)
_job_info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
win32job.SetInformationJobObject(_job_handle, win32job.JobObjectExtendedLimitInformation, _job_info)


def _run_naps2(args, timeout=60):
    proc = subprocess.Popen(
        [NAPS2_CONSOLE] + args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    try:
        win32job.AssignProcessToJobObject(_job_handle, proc._handle)
    except Exception:
        log.warning("Impossible de rattacher NAPS2.Console.exe au job object")

    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.communicate()
        raise ScanFailed(f"NAPS2 n'a pas repondu dans les {timeout}s, processus tue")

    return proc.returncode, stdout, stderr


# ---------------------------------------------------------------------------
# Icone barre des taches
# ---------------------------------------------------------------------------

def make_icon_image(color):
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((8, 8, 56, 56), fill=color)
    return img


ICONS = {
    "idle": make_icon_image((70, 130, 180, 255)),
    "collecting": make_icon_image((60, 170, 80, 255)),
    "error": make_icon_image((220, 20, 60, 255)),
}

tray_icon = None


def set_status(state, tooltip):
    if tray_icon is not None:
        tray_icon.icon = ICONS[state]
        tray_icon.title = f"Autoscan DS-640 - {tooltip}"


# ---------------------------------------------------------------------------
# Scan via NAPS2 (image brute, pas encore d'OCR ni de rotation)
# ---------------------------------------------------------------------------

def scan_page_to_image(output_path):
    """Retourne True si une page a ete scannee, False si le bac est vide."""
    args = [
        "--driver", "wia",
        "--device", DEVICE_NAME,
        "--source", "feeder",
        "--dpi", str(DPI),
        "--bitdepth", COLOR_MODE,
        "--pagesize", PAGE_SIZE,
        "-o", output_path,
        "--force",
    ]
    if DESKEW:
        args.append("--deskew")

    returncode, stdout, stderr = _run_naps2(args, timeout=60)

    if returncode == 0 and os.path.exists(output_path):
        return True

    combined_output = f"{stdout} {stderr}".lower()
    if any(marker in combined_output for marker in EMPTY_FEEDER_MARKERS):
        return False

    raise ScanFailed(
        f"NAPS2 a echoue (code {returncode}) : {stdout.strip()} {stderr.strip()}"
    )


# ---------------------------------------------------------------------------
# Correction d'orientation (page inseree a l'envers ou de travers a 90/180/270)
# ---------------------------------------------------------------------------

def correct_orientation(image_path):
    try:
        with Image.open(image_path) as img:
            osd = pytesseract.image_to_osd(img, output_type=pytesseract.Output.DICT)
            rotate = int(osd.get("rotate", 0))
            if rotate:
                corrected = img.rotate(-rotate, expand=True)
                corrected.save(image_path)
                log.info(
                    "Page %s : rotation corrigee de %d degres (confiance %.1f)",
                    os.path.basename(image_path), rotate, osd.get("orientation_conf", 0),
                )
    except pytesseract.TesseractError as e:
        log.warning(
            "Detection d'orientation impossible pour %s (page probablement peu de texte) : %s",
            os.path.basename(image_path), e,
        )
    except Exception:
        log.exception("Erreur inattendue lors de la correction d'orientation de %s", image_path)


# ---------------------------------------------------------------------------
# Finalisation : rotation + fusion + OCR en un seul PDF, via NAPS2
# ---------------------------------------------------------------------------

def finalize_pdf(page_paths, created_at):
    if AUTO_ROTATE:
        for p in page_paths:
            correct_orientation(p)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    date_string = datetime.fromtimestamp(created_at).strftime("%Y-%m-%d_%H-%M-%S")
    output_path = os.path.join(OUTPUT_FOLDER, f"scan_{date_string}.pdf")

    import_arg = ";".join(page_paths)
    args = [
        "-i", import_arg,
        "-n", "0",
        "--ocrlang", OCR_LANG,
        "--enableocr",
        "-o", output_path,
        "--force",
    ]
    returncode, stdout, stderr = _run_naps2(args, timeout=60 + 20 * len(page_paths))

    if returncode != 0 or not os.path.exists(output_path):
        log.error(
            "Echec de la fusion/OCR (code %s) : %s %s -- fallback sans OCR",
            returncode, stdout.strip(), stderr.strip(),
        )
        writer = PdfWriter()
        for p in page_paths:
            writer.append(p)
        with open(output_path, "wb") as f:
            writer.write(f)
        writer.close()

    for p in page_paths:
        os.remove(p)

    log.info("PDF finalise : %s (%d page(s))", output_path, len(page_paths))


# ---------------------------------------------------------------------------
# Boucle de scan (tourne dans un thread a part, l'icone occupe le thread principal)
# ---------------------------------------------------------------------------

def scan_loop(stop_event):
    os.makedirs(SESSION_DIR, exist_ok=True)
    log.info("Autoscan demarre, en attente de documents")
    set_status("idle", "en attente d'un document")

    session_pages = []
    session_created_at = None
    last_activity = None

    while not stop_event.is_set():
        tmp_path = os.path.join(SESSION_DIR, f"page_{uuid.uuid4().hex}.png")

        try:
            scanned = scan_page_to_image(tmp_path)
        except ScanFailed as e:
            log.error("Echec du scan : %s", e)
            set_status("error", str(e)[:120])
            stop_event.wait(POLL_INTERVAL)
            continue

        if scanned:
            if not session_pages:
                session_created_at = time.time()
                log.info("Nouveau document, premiere page scannee")
            session_pages.append(tmp_path)
            last_activity = time.time()
            log.info("Page %d scannee", len(session_pages))
            set_status("collecting", f"{len(session_pages)} page(s) - en attente de la suite")
        else:
            if session_pages and (time.time() - last_activity) > WAIT_SECONDS:
                finalize_pdf(session_pages, session_created_at)
                session_pages = []
                session_created_at = None
                last_activity = None
                set_status("idle", "en attente d'un document")
            stop_event.wait(POLL_INTERVAL)


# ---------------------------------------------------------------------------
# Menu de l'icone
# ---------------------------------------------------------------------------

def open_output_folder(icon, item):
    os.startfile(OUTPUT_FOLDER)


def quit_app(icon, item, stop_event):
    stop_event.set()
    icon.stop()


def main():
    global tray_icon

    stop_event = threading.Event()
    worker = threading.Thread(target=scan_loop, args=(stop_event,), daemon=True)

    menu = pystray.Menu(
        pystray.MenuItem("Ouvrir le dossier de sortie", open_output_folder),
        pystray.MenuItem("Quitter", lambda icon, item: quit_app(icon, item, stop_event)),
    )
    tray_icon = pystray.Icon("autoscan", ICONS["idle"], "Autoscan DS-640", menu)

    worker.start()
    tray_icon.run()  # bloque le thread principal jusqu'a "Quitter"


if __name__ == "__main__":
    main()
