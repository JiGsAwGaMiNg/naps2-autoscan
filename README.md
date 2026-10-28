# Autoscan DS-640

Scan automatique multi-pages pour scanner Brother DS-640 (ou tout scanner
WIA a bac feuille-a-feuille) : des qu'une feuille est inseree, elle est
scannee automatiquement ; si une nouvelle feuille arrive dans le delai
imparti, elle est ajoutee au meme PDF ; sinon le PDF est finalise et
enregistre.

Chaque PDF final est recadre sur la taille de page reelle, redresse
(deskew), reoriente automatiquement si une page a ete inseree a l'envers
ou de cote, et passe a l'OCR pour produire un PDF dont le texte est
cherchable/copiable.

Tourne en tache de fond, sans fenetre, avec une icone dans la barre des
taches (bleu = en attente, vert = document en cours de constitution, rouge
= erreur).

## Pourquoi ce projet

L'API WIA "Automation Layer" (`wiaaut.dll`, celle qu'on pilote facilement
depuis Python via `pywin32`/`win32com.client.Dispatch("WIA.DeviceManager")`)
s'est averee **incompatible avec le driver Brother DS-640** : tout appel a
`Item.Transfer()` echoue systematiquement apres 10 a 30 secondes, meme avec
une feuille correctement chargee, alors que "Telecopie et numerisation
Windows" (qui utilise l'API WIA 2.0 native) fonctionne parfaitement.

Ce projet contourne le probleme en delegant le scan a
[NAPS2](https://github.com/cyanfish/naps2) (`NAPS2.Console.exe`), qui
utilise la meme API native, et echoue en ~2 secondes (au lieu de 10-30s)
quand le bac est vide -- ce qui permet un vrai polling reactif.

## Prerequis

- Windows
- Python 3.10+
- Un scanner deja installe et fonctionnel dans Windows (visible dans
  "Telecopie et numerisation Windows" ou l'app "Numeriser")
- [NAPS2](https://github.com/cyanfish/naps2/releases) version portable
  Windows (`naps2-x.y.z-win-x64.zip`) -- **non fourni dans ce depot**,
  a extraire dans le dossier `NAPS2/` a la racine du projet (ou ailleurs,
  en renseignant `NAPS2_CONSOLE` dans `.env`)
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (utilise
  uniquement pour la detection d'orientation des pages, distinct du moteur
  OCR interne de NAPS2) :
  ```bash
  winget install -e --id UB-Mannheim.TesseractOCR
  ```

## Installation

```bash
pip install -r requirements.txt
```

Telecharger NAPS2 (version `win-x64.zip`) depuis les
[releases GitHub](https://github.com/cyanfish/naps2/releases) et extraire
le contenu du zip dans un dossier `NAPS2/` a la racine du projet, de sorte
que `NAPS2/App/NAPS2.Console.exe` existe.

Copier `.env.example` vers `.env` et l'adapter :

```bash
cp .env.example .env
```

Pour connaitre le nom exact de ton scanner :

```bash
NAPS2\App\NAPS2.Console.exe --driver wia --listdevices
```

Renseigner ce nom dans `DEVICE_NAME` (une correspondance partielle
suffit).

Installer les packs de langue OCR de NAPS2 (independants du Tesseract
systeme) pour les langues souhaitees, ex. francais et anglais :

```bash
NAPS2\App\NAPS2.Console.exe --install ocr-fra
NAPS2\App\NAPS2.Console.exe --install ocr-eng
```

## Utilisation

Lancement manuel (visible, avec console) :

```bash
python autoscan.py
```

Lancement invisible (recommande, icone barre des taches uniquement) :

```bash
pythonw autoscan.py
```

Ou via [run_autoscan.bat](run_autoscan.bat).

### Lancement automatique a l'ouverture de session

Placer un raccourci vers `run_autoscan.bat` dans :

```
%AppData%\Microsoft\Windows\Start Menu\Programs\Startup
```

## Configuration (`.env`)

| Variable        | Description                                                        | Defaut                                  |
|-----------------|----------------------------------------------------------------------|------------------------------------------|
| `OUTPUT_FOLDER` | Dossier ou sont enregistres les PDF et le fichier `autoscan.log`   | *(obligatoire)*                          |
| `WAIT_SECONDS`  | Delai d'inactivite avant de finaliser le PDF en cours (secondes)   | `30`                                      |
| `POLL_INTERVAL` | Intervalle entre deux tentatives quand le bac est vide (secondes)  | `2`                                       |
| `NAPS2_CONSOLE` | Chemin vers `NAPS2.Console.exe`                                     | `<dossier du script>/NAPS2/App/NAPS2.Console.exe` |
| `DEVICE_NAME`   | Nom (ou partie du nom) du scanner                                   | *(obligatoire)*                          |
| `DPI`           | Resolution de scan (300 = bon compromis OCR, 600 = qualite max du DS-640) | `300`                               |
| `COLOR_MODE`    | `color`, `gray` ou `bw`                                             | `color`                                   |
| `JPEG_QUALITY`  | Qualite JPEG (0-100) des images dans le PDF final                   | `90`                                      |
| `PAGE_SIZE`     | Zone de capture maximale (pas la taille du papier, voir `AUTO_CROP`)  | `8.5x14in`                                |
| `DESKEW`        | Redresser automatiquement les pages legerement de travers            | `true`                                    |
| `AUTO_CROP`     | Recadrer automatiquement chaque page sur son contenu reel, quel que soit son format | `true`                   |
| `AUTO_CROP_THRESHOLD` | Seuil (0-255) distinguant papier et fond hors-page              | `25`                                      |
| `AUTO_ROTATE`   | Corriger l'orientation d'une page inseree a l'envers/de cote          | `true`                                    |
| `OCR_LANG`      | Langue(s) OCR (codes 3 lettres, combinables avec `+`)                | `fra+eng`                                 |
| `TESSERACT_CMD` | Chemin vers `tesseract.exe` (pour `AUTO_ROTATE` uniquement)          | `C:\Program Files\Tesseract-OCR\tesseract.exe` |

## Fonctionnement

1. Boucle en continu : tente un scan depuis le bac (`--source feeder`) sur
   la zone maximale `PAGE_SIZE`, redresse (`--deskew`).
2. Bac vide -> echec rapide (~2s), on reessaie apres `POLL_INTERVAL`.
3. Feuille scannee -> image ajoutee au document en cours.
4. Si plus aucune feuille pendant `WAIT_SECONDS` :
   - chaque page est recadree automatiquement sur son contenu reel
     (`AUTO_CROP`), quel que soit son format d'origine. Le cadre se cale
     sur le texte dense, puis est repousse cote par cote si un element
     epars (filet, bandeau, tampon...) le traverse, pour ne rien couper
     (l'extension est notee dans `autoscan.log` : « cadre etendu ») ;
   - chaque page est reorientee automatiquement si necessaire (detection
     Tesseract OSD, `AUTO_ROTATE`) ;
   - toutes les pages sont fusionnees et passees a l'OCR en un seul PDF
     (`scan_AAAA-MM-JJ_HH-MM-SS.pdf`) dans `OUTPUT_FOLDER` ;
   - l'attente reprend pour un nouveau document.

## Tickets de caisse et documents longs

Le DS-640 limite les scans a 14 pouces (35,6 cm) par defaut via l'API WIA
standard. Le mode "Long Receipts" du driver Brother necessite l'API WIA
bas niveau (pas la couche simplifiee utilisable facilement depuis Python,
qui est aussi celle dont le transfert d'image est casse sur ce modele --
voir plus haut). Tenter de forcer une taille superieure a 14 pouces
provoque un **bourrage papier confirme** (pas juste une troncature).

Pour un document/ticket plus long que 14 pouces : scanne-le en plusieurs
passages de 14 pouces maximum, en reinserant la suite avant l'expiration
de `WAIT_SECONDS`. Chaque passage devient une page du meme PDF final.

## Limitations connues

- L'evenement "bouton de numerisation" du scanner (declenchement physique)
  n'est pas utilise : Windows ne permet pas d'enregistrer facilement un
  script tiers comme gestionnaire de cet evenement sur les versions
  recentes (le mecanisme `StillImage\Events` classique est en grande
  partie obsolete). L'approche par polling s'en passe entierement.
- La mise en veille automatique du scanner (Auto Power Off) n'est pas
  geree par ce script -- desactive-la directement sur l'appareil via
  l'utilitaire Brother Remote Setup si elle te derange.

## Licence

MIT
