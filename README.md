# Serbisch B2

Persönliche Vokabel-PWA für Serbisch B2. FSRS-Spaced-Repetition, GitHub-Pages-gehostet, Offline-fähig.

## Setup auf GitHub

1. Neues Repo auf github.com erstellen, z.B. `serbisch-b2` (kann ein beliebiger obskurer Name sein, niemand sucht aktiv danach)
2. Alle Files aus diesem ZIP unverändert ins Repo pushen:
   ```bash
   cd serbisch-b2
   git init
   git add .
   git commit -m "init"
   git branch -M main
   git remote add origin https://github.com/<dein-username>/serbisch-b2.git
   git push -u origin main
   ```
3. Im Repo unter `Settings → Pages`:
   - **Source**: GitHub Actions
4. Erster Push triggert die Action automatisch. Nach ~1 Minute ist die App live unter `https://<dein-username>.github.io/serbisch-b2/`
5. URL auf dem Handy öffnen, im Browser-Menü "Zum Home-Bildschirm hinzufügen". Fertig.

## Wortliste pflegen

Die Wortliste lebt in `source/wortliste.xlsx`. Bei Änderungen einfach committen und pushen — die GitHub Action regeneriert `data/words.json` und deployed automatisch.

```bash
# Datei lokal bearbeiten
git add source/wortliste.xlsx
git commit -m "neue Wörter"
git push
```

## Lokal testen

```bash
# Python http-server, dann im Browser auf http://localhost:8000
python3 -m http.server 8000
```

Service Worker funktioniert nur über HTTPS oder localhost — auf einer "geöffneten" file://-URL also nicht.

## Datenformat

`data/words.json` — wird automatisch von `scripts/build.py` aus der xlsx generiert.

`data/examples.json` — manuell gepflegte Beispielsätze für Verben:
```json
{
  "bringen|donositi / doneti": {
    "ipf": ["serbischer Satz", "deutsche Übersetzung"],
    "pf":  ["serbischer Satz", "deutsche Übersetzung"]
  }
}
```
Match über Schlüssel `<DE>|<SR-Latinica>`.

`data/forms.json` — grammatische Anreicherung pro Wort, gleiches Key-Schema:
```json
{
  "schwer|težak": {
    "mfn": ["težak", "teška", "teško"]
  },
  "Mensch|čovek": {
    "genus": "m",
    "plural": "ljudi"
  },
  "gehen|ići": {
    "pres1": "idem"
  }
}
```

`data/corpus.json` — wird aus den Idem-Dalje-docx in `source/` generiert. Interne Referenz für Beispielsatz-Generierung, nicht in der App sichtbar.

## Privacy

`<meta name="robots" content="noindex, nofollow">` blockiert Suchmaschinen-Indexing. Die App ist über URL zugänglich, aber niemand wird zufällig drauf stoßen. Lernfortschritt liegt im localStorage des Browsers, nicht auf dem Server.

## Backup

In der App: "Backup speichern" lädt eine JSON-Datei mit allem Fortschritt herunter. Diese in iCloud Drive / Google Drive ablegen. Bei Datenverlust (z.B. iOS räumt Storage) kannst du die Datei über Settings → "Fortschritt importieren" wiederherstellen.

## Struktur

```
serbisch-b2/
├── index.html                     # Markup
├── styles.css                     # Styling
├── app.js                         # Main app
├── fsrs.js                        # FSRS-Algorithmus
├── sw.js                          # Service Worker (offline cache)
├── manifest.json                  # PWA Manifest
├── icons/                         # PWA Icons
├── data/                          # Generierte JSON-Daten
│   ├── words.json
│   ├── examples.json
│   ├── forms.json
│   └── corpus.json
├── scripts/
│   ├── build.py                   # xlsx → words.json
│   └── parse_idemo.py             # docx → corpus.json
├── source/                        # Source of Truth
│   ├── wortliste.xlsx
│   └── Serbisch_-_Idemo_Dalje_*.docx
└── .github/workflows/
    └── build-data.yml             # CI: build & deploy
```
