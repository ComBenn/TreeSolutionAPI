# TreeSolutionHelper

TreeSolutionHelper ist ein Windows-GUI-Tool zur Aufbereitung von Benutzerexporten fuer den Upload in Zielsysteme.
Das Programm laedt Benutzerdaten aus Excel oder CSV, markiert technische Accounts und Duplikate, verarbeitet Mitarbeiterlisten,
zeigt die aktuelle Auswahl in einer Filter-/Sortieransicht an und erzeugt daraus Upload-CSV-Dateien.

Der normale Einsatz ist die Weitergabe der erzeugten `TreeSolutionHelper.exe`.
Der Python-Quellstand dient dem Build und der Wartung.

## Hauptfunktionen

- Benutzerdatei aus Excel oder CSV laden
- Technische Accounts ueber Keyword-Datei markieren und ausschliessen
- Duplikate ueber Email, Username oder Nachname/Vorname pruefen und gezielt ausschliessen
- Mitarbeiterlisten als wiederverwendbare Vorlagen speichern
- Aktuelle Auswahl interaktiv filtern, sortieren und pruefen
- Upload-CSV fuer die aktuelle Auswahl exportieren
- Batch-Export mit persistenter Merkliste bereits exportierter IDs

## Arbeitsweise

### 1. Benutzerdatei laden

Die Benutzerdatei ist die zentrale Datenquelle.
Erwartet werden mindestens diese Spalten:

- `id`
- `username`
- `email`
- `firstname`
- `lastname`

Je nach Export werden spaeter auch diese Spalten geschrieben oder ueberschrieben:

- `institution`
- `department`
- `auth`

CSV-Dateien werden mit mehreren Encodings eingelesen.
Excel-Dateien werden ueber `pandas` und `openpyxl` verarbeitet.

### 2. Technische Accounts erkennen

Technische Accounts werden ueber die Datei `keywords_technische_accounts.txt` erkannt.

Die Erkennung arbeitet aktuell mit mehreren Regeln:

- exakter Treffer auf `id`
- exakter Treffer auf `firstname`
- exakter Treffer auf `lastname`
- exakter Treffer auf kombinierte Namen wie `uro andro`
- Token-Treffer innerhalb von `firstname` oder `lastname`
- Teilstring-Treffer fuer laengere Keywords
- numerische `firstname` / `lastname`

Wichtig:

- Keywords werden case-insensitive verarbeitet
- kuerzere Begriffe sollten moeglichst als exakte oder kombinierte Keywords gepflegt werden
- problematische Teilstrings wie `andro` wurden bewusst in explizite Kombinationen umgebaut, um False Positives zu vermeiden

Die GUI stellt automatisch eine interne Vorlage `Technische Accounts (Auto)` bereit.
Diese Vorlage ist fest auf `ausschliessen` gesetzt und wird bei jeder Aktualisierung aus der Keyword-Datei neu aufgebaut.

### 3. Mitarbeiterlisten anwenden

Mitarbeiterlisten koennen als Vorlagen gespeichert werden.
Eine Vorlage kann auf `einschliessen` oder `ausschliessen` stehen.

Der Abgleich erfolgt ueber:

- `email`
- Namensvarianten aus `firstname` / `lastname`
- kombinierte Namensspalten in Mitarbeiterlisten

Unterstuetzte typische Spaltennamen in Mitarbeiterlisten:

- `email`
- `firstname`
- `lastname`
- `vorname`
- `nachname`
- kombinierte Spalten wie `lastname firstname` oder `vorname nachname`

Die GUI speichert Vorlagen in der Runtime-`ui_state.json`, nicht im Quellcode.

### 4. Vorschau und Tabellenansicht

Die aktuelle Auswahl kann in einer Tabellenansicht geoeffnet werden.
Dort sind folgende Aktionen moeglich:

- Sortieren per Klick auf den Spaltenkopf
- Filtern per Rechtsklick auf den Spaltenkopf
- mehrere Vorlagen anwenden
- selektierte Eintraege aus der aktuellen Ansicht entfernen
- direkte Exportkontrolle vor dem Schreiben der CSV

### 5. Export

Der Export erzeugt eine semikolon-separierte CSV in `UTF-8 mit BOM`, damit Excel Umlaute korrekt erkennt.

Beim Export werden technische Hilfsspalten entfernt:

- Spalten beginnend mit `flag_`
- Spalten beginnend mit `__`

Zusatzlogik im Export:

- `institution` wird auf `Sonic Suisse SA` gesetzt
- `auth` wird auf `iomadoidc` gesetzt
- `department` kann optional ueber die GUI ueberschrieben werden

## Batch-Export

Der Batch-Export ist fuer schrittweise Auslieferungen gedacht.

Verhalten:

- exportiert nur Datensaetze mit gefuellter `id`
- merkt sich bereits exportierte IDs persistent
- schliesst bereits exportierte IDs bei weiteren Batch-Laeufen automatisch aus
- zeigt an, wie viele Eintraege noch offen sind

Die Merkliste wird zur Laufzeit als `batch_export_tracker.json` neben der EXE gespeichert.

## Runtime-Dateien

Beim Start legt das Programm im Laufzeitverzeichnis automatisch benoetigte Dateien an.
Im EXE-Betrieb ist das typischerweise der Ordner der EXE.

Wichtige Runtime-Dateien:

- `keywords_technische_accounts.txt`
- `ui_state.json`
- `batch_export_tracker.json`

Diese Dateien gehoeren nicht in das Repository und werden lokal erzeugt oder fortgeschrieben.

## Projektstruktur

```text
TreeSolutionAPI/
  src/
    treesolution_helper/
      files/
        main.py
        ui_app.py
        config.py
        io_utils.py
        filters_technical.py
        filters_employee_list.py
        exporter.py
        keywords_technische_accounts.txt
  TreeSolutionHelper.spec
  requirements.txt
```

## Zentrale Module

- `main.py`
  Startpunkt der Anwendung. Startet direkt die GUI.

- `ui_app.py`
  Hauptlogik der Anwendung und komplette GUI.

- `filters_technical.py`
  Erkennung technischer Accounts ueber Keywords und Namensregeln.

- `filters_employee_list.py`
  Matching von Benutzerdaten gegen Mitarbeiterlisten und Vorlagen.

- `exporter.py`
  Aufbau des finalen Upload-Exports.

- `io_utils.py`
  Einlesen, Textnormalisierung und Dateihilfen.

- `config.py`
  Standarddateien, erwartete Spaltennamen und Export-Fixwerte.

## Installation fuer Entwicklung

### Voraussetzungen

- Windows
- Python 3.12

### Setup

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller jinja2
```

### Start aus dem Quellcode

```powershell
.\.venv\Scripts\python.exe src\treesolution_helper\files\main.py
```

## EXE bauen

```powershell
.\.venv\Scripts\python.exe -m PyInstaller TreeSolutionHelper.spec
```

Das Build-Ergebnis liegt danach unter:

```text
dist/TreeSolutionHelper.exe
```

## Versionsverwaltung der EXE

Die fortlaufende EXE-Version wird zentral in [VERSION.txt](/C:/workarea/workspace/TreeSolutionAPI/VERSION.txt) gepflegt.
Der aktuelle Ausgangsstand ist `5.2`.

Fuer lokale versionierte Builds gibt es das Skript:

```powershell
.\scripts\build-versioned-exe.ps1
```

Verhalten des Skripts:

- liest die aktuelle Version aus `VERSION.txt`
- erhoeht die Minor-Version automatisch um `1`
- baut die EXE
- behaelt im `dist`-Ordner nur die versionierte Datei wie `TreeSolutionHelper (V5.3).exe`
- schreibt die neue Version zurueck nach `VERSION.txt`

Beispiel:

- vor dem Build: `5.2`
- nach dem Build: `5.3`
- erzeugte Datei: `dist/TreeSolutionHelper (V5.3).exe`

## GitHub Releases

Das Repository ist so vorbereitet, dass GitHub Actions die aktuelle `TreeSolutionHelper.exe` automatisch baut.

Workflow:

- manueller Start ueber `Actions > Build TreeSolutionHelper > Run workflow`
- automatischer Build beim Veroeffentlichen eines GitHub Releases

Verhalten des Workflows:

- Build auf `windows-latest`
- Installation von `requirements.txt`, `pyinstaller` und `jinja2`
- Build von `TreeSolutionHelper.exe`
- Upload der EXE als Workflow-Artefakt
- bei einem veroeffentlichten GitHub Release zusaetzlich Upload der EXE direkt an das Release

Empfohlener Release-Ablauf:

1. Aenderungen committen und pushen
2. auf GitHub ein neues Release erstellen und veroeffentlichen
3. GitHub Actions baut automatisch die aktuelle EXE
4. die fertige `TreeSolutionHelper.exe` steht danach im Release als Asset bereit

Wichtig:

- du musst die EXE bei diesem Ablauf nicht selbst auf GitHub hochladen
- du musst nur den Code pushen und das GitHub Release veroeffentlichen
- das Hochladen der EXE uebernimmt GitHub Actions

## Aktueller Repository-Ansatz

Das Repository ist auf Quellcode fokussiert.
Generierte Artefakte werden ignoriert:

- `build/`
- `dist/`
- `__pycache__/`
- lokale Runtime-State-Dateien

Wenn alte Build-Dateien bereits historisch versioniert wurden, muessen diese einmalig per Commit aus dem Repository entfernt werden.

## Typische Pflegeaufgaben

### Neues technisches Keyword pflegen

Eintrag in `keywords_technische_accounts.txt` ergaenzen.
Danach die technische Markierung in der GUI aktualisieren oder die App neu starten.

### False Positive bei Personennamen vermeiden

Keine zu generischen Teilstrings als Einzelkeyword pflegen.
Stattdessen moeglichst:

- exakte Keywords
- Token-basierte Begriffe
- explizite Namenskombinationen wie `uro andro`

### Neue EXE fuer Fachbereich erzeugen

1. Quellcode anpassen
2. Build mit PyInstaller ausfuehren
3. `dist/TreeSolutionHelper.exe` weitergeben

## Bekannte technische Hinweise

- Die GUI ist der einzige aktiv gepflegte Einstiegspunkt.
- Die alte CLI-Menue-Struktur wurde entfernt.
- `jinja2` ist in der Entwicklungs-`.venv` installiert, damit der PyInstaller-Build sauberer bleibt.
- Runtime-State wird bewusst nicht aus dem Repository geladen, sondern lokal erzeugt.

## Wartungsempfehlungen

- Keyword-Datei fachlich sauber halten, um False Positives gering zu halten
- Mitarbeiterlisten als Vorlagen statt als einmalige Ad-hoc-Dateien verwenden
- Build-Artefakte nicht versionieren
- Vor groesseren Aenderungen immer mindestens einen kompletten EXE-Build testen
