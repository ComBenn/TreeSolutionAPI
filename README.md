# TreeSolutionHelper

TreeSolutionHelper ist ein Windows-GUI-Tool zur Aufbereitung von Benutzerexporten fuer den Upload in Zielsysteme.
Die Anwendung laedt Benutzerdaten aus Excel oder CSV, markiert technische Accounts und Duplikate, verarbeitet Mitarbeiterlisten als Vorlagen, zeigt die aktuelle Auswahl in Tabellenfenstern an und erzeugt daraus Upload-CSV-Dateien.

Der normale Einsatz ist die Weitergabe der erzeugten `TreeSolutionHelper.exe`.
Der Python-Quellstand dient dem Build, der Pflege und dem Nachvollziehen der fachlichen Regeln.

## Status des Projekts

- aktiv gepflegte Windows-Tkinter-GUI
- Einstiegspunkt ist `src/treesolution_helper/files/main.py`
- Build ueber PyInstaller
- automatischer GitHub-Workflow fuer Release- und manuelle Builds vorhanden
- automatischer Runtime-State fuer UI-Einstellungen, Keywords und Batch-Merkliste vorhanden

Aktuelle Versionsdatei: [VERSION.txt](/C:/workarea/workspace/TreeSolutionAPI/VERSION.txt)

## Hauptfunktionen

- Benutzerdatei aus Excel oder CSV laden
- technische Accounts ueber Keyword-Datei automatisch markieren
- Duplikate ueber Email, Username und Namenskombinationen erkennen
- Duplikate manuell in einem eigenen Prueffenster ausschliessen oder beibehalten
- Mitarbeiterlisten als wiederverwendbare Vorlagen speichern
- Vorlagen auf `einschliessen` oder `ausschliessen` setzen
- aktuelle Auswahl in einer interaktiven Tabellenansicht pruefen
- regulaeren Export und Batch-Export erzeugen
- bereits batch-exportierte IDs persistent merken

## Voraussetzungen

- Windows
- Python 3.12 fuer Entwicklung und Build

Python-Abhaengigkeiten aus [requirements.txt](/C:/workarea/workspace/TreeSolutionAPI/requirements.txt):

- `pandas>=2.2,<3.0`
- `openpyxl>=3.1,<4.0`

Fuer den EXE-Build werden zusaetzlich verwendet:

- `pyinstaller`
- `jinja2`

## Projektstruktur

```text
TreeSolutionAPI/
  .github/
    workflows/
      build-release.yml
  scripts/
    build-versioned-exe.ps1
  src/
    treesolution_helper/
      files/
        main.py
        ui_app.py
        state.py
        io_utils.py
        config.py
        exporter.py
        export_service.py
        export_dialogs.py
        duplicate_dialogs.py
        filters_technical.py
        filters_duplicates.py
        filters_employee_list.py
        template_service.py
        auto_template_service.py
        keywords_technische_accounts.txt
  tests/
  TreeSolutionHelper.spec
  VERSION.txt
  requirements.txt
```

## Bedienkonzept

### 1. Benutzerdatei laden

Die Benutzerdatei ist die zentrale Datenquelle.
Erwartet werden mindestens diese Spalten:

- `id`
- `username`
- `email`
- `firstname`
- `lastname`

Je nach Export werden ausserdem diese Spalten gesetzt oder neu erzeugt:

- `institution`
- `department` oder `department1`, `department2`, ...
- `auth`

CSV-Dateien werden mit mehreren Encodings eingelesen:

- `utf-8-sig`
- `utf-8`
- `cp1252`
- `latin-1`

Der Trenner fuer CSV wird automatisch erkannt.
Excel-Dateien werden ueber `pandas.read_excel(...)` gelesen; bei leerem Sheet-Eintrag wird das erste Tabellenblatt verwendet.

### 2. Technische Accounts erkennen

Technische Accounts werden ueber die Keyword-Datei `keywords_technische_accounts.txt` erkannt.
Die aktuelle Erkennung verwendet unter anderem:

- exakte Treffer auf `id`
- exakte Treffer auf `firstname`
- exakte Treffer auf `lastname`
- exakte Treffer auf kombinierte Namen
- Token-Treffer innerhalb von `firstname` oder `lastname`
- Teilstring-Treffer fuer laengere Keywords
- numerische `firstname` oder `lastname`

Keywords werden case-insensitive verarbeitet.

Die GUI verwaltet automatisch die interne Auto-Vorlage:

- `Technische Accounts (Auto)`
- Modus: `ausschliessen`
- `readonly`
- Quelle: `<auto:keywords_technische_accounts>`

Diese Vorlage wird aus den aktuellen Markierungen neu aufgebaut.

### 3. Duplikate pruefen

Duplikate werden aus der Original-Benutzerdatei erzeugt und nach folgenden Merkmalen gruppiert:

- gleiche Email
- gleicher Username
- gleiche Namenskombination

Das Duplikatfenster zeigt nur als Duplikat markierte Datensaetze.
Dort sind aktuell moeglich:

- Sortieren per Linksklick auf den Spaltenkopf
- Filtern per Rechtsklick auf den Spaltenkopf
- Ausschluss einzelner Zeilen per Checkbox
- `Alle außer erster ausschließen`
- Export der eingeschlossenen Duplikate
- Export der ausgeschlossenen Duplikate
- Speichern der Auswahl

Wichtige Regel:
Pro Duplikat-Gruppe muss mindestens ein Datensatz aktiv bleiben.

Die gespeicherten Ausschluesse landen in der internen Auto-Vorlage:

- `Duplikate ausgeschlossen (Auto)`
- Modus: `ausschliessen`
- `readonly`
- Quelle: `<auto:duplicate_review>`

Zusatzlich werden die ausgeschlossenen IDs in `ui_state.json` gespeichert.

### 4. Mitarbeiterlisten als Vorlagen

Mitarbeiterlisten koennen als Vorlagen gespeichert und spaeter erneut angewendet werden.
Eine Vorlage kann auf `einschliessen` oder `ausschliessen` stehen.

Der Abgleich erfolgt ueber:

- `email`
- `firstname` und `lastname`
- kombinierte Namensspalten

Unterstuetzte typische Spaltenbezeichnungen in Mitarbeiterlisten:

- `email`
- `firstname`
- `lastname`
- `vorname`
- `nachname`
- `first name`
- `given name`
- `lastname firstname`
- `firstname lastname`
- `vorname nachname`

Vorlagen werden in der UI verwaltet und in `ui_state.json` persistiert.
Die Auto-Vorlagen fuer technische Accounts und Duplikate erscheinen ebenfalls in dieser Liste und sind schreibgeschuetzt.

### 5. Tabellenansichten

Es gibt mehrere Tabellenfenster im Projekt:

- `Aktuelle Auswahl anzeigen und exportieren`
- Batch-Export-Fenster
- Duplikat-Prueffenster
- Auswahlansicht fuer gespeicherte Vorlagen

Die normale Tabellenansicht fuer die aktuelle Auswahl unterstuetzt:

- Sortieren per Klick auf den Spaltenkopf
- Filtern per Rechtsklick auf den Spaltenkopf
- mehrere gleichzeitige Spaltenfilter
- `Alle Filter löschen`
- Entfernen selektierter Zeilen aus der aktuellen Auswahl
- direkten Export der aktuellen Auswahl im Fenster

Wichtig:
Die Header-Filter in den Tabellenfenstern filtern die Anzeige.
Der regulaere Export arbeitet mit der aktuellen Fensterauswahl nach echten Aenderungen wie Zeilenentfernung, nicht nur mit der gerade sichtbaren Filteransicht.

Die Batch-Ansicht bietet dieselben Sortier- und Filtermechanismen fuer die Anzeige der noch nicht exportierten Datensaetze.
Die Batch-Filter aendern nicht die Exportlogik; exportiert wird weiterhin der naechste Batch aus der ungefilterten verbleibenden Auswahl.

### 6. Export

Der regulaere Export erzeugt eine semikolon-separierte CSV in `UTF-8 mit BOM` (`utf-8-sig`), damit Excel Umlaute sauber erkennt.

Beim Export werden technische Hilfsspalten entfernt:

- Spalten beginnend mit `flag_`
- Spalten beginnend mit `__`

Exportregeln:

- `institution` wird immer auf `Sonic Suisse SA` gesetzt
- `auth` wird immer auf `iomadoidc` gesetzt
- bei Department-Overrides wird `department` entfernt und durch `department1`, `department2`, ... ersetzt
- ohne Override bleibt `department` als normale Spalte erhalten oder wird leer angelegt

Mehrere Department-Werte koennen in der GUI eingetragen werden.

## Batch-Export

Der Batch-Export ist fuer schrittweise Auslieferungen gedacht.

Verhalten:

- exportiert nur Datensaetze mit gefuellter `id`
- schliesst bereits gemerkte IDs automatisch aus
- zeigt an, wie viele IDs batch-faehig, bereits exportiert und noch offen sind
- merkt neu exportierte IDs persistent

Die Merkliste liegt in:

- `batch_export_tracker.json`

Falls diese Datei ungueltig ist, wird sie nicht verwendet.
Der Code legt dann eine Sicherungsdatei wie `batch_export_tracker.json.invalid` an und verlangt vor weiterem Batch-Export ein Zuruecksetzen der Merkliste.

## Runtime-Dateien

Beim Start erzeugt oder kopiert die Anwendung benoetigte Dateien in das Laufzeitverzeichnis.
Im EXE-Betrieb ist das typischerweise der Ordner neben der EXE.
Im Quellbetrieb ist es das Verzeichnis `src/treesolution_helper/files`.

Wichtige Runtime-Dateien:

- `README.md`
- `keywords_technische_accounts.txt`
- `ui_state.json`
- `batch_export_tracker.json`

Persistiert werden unter anderem:

- letzter Pfad zur Benutzerdatei
- Users-Sheet
- Keyword-Datei
- Output-Datei
- Department-Overrides
- Mitarbeiterlisten-Vorlagen
- ausgeschlossene Duplikat-IDs

## Wichtige Module

- [main.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/main.py)
  Einstiegspunkt, startet die GUI.

- [ui_app.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/ui_app.py)
  Hauptfenster, UI-Zustand, Dateiauswahl, Vorlagenverwaltung, Exportfluss.

- [state.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/state.py)
  Runtime-Dateien, Laden/Zuruecksetzen von Daten, Batch-Merkliste.

- [filters_technical.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/filters_technical.py)
  Markierung technischer Accounts.

- [filters_duplicates.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/filters_duplicates.py)
  Duplikaterkennung und Gruppierung.

- [filters_employee_list.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/filters_employee_list.py)
  Matching gegen Mitarbeiterlisten.

- [duplicate_dialogs.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/duplicate_dialogs.py)
  Duplikat-Prueffenster mit Sortierung und Header-Filter.

- [export_dialogs.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/export_dialogs.py)
  Tabellenansicht der aktuellen Auswahl und Batch-Export-Fenster.

- [exporter.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/exporter.py)
  Aufbau des finalen Upload-DataFrames und CSV-Export.

- [template_service.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/template_service.py)
  Persistente Vorlagen, interne ID-Listen, Include-/Exclude-Anwendung.

- [auto_template_service.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/auto_template_service.py)
  Aufbau und Aktualisierung der internen Auto-Vorlagen.

- [io_utils.py](/C:/workarea/workspace/TreeSolutionAPI/src/treesolution_helper/files/io_utils.py)
  Laden von Tabellen, Textnormalisierung, Keyword-Dateien.

## Entwicklung

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

### Tests

Die Tests liegen unter `tests/` und werden mit `unittest` ausgefuehrt.
Beispiel:

```powershell
python -m unittest
```

## EXE bauen

### Direktbuild

```powershell
.\.venv\Scripts\python.exe -m PyInstaller TreeSolutionHelper.spec
```

Ergebnis:

```text
dist/TreeSolutionHelper.exe
```

### Versionierter lokaler Build

Fuer lokale versionierte Builds existiert das Skript [build-versioned-exe.ps1](/C:/workarea/workspace/TreeSolutionAPI/scripts/build-versioned-exe.ps1).

```powershell
.\scripts\build-versioned-exe.ps1
```

Verhalten des Skripts:

- liest die aktuelle Version aus `VERSION.txt`
- erhoeht die Minor-Version automatisch um `1`
- baut die EXE
- erstellt zusaetzlich eine versionierte Datei wie `TreeSolutionHelper (V5.5).exe`
- entfernt die unversionierte Datei `dist/TreeSolutionHelper.exe`
- schreibt die neue Version nach `VERSION.txt`

Beispiel:

- vor dem Build: `5.4`
- nach dem Build: `5.5`
- erzeugte Datei: `dist/TreeSolutionHelper (V5.5).exe`

## GitHub Actions

Der Workflow [build-release.yml](/C:/workarea/workspace/TreeSolutionAPI/.github/workflows/build-release.yml) ist vorhanden.

Ausloeser:

- manuell ueber `workflow_dispatch`
- automatisch bei einem veroeffentlichten GitHub Release

Verhalten:

- Build auf `windows-latest`
- Python `3.12`
- Installation von `requirements.txt`, `pyinstaller` und `jinja2`
- Build mit `pyinstaller TreeSolutionHelper.spec`
- Kopie der EXE in eine versionierte Datei basierend auf `VERSION.txt`
- Upload als Workflow-Artefakt
- bei Release-Events zusaetzlich Upload an das GitHub Release

## Bekannte technische Hinweise

- Die GUI ist der einzige aktiv gepflegte Einstiegspunkt.
- Das Projekt ist auf Windows-Betrieb ausgelegt.
- Im Code werden ASCII-Dateien bevorzugt; das README ist ebenfalls ASCII-kompatibel gehalten.
- `ui_state.json` und `batch_export_tracker.json` sind Laufzeitdateien und keine fachlichen Stammdaten.
- Die EXE bringt `README.md` und `keywords_technische_accounts.txt` als Bundledaten mit und kopiert sie beim ersten Start ins Laufzeitverzeichnis.

## Pflegehinweise

- keine zu generischen Keywords pflegen, um False Positives bei technischen Accounts zu vermeiden
- Mitarbeiterlisten nach Moeglichkeit als wiederverwendbare Vorlagen pflegen statt ad hoc zu laden
- vor einer Weitergabe an Fachbereiche mindestens einen kompletten EXE-Build testen
- `dist/`, `build/`, `__pycache__/` und Runtime-State-Dateien nicht versionieren
