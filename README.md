# Patent-Tools (Word)

Ribbon-Tab „Patent-Tools“ mit VBA-Makros für Patentansprüche (DE/EN): Neunummerierung, Anspruchsprüfung, Formatierung, Reinschrift/markierte Fassung, Textbausteine, Entwurfshelfer.

- Anforderungen & Entscheidungen: `docs/project-context.md`
- Arbeitsanweisungen für Claude Code: `CLAUDE.md`
- Referenzmaterial: `docs/reference/` (EP 20 187 485.6)
- Versionen: `CHANGELOG.md`

## Stand (Version 0.3.0)

| Funktion | Status |
|---|---|
| **Neu nummerieren** (MVP-1) | fertig, in Word noch zu testen |
| **Vollprüfung, Einzelprüfung, Prüfmarkierungen entfernen, Prüfbericht** (MVP-2) | fertig, in Word noch zu testen |
| Einstellungen öffnen, Info | fertig |
| Formatierung, Reinschrift/markierte Fassung, Textbausteine, Entwurfshelfer | in Arbeit (siehe Roadmap in `CLAUDE.md`) |

### Prüfung – was passiert
- **Vollprüfung** führt alle in `settings.json` (`check.enabled`) aktivierten Prüfungen aus; **Einzelprüfung** nur eine.
- Jeder Befund wird als Kommentar von „PatentCheck“ mit vorangestelltem Prüfungsnamen (z. B. „[Bezugsbasis] …“) eingefügt und farbig markiert (Farbe je Prüfung: `check.highlightColors`). Die Änderungsverfolgung wird dabei kurz ausgeschaltet, damit keine Formatänderungen aufgezeichnet werden.
- Vor jeder Prüfung werden die alten Prüfmarkierungen entfernt. **Prüfmarkierungen entfernen** löscht nur Kommentare/Markierungen des Werkzeugs – eigene Kommentare bleiben.
- **Prüfbericht** listet alle Befunde als Tabelle in einem neuen Dokument; das Original bleibt unverändert.
- Ansatz „flag it, you decide“: Lieber ein Hinweis zu viel als einer zu wenig, besonders bei Bezugsbasis und unklaren Begriffen.

### Neu nummerieren – was passiert
- Alle Ansprüche (Absatz beginnt mit `Nummer.` + Tab) werden **in Dokumentreihenfolge** fortlaufend 1, 2, 3 … nummeriert – auch Zusatznummern wie `1a`, `1.1`, `1a.1`, `1a1` (jede Nummer ist ein eigenes Kennzeichen, `1a1` ≠ `1a.1`).
- Alle Rückbezüge werden angepasst: „claim 1a“, „claims 1 to 3“, „any one of claims 1, 2 or 4“, „claims 1-3“, „Anspruch 1a“, „Ansprüche 1 bis 3“, „gemäß Patentanspruch 2“, „des Anspruchs 4“ …
- Text vor dem ersten Anspruch (z. B. Überschrift „Patentansprüche“) bleibt unberührt.
- Bei eingeschalteter Änderungsverfolgung werden die Änderungen **als Änderungen aufgezeichnet** (`renumber.trackChangesMode`, alternativ `"silent"`).
- Bereits vorhandene Änderungsmarkierungen werden berücksichtigt: gelöschter Text zählt nicht.
- **Doppelte Nummern** → Abbruch mit Meldung, Dokument bleibt unverändert.
- Unklare Stellen werden nicht geraten, sondern als **Kommentar von „PatentCheck“** markiert: Rückbezug auf nicht vorhandenen Anspruch; Bereich („1 bis 3“), der durch eingefügte/verschobene Ansprüche jetzt etwas anderes umfasst; Nummer, die wegen Feld oder Änderungsmarkierung nicht sicher ersetzt werden kann.
- Alles lässt sich mit **einem** Strg+Z rückgängig machen.

## Installation (Testbetrieb auf deinem PC)

Einmalig in Word: *Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Makroeinstellungen →* „Zugriff auf das VBA-Projektobjektmodell vertrauen“ aktivieren (nur für den Import-Helfer nötig).

1. Repository herunterladen (GitHub → *Code → Download ZIP*) und entpacken.
2. `PatentTools.dotm` besorgen: entweder die von Claude gelieferte Datei oder selbst erzeugen mit `python tools/build.py` (liegt dann in `build/`).
3. Import-Helfer einmalig in die Normal.dotm holen: Word → Alt+F11 → links „Normal“ markieren → *Datei → Datei importieren…* → `tools/PT_ImportModules.bas`.
4. In Word `PatentTools.dotm` über **Datei → Öffnen** öffnen (nicht per Doppelklick – das würde ein neues Dokument daraus erzeugen).
5. Alt+F8 → `PT_ImportModules` → den Ordner `src` des Repositorys wählen. Alle Module werden importiert (ältere Versionen ersetzt).
6. Strg+S, `PatentTools.dotm` schließen.
7. `PatentTools.dotm` in den STARTUP-Ordner kopieren: `%APPDATA%\Microsoft\Word\STARTUP`. Word neu starten → Tab „Patent-Tools“ erscheint.

Beim ersten Start legt das Makro `%APPDATA%\PatentTools\` mit `settings.json`, `blacklist.json` und `textblocks\` an.

**Update auf eine neue Version:** Schritte 4–7 wiederholen (bei geändertem Menüband vorher eine frische `PatentTools.dotm` verwenden).

Falls beim Kompilieren `IRibbonControl` oder `msoFileDialogFolderPicker` unbekannt ist: im VBA-Editor *Extras → Verweise* „Microsoft Office 16.0 Object Library“ aktivieren.

## Testen

**In Word** (empfohlen nach jedem Update):
- `tests/fixtures/renumber_*.docx` öffnen → *Patent-Tools → Neu nummerieren*. Das erwartete Ergebnis steht in `<Name>.expected.txt` (erste Zeile: Anzahl erwarteter Kommentare).
- `tests/fixtures/check_*.docx` und `EP20187485_*.docx` öffnen → *Vollprüfung*. Die erwartete Anzahl Befunde je Prüfung steht in `<Name>.check.txt` (die Meldung am Ende zeigt die Zahlen je Prüfung).
- `renumber_EN_large.docx` (218 Ansprüche) dient auch als Geschwindigkeitstest.
- Optional die Unit-Tests: `tests/vba/modTestRenumber.bas` und `modTestJson.bas` zusätzlich importieren, Alt+F8 → `PT_RunRenumberTests` / `PT_RunJsonTests`.

**Ohne Word** (Entwicklung, Linux/CI): `python3 tools/run_vba_tests.py` führt die Text-Logik in LibreOffice Basic aus (Unit-Tests + alle Fixtures).

## Build

`src/` ist die Quelle. `python3 tools/build.py`
- erzeugt `src/modules/modDefaults.bas` aus `config/settings.json` und `config/blacklist.json` (Standardwerte für den ersten Start),
- prüft die VBA-Quellen (nur ASCII, `Option Explicit`, Modulname = Dateiname),
- baut `build/PatentTools.dotm` mit dem Menüband aus `src/customUI/customUI14.xml`.

Die VBA-Module kommen anschließend per Import-Helfer in die .dotm (siehe oben), weil `vbaProject.bin` nur Office selbst zuverlässig schreiben kann. `python3 tools/build.py --update <Datei>.dotm` tauscht nur das Menüband in einer bestehenden .dotm aus und behält den importierten Code.

Testdokumente neu erzeugen: `python3 tools/make_fixtures.py` (benötigt python-docx).
