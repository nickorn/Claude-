# Patent-Tools (Word)

Ribbon-Tab „Patent-Tools“ mit VBA-Makros für Patentansprüche (DE/EN): Neunummerierung, Anspruchsprüfung, Formatierung, Reinschrift/markierte Fassung, Textbausteine, Entwurfshelfer.

- Anforderungen & Entscheidungen: `docs/project-context.md`
- Arbeitsanweisungen für Claude Code: `CLAUDE.md`
- Referenzmaterial: `docs/reference/` (EP 20 187 485.6)
- Versionen: `CHANGELOG.md`

## Stand (Version 0.8.1)

| Funktion | Status |
|---|---|
| **Neu nummerieren** (MVP-1) | fertig, in Word noch zu testen |
| **Vollprüfung, Einzelprüfung, Prüfmarkierungen entfernen, Prüfbericht** (MVP-2) | fertig, in Word noch zu testen |
| Einstellungen öffnen, Info | fertig |
| **Formatierung anwenden, Doppelte Leerzeichen, Zeichensetzung, Bezugszeichen-Format** (MVP-3) | fertig, in Word noch zu testen |
| **Reinschrift erzeugen, Markierte Fassung erzeugen, Fußzeile/Status** (Roadmap 4) | fertig, in Word noch zu testen |
| **Textbausteine** (Roadmap 5) | fertig, in Word noch zu testen |
| **Anspruchsbaum, Abhängigen Anspruch anlegen, Ansprüche zusammenführen** (Roadmap 6) | fertig, in Word noch zu testen |
| **Verfahren ↔ Vorrichtung, Ansprüche → Beschreibung, Bezugszeichen entfernen/einfügen** (Roadmap 7) | fertig, in Word noch zu testen |

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

### Einmalige Vorbereitung in Word
1. *Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Makroeinstellungen →* „Zugriff auf das VBA-Projektobjektmodell vertrauen“ anhaken (nur für den Import-Helfer nötig).
2. *Datei → Optionen → Erweitert → Allgemein →* „Fehler in Benutzeroberfläche von Add-Ins anzeigen“ anhaken. Sonst lässt Word das Menüband bei einem Fehler kommentarlos weg; so siehst du die Fehlermeldung (nur für den Testbetrieb).

### Installation
1. Repository herunterladen (GitHub → *Code → Download ZIP*). **Vor dem Entpacken:** Rechtsklick auf die ZIP-Datei → *Eigenschaften* → unten „Zulassen“ anhaken → *OK*. Sonst markiert Windows alle Dateien als „aus dem Internet“ und Office blockiert die Makros. Dasselbe für eine separat erhaltene `PatentTools.dotm`.
2. ZIP entpacken. `PatentTools.dotm` (von Claude geliefert oder mit `python tools/build.py` erzeugt, liegt dann in `build/`) in einen Arbeitsordner legen, z. B. den Desktop – **noch nicht** in den STARTUP-Ordner.
3. Import-Helfer einmalig in die Normal.dotm holen: Word → Alt+F11 → links „Normal“ markieren → *Datei → Datei importieren…* → `tools/PT_ImportModules.bas`.
4. In Word die `PatentTools.dotm` über **Datei → Öffnen** öffnen (nicht per Doppelklick – das würde ein neues Dokument daraus erzeugen).
5. Alt+F8 → `PT_ImportModules` → *Ausführen* → den Ordner `src` des entpackten Repositorys wählen. Die Meldung nennt die Zahl der importierten Module.
6. Alt+F11 → Menü *Debuggen → Kompilieren von TemplateProject* (bzw. „PatentTools“). Es darf **keine** Meldung erscheinen. Falls doch: Screenshot an Claude.
7. Strg+S, die `PatentTools.dotm` schließen.
8. **Word ganz schließen.** Die Datei in den STARTUP-Ordner kopieren: im Explorer `%APPDATA%\Microsoft\Word\STARTUP` in die Adresszeile eingeben.
9. Word starten → Tab „Patent-Tools“ erscheint. *Info* zeigt die Version.

Beim ersten Start legt das Makro `%APPDATA%\PatentTools\` mit `settings.json`, `blacklist.json` und `textblocks\` an.

**Textbausteine einrichten:** *Textbausteine → Beispielbausteine importieren…* → Ordner `config\textblocks` des Repositorys wählen (kopiert die neutralen Beispiele, überschreibt nichts). Eigene Bausteine einfach als `.docx` in den Ordner legen (*Ordner öffnen*), z. B. `EPA\15_Mein Baustein.docx` – erscheint sofort im Menü.

### Update auf eine neue Version
Word schließen, die neue `PatentTools.dotm` (oder die alte aus dem Arbeitsordner) verwenden und die Schritte 1–9 wiederholen; im STARTUP-Ordner die alte Datei überschreiben. Word sperrt die Datei im STARTUP-Ordner, solange es läuft – daher vorher schließen. Deine Einstellungen in `%APPDATA%\PatentTools\` bleiben erhalten; neue Einstellungen gelten automatisch mit ihren Standardwerten.

### Wenn etwas nicht klappt
| Symptom | Ursache / Abhilfe |
|---|---|
| Tab „Patent-Tools“ fehlt | Liegt die Datei im STARTUP-Ordner? *Datei → Optionen → Add-Ins → Verwalten: Vorlagen → Gehe zu* – ist „PatentTools.dotm“ angehakt? Mit Vorbereitung 2 zeigt Word eine Fehlermeldung zum Menüband – bitte an Claude schicken. |
| Gelber/roter Balken „Makros deaktiviert“ bzw. „Sicherheitsrisiko“ | Datei war als „aus dem Internet“ markiert: Installation Schritt 1 („Zulassen“), dann Schritte 2–9. Der STARTUP-Ordner ist standardmäßig ein vertrauenswürdiger Speicherort (*Trust Center → Vertrauenswürdige Speicherorte*). |
| Beim Kompilieren „Mehrdeutiger Name“ | Ein Modul ist doppelt vorhanden (z. B. `modCheck` und `modCheck1`): das doppelte im VBA-Editor entfernen und Schritt 5 wiederholen. |
| `IRibbonControl` oder `msoFileDialogFolderPicker` unbekannt | VBA-Editor → *Extras → Verweise* → „Microsoft Office 16.0 Object Library“ anhaken. |
| `PT_ImportModules` meldet einen Fehler | Vorbereitung 1 fehlt, oder die `PatentTools.dotm` ist nicht über *Datei → Öffnen* geöffnet. |
| Eine Funktion meldet „Unerwarteter Fehler“ | Mit Strg+Z rückgängig machen und die Meldung samt Testdokument an Claude schicken. |

## Einstellungen (`%APPDATA%\PatentTools\settings.json`)

Öffnen mit *Einstellungen → Einstellungen öffnen* (Editor). Änderungen gelten beim nächsten Klick auf eine Funktion – kein Neustart nötig. Fehlt ein Wert oder ist die Datei fehlerhaft, gilt der Standardwert (mit Hinweis).

| Schlüssel | Bedeutung (Standard) |
|---|---|
| `language.claimLanguageDetection` | Anspruchssprache: `"auto"`, `"de"` oder `"en"` |
| `format.fontName`, `fontSize`, `lineSpacing`, `spaceBefore`, `spaceAfter` | Soll-Format für Prüfung und *Formatierung anwenden* (Arial, 11, 1,5, 0, 0) |
| `format.subFeatureIndentCm` | Einzugstiefe der Merkmale (1,25 cm) |
| `format.applyToWholeDocument` | Format für das ganze Dokument statt nur ab Anspruch 1 (`true`) |
| `format.fixEmptyParagraphs` | *Formatierung anwenden* setzt genau einen Leerabsatz zwischen Ansprüchen (`true`) |
| `format.layoutTrackChanges` / `textTrackChanges` | `"silent"` = ohne Änderungsverfolgung, `"tracked"` = wie im Dokument eingestellt (`"silent"` / `"tracked"`) |
| `punctuation.enforceAndBeforeLastFeature` | „; and“/„; und“ vor dem letzten Merkmal verlangen (`true`) |
| `punctuation.flagWhereinAfterFullStop` | Text nach dem Schlusspunkt melden (`true`) |
| `referenceSigns.style` | `"parentheses"` = „(150)“; `"plain"` schaltet die Bezugszeichen-Prüfung ab |
| `referenceSigns.multipleSeparator` | Trennzeichen bei mehreren Bezugszeichen (`", "`) |
| `referenceSigns.normaliseSemicolon` | „;“ zwischen Bezugszeichen melden (`true`) |
| `renumber.trackChangesMode` | Neu nummerieren/Zusammenführen: `"tracked"` oder `"silent"` (`"tracked"`) |
| `check.commentAuthor` | Autorname der Prüfkommentare (`"PatentCheck"`) |
| `check.createReport` | nach der Vollprüfung zusätzlich den Prüfbericht öffnen (`false`) |
| `check.flagEveryMultipleDependency` | jeden mehrfach abhängigen Anspruch melden, nicht nur mehrfach auf mehrfach (`false`) |
| `check.maxFormattingFindings` | höchstens so viele Formatbefunde einzeln kommentieren (20) |
| `check.enabled.<prüfung>` | einzelne Prüfungen der Vollprüfung ein/aus |
| `check.highlightColors.<prüfung>` | Markierungsfarbe je Prüfung (`wdYellow`, `wdRed`, `wdBrightGreen`, …, `none`) |
| `footer.status.de/en.*` | Texte für Fußzeile links |
| `versions.markedUpKeepComments` | Kommentare in der markierten Fassung behalten (`false`) |
| `description.introPhrases.de/en` | Auswahl der Einleitungen für *Ansprüche → Beschreibung* |
| `textblocks.folder` | Ordner der Textbausteine (später z. B. ein Netzlaufwerk) |
| `textblocks.askForPlaceholders` / `highlightUnfilledPlaceholders` | Platzhalter abfragen / unausgefüllte gelb markieren (`true`) |

Die Liste unklarer Begriffe steht in `blacklist.json` (`"de"` und `"en"`, ganze Wörter, Groß-/Kleinschreibung egal): *Einstellungen → Blacklist bearbeiten*.

## Testen

**In Word** (empfohlen nach jedem Update):
- `tests/fixtures/renumber_*.docx` öffnen → *Patent-Tools → Neu nummerieren*. Das erwartete Ergebnis steht in `<Name>.expected.txt` (erste Zeile: Anzahl erwarteter Kommentare).
- `tests/fixtures/check_*.docx` und `EP20187485_*.docx` öffnen → *Vollprüfung*. Die erwartete Anzahl Befunde je Prüfung steht in `<Name>.check.txt` (die Meldung am Ende zeigt die Zahlen je Prüfung).
- `renumber_EN_large.docx` (218 Ansprüche) dient auch als Geschwindigkeitstest.
- `renumber_EN_small.docx`: Cursor in Anspruch 2 → *Abhängigen Anspruch anlegen* (fügt „2a. The method according to claim 2, wherein“ ein); Cursor in Anspruch 2 → *Ansprüche zusammenführen* (übernimmt ihn in Anspruch 1); *Anspruchsbaum* zeigt die Abhängigkeiten.
- `EP20187485_original.docx`: Cursor in Anspruch 1 → *Verfahren ↔ Vorrichtung* (fügt Anspruch 3 „Apparatus for authorizing …, the apparatus being configured to: obtain …“ an); *Ansprüche → Beschreibung*; *Bezugszeichen entfernen* und danach *Bezugszeichen einfügen* (Quelle: dieselbe Datei).
- `versions_tracked_footer.docx` (vorher speichern!) → *Reinschrift erzeugen*: neues Dokument ohne Änderungen/Kommentar, Fußzeile „Reinschrift⇥20 187 485.6“. *Markierte Fassung erzeugen*: Änderungen bleiben sichtbar, Fußzeile „Markierte Fassung⇥20 187 485.6“. Das Original bleibt jeweils unverändert.
- Optional die Unit-Tests: `tests/vba/modTestRenumber.bas` und `modTestJson.bas` zusätzlich importieren, Alt+F8 → `PT_RunRenumberTests` / `PT_RunJsonTests`.

**Ohne Word** (Entwicklung, Linux/CI): `python3 tools/run_vba_tests.py` führt die Text-Logik in LibreOffice Basic aus (Unit-Tests + alle Fixtures).

## Build

`src/` ist die Quelle. `python3 tools/build.py`
- erzeugt `src/modules/modDefaults.bas` aus `config/settings.json` und `config/blacklist.json` (Standardwerte für den ersten Start),
- prüft die VBA-Quellen (nur ASCII, `Option Explicit`, Modulname = Dateiname),
- baut `build/PatentTools.dotm` mit dem Menüband aus `src/customUI/customUI14.xml`.

Die VBA-Module kommen anschließend per Import-Helfer in die .dotm (siehe oben), weil `vbaProject.bin` nur Office selbst zuverlässig schreiben kann. `python3 tools/build.py --update <Datei>.dotm` tauscht nur das Menüband in einer bestehenden .dotm aus und behält den importierten Code.

Testdokumente neu erzeugen: `python3 tools/make_fixtures.py` (benötigt python-docx).
