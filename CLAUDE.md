# Patent-Tools – Word VBA toolset for patent claims

Word ribbon ("Patent-Tools") + VBA macros for drafting, renumbering and checking patent claims (EN + DE).
Owner/user: Niklas (Patentanwaltskandidat). Later possibly colleagues.

**Read `docs/project-context.md` before any work.** It holds every decision from two Q&A rounds (60 + 14 questions). If anything here conflicts with it, `docs/project-context.md` wins. Update it (and `CHANGELOG.md`) whenever a decision changes.

## Hard constraints
- Target: Microsoft Office LTSC Standard 2024, Word desktop, Windows. VBA, unsigned macros.
- Delivery: ONE global template `build/PatentTools.dotm` for the Word STARTUP folder; IT installs it.
- STRICTLY OFFLINE: no network calls, no web APIs, no telemetry, no online dependencies – confidentiality.
- UI language German (ribbon labels, MsgBox texts, comments, report). Code: English identifiers + English comments.
- `Option Explicit` in every module. One module per feature, prefixed: `modRenumber`, `modCheckAntecedent`, `modFormat`, ...
- Performance: claim sets up to ~300 claims. Use `Range`/`Find`, read paragraphs into arrays, no per-character loops, `Application.ScreenUpdating = False` with guaranteed restore in error handlers.
- Never modify the original document for clean/marked-up versions – always create a new document.
- Settings: JSON in `%APPDATA%\PatentTools\` (`settings.json`, `blacklist.json`, `textblocks\`), created from defaults on first run. Defaults live in `config/`. Parsed by our own minimal parser `src/modules/modJson.bas` (no third-party code). Defaults are embedded via the generated `modDefaults.bas` (`tools/build.py`); never edit it by hand.
- No hard-coded personal paths. Use `Environ("APPDATA")`.

## Document model (what macros can rely on)
- Claims: paragraph starting with `<label>.` + TAB, manual numbering (no Word auto-numbering, no special styles).
- Label grammar: `^\d+[0-9A-Za-z.]*$` (1, 1a, 1.1, 1a.1, 1a1, 12b.3). Labels are opaque tokens – `1a1` ≠ `1a.1`.
- Anything before the first claim (e.g. heading "Claims" / "Patentansprüche" / "Ansprüche") is ignored.
- Sub-features: indent only (no dashes/letters); indent depth = constant at top of module, mirrored in settings.json.
- Between claims: exactly one empty paragraph.
- Formatting everywhere: Arial 11 pt, line spacing 1.5, space before/after 0.
- Renumbering records number/back-reference changes as tracked changes (setting `renumber.trackChangesMode`, default "tracked").
- Footer left: status text by macro, language = claim language. Footer right: application number, typed by hand – NEVER touch it.

## Repository layout
```
CLAUDE.md                 this file
docs/project-context.md   full requirements & decisions (source of truth)
docs/reference/           EP 20 187 485.6 reference material (published, OK to use)
src/modules/*.bas         standard modules (source of truth, imported into .dotm)
src/classes/*.cls         class modules
src/forms/*.frm/.frx      UserForms
src/customUI/             customUI14.xml + images/ (ribbon)
config/                   default settings.json, blacklist.json, textblocks/ (neutral examples only!)
tests/fixtures/           test .docx with seeded errors (EN + DE, small + ~200 claims)
tests/expected/           Niklas' marked-up corrections = expected results
build/                    generated .dotm (not hand-edited)
CHANGELOG.md              version history; version also shown in ribbon "Info"
```

## Build approach
- `src/` is the source of truth. The `.dotm` container + `customUI/customUI14.xml` + ribbon images are assembled by a script (zip/OOXML, e.g. Python) into `build/`.
- `vbaProject.bin` cannot be generated reliably outside Office → VBA modules are imported into the .dotm with the helper `tools/PT_ImportModules.bas` (lives in Normal.dotm), then the .dotm is saved. Steps in README.
- `python3 tools/build.py` regenerates `modDefaults.bas`, checks sources and builds `build/PatentTools.dotm` (ribbon only).
- VBA sources must be pure ASCII with CRLF line endings (VBA editor imports in the ANSI code page; `tools/build.py` checks both): German texts use the tokens of `modStrings.Txt` (`{ue}` …) or `ChrW`.
- Keep Word-independent logic in pure modules (e.g. `modClaimParse`: plain strings in, planned edits out) and the Word object model in thin modules (e.g. `modRenumber`). Pure logic is tested headless with `python3 tools/run_vba_tests.py` (LibreOffice Basic): unit tests in `tests/vba/` plus every `tests/fixtures/*.docx` against its `.expected.txt` (`tools/make_fixtures.py`). Run it before every commit; Word-level behaviour still needs a check in Word by Niklas.
- Ribbon callbacks: `Public Sub OnAction_xyz(control As IRibbonControl)`; dynamic menus via `getContent`.
- `tools/vba_lint.py` (run by the build) is the compile check for code LibreOffice cannot run: undeclared names, calls to Private procedures of other modules, ambiguous public names, ByRef type mismatches, block balance, reserved words as variable names (`name`, `date`, ...), VBA line limits. Add Word/VBA names to its allow-lists only after checking they really exist. The build also validates the ribbon XML attributes and that every callback exists in `modRibbon`.
- Portability notes from LibreOffice runs: avoid typed object arrays (`x() As Collection`, use Variant), always initialise Variant function results, avoid `Replace(..., vbTextCompare)` (use `ReplaceText`), and remember LibreOffice passes ByVal arrays by reference.

## Roadmap (in order)
1. ✅ **MVP-1** Claim renumbering + back-reference update (supplementary numbers, position-based mapping, EN+DE references, ranges).
2. ✅ **MVP-2** Full final claim check → Word comments (author "PatentCheck") + highlight colour per category + "Prüfmarkierungen entfernen" + optional report document.
3. ✅ Formatting fixer.
4. ✅ Clean version / marked-up version + footer status.
5. ✅ Text blocks (dynamic menu from files `<Subfolder>\NN_Label.docx`, prefix = sort order only; placeholders `[DATUM]`, `[D1]`, ...).
6. ✅ Claim tree, claim merge, dependent-claim skeletons.
7. ✅ Method ↔ apparatus, claims → description, reference signs strip/insert.

## Open points (ask Niklas before deciding)
- none currently

## Working rules for Claude
- Deliver complete, importable modules – never fragments.
- German UI strings centralised (e.g. `modStrings`) so wording can be changed in one place.
- Every check is switchable in settings.json.
- Ship NO firm content in text blocks – confidential; only neutral examples.
- Test against `tests/fixtures` and `docs/reference` (EP 20 187 485.6) before claiming a feature works.
