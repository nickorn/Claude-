# Changelog

## [0.4.0] – 2026-10-31
### Added
- MVP-3 Format group: "Formatierung anwenden" (font, size, line spacing, space before/after, sub-feature indents as whole multiples of `subFeatureIndentCm`, exactly one empty paragraph between claims), "Doppelte Leerzeichen" (also spaces before punctuation / at paragraph end), "Zeichensetzung" (";" / ":" / "; and"/"; und" / final full stop), "Bezugszeichen-Format" ("(150; 450)" → "(150, 450)", space before "(").
- Settings `format.fixEmptyParagraphs`, `format.layoutTrackChanges` (default "silent"), `format.textTrackChanges` (default "tracked").
- Shared Word helpers `BeginDocWork`/`EndDocWork`, `ApplyTextOps` (replace + insert, verified); build checks for declaration order and Sub/Function block balance.

## [0.3.0] – 2026-10-31
### Added
- MVP-2 "Vollprüfung": numbering, punctuation (";", "; and"/"; und" before the last feature incl. nested lists, full stop, text after the full stop), spacing (empty paragraphs between claims, double spaces, space before punctuation), formatting (font, size, line spacing, paragraph spacing), back-references (unknown, self, forward), independent claims with same-category references, multiple dependencies (multiple-dependent on multiple-dependent; every one optional), antecedent basis EN/DE along the dependency chain incl. singular/plural, reference signs (same term = same sign, same sign = same term, missing at first mention per claim, format ";"/", "), blacklist, categories (dependent claim of another category, second independent claim of a category).
- Findings as PatentCheck comments with highlight colour per check; "Einzelprüfung" menu; "Prüfmarkierungen entfernen" (only tool comments/highlights); "Prüfbericht" as new document.
- Settings `check.createReport`, `check.flagEveryMultipleDependency`, `check.maxFormattingFindings`.
- Check fixtures with expected counts (EN/DE seeded errors, EP 20 187 485.6 original + amended); unit tests incl. the reference table.

## [0.2.0] – 2026-10-30
### Added
- MVP-1 "Neu nummerieren": renumbers claims in document order (supplementary labels 1a, 1.1, 1a.1, 1a1 …), updates EN/DE back-references incl. lists, ranges ("to", "bis", "-"), compounds ("Patentanspruch") and inflections ("Anspruchs", "Ansprüchen"). Respects existing Track Changes; aborts on duplicate labels; flags unknown references and changed ranges as PatentCheck comments; one undo step.
- Ribbon tab "Patent-Tools" (groups Ansprüche, Einstellungen), settings in `%APPDATA%\PatentTools\` created on first run, own JSON parser (`modJson`).
- Build script `tools/build.py`, import helper `tools/PT_ImportModules.bas`, fixture generator `tools/make_fixtures.py`, headless test runner `tools/run_vba_tests.py` (LibreOffice Basic).
- Test fixtures (EN/DE small, tracked changes, 218 claims) with expected results; VBA unit tests.

### Changed
- R2-2 decided: renumbering records changes as tracked changes (`renumber.trackChangesMode` default "tracked").
- Text-block naming scheme `<Subfolder>\NN_Label.docx` confirmed.

## [0.1.0] – 2026-10-30
### Added
- Project context from Q&A rounds 1 (Q1–60) and 2 (R2-1–14).
- CLAUDE.md, repository layout, default `settings.json` and `blacklist.json` seed (EN + DE).
- Neutral example text blocks (EPA, DPMA, PCT_EESR, Allgemein).
- Reference material EP 20 187 485.6 (claims, figures, description) with expected check findings.
