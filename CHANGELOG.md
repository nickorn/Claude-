# Changelog

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
