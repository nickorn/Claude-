# Project context: Word patent claim toolset (ribbon + VBA macros)

Status: living document. Round 1 (Q1-60) and round 2 (R2-1 to R2-14) answered on 30.10.2026. Open points are marked [OPEN] and collected at the end.

## 1. Users and scope
- Primary user: Niklas, patent attorney trainee (Patentanwaltskandidat). The tool must work equally well for qualified patent attorneys.
- Starts as a single user; colleagues may follow after testing. Design for multi-user from day one: no hard-coded personal paths, one settings file per user.
- Jurisdictions: all, with priority on EP, DE and US.
- Claim languages: English and German. Every check has to recognise both.
- Document types: new drafts, responses to office actions, amended claim sets.
- Claim set size: 25 to ~200-300 claims, typically ~60. Performance matters, so use Range/Find and arrays and avoid per-character loops.
- Reference material: EP 20 187 485.6 is published and fine to use. Niklas will upload the remaining pages.

## 2. Technical environment and delivery
- Microsoft Office LTSC Standard 2024 (Word desktop, Windows).
- VBA is allowed; macros are unsigned.
- IT does the installs. Delivery is ONE global template (.dotm) with customUI ribbon XML, placed in the Word STARTUP folder.
- Build approach: the source of truth is plain-text modules (.bas/.cls/.frm) plus customUI XML. Claude builds the .dotm container with the ribbon. The VBA modules are imported into the .dotm in Word (helper macro `PT_ImportModules`), because vbaProject.bin can only be created reliably by Office itself.
- There are no existing firm macros or templates to coexist with.
- DMS: PatOrg, with no naming rules.
- STRICTLY OFFLINE: no network or API calls of any kind (confidentiality).
- UI language: German (ribbon labels, messages, report texts).
- Code language: English identifiers and English comments; `Option Explicit` everywhere; one module per feature with a prefix (e.g. `modRenumber`, `modCheckAntecedent`).
- Settings: JSON files in `%APPDATA%\PatentTools\` (`settings.json`, `blacklist.json`, `textblocks\`). Parsed with an embedded VBA JSON parser (own code `modJson`, offline). Defaults are created on first run.
- Versioning: Git-style source folder and a CHANGELOG; the version number is shown in the ribbon's "Info" button.
- Icons: built-in imageMso where an icon clearly describes the function, otherwise custom PNG icons embedded in the .dotm.

## 3. Document conventions (hard rules)
- Claim numbering: manual "1." + TAB, no Word auto-numbering, so documents survive hand-over to inventors.
- No special claim styles. Claim 1 is independent, followed by dependents and possibly further independent claims.
- Sub-steps are indented only: no dashes, no letters. The indent depth is a configurable constant at the top of the macro code (mirrored in settings.json).
- Punctuation: each feature ends with ";", the last one with ".", and the claim ends with exactly one full stop. "; and" before the last feature is enforced (Q24 = yes) but can be switched off in settings, because practice varies.
- Formatting ALWAYS: Arial 11 pt, line spacing 1.5, space before/after = 0.
- No double or multiple spaces.
- Margin line numbering: not needed.
- Footer: bottom left = document status, set by macro, language follows claim language: DE "Markierte Fassung" / "Reinschrift" / "Vom Erfinder freigegeben"; EN "Marked-up version" / "Clean version" / "Approved by inventor". Bottom right = application number, entered BY HAND (the macro never touches the right footer content).
- Documents usually contain ONLY claims (possibly with a "Claims"/"Patentansprüche"/"Ansprüche" heading). Claims are detected by the pattern "number + . + TAB" at paragraph start; anything before the first claim is ignored.
- Formatting rules (Arial 11, 1.5, 0/0 spacing) apply to the WHOLE document.
- Between two claims: exactly one empty paragraph. Inside a claim: empty paragraphs are NOT checked.
- Reference signs: configurable per attorney, either (a) without parentheses or (b) in parentheses. Multiple signs are separated by ", " (e.g. "(150, 450)"). The macro can normalise ";" to ",".

## 4. Supplementary claim numbering (drafting)
- During drafting, inserted claims get supplementary numbers such as 1a, 1.1, 1a.1, 1a1.
- Supplementary claims can sit ANYWHERE in the document. The new order is the order of the claims IN THE DOCUMENT (paragraph position), not a sort by number. The macro first reads all claim numbers, builds an old->new map by position, then rewrites.
- 1a1 and 1a.1 are DIFFERENT, distinct numbers (treat every label as an opaque string token; never normalise one into the other).
- Accepted label grammar: digits followed by any mix of letters, digits and dots (e.g. 1, 1a, 1.1, 1a.1, 1a1, 12b.3). Duplicate labels are an error: abort and report, do not guess.
- The renumbering macro converts these into a clean consecutive integer sequence and updates every back-reference ("claim 1a", "claims 1.1 to 3", "any one of claims 1a or 2", "Anspruch 1a", "einem der Ansprüche 1 bis 3").
- Ranges ("claims 1 to 3", "Ansprüche 1a bis 4") are resolved by document position after renumbering; if a range becomes non-contiguous or unclear, it is rewritten as far as safe and flagged with a comment.
- Track Changes (R2-2, decided 30.10.2026): number and back-reference changes are RECORDED AS TRACKED CHANGES (`renumber.trackChangesMode = "tracked"`). The setting can be switched to "silent" in settings.json.

## 5. Final claim check (wanted)
- Claim numbering: consecutive, no gaps or duplicates, no supplementary numbers left over.
- Missing full stop at the end of a claim; ";" / "; and" / "." rules for features.
- Spacing between claims, double spaces, font and spacing compliance.
- Independent claims contain no back-reference (independent claims are detected automatically). EXCEPTION, allowed and not flagged: claims of another category that refer to a claim of a different category, e.g. "System/Device configured to carry out the method of claim 1", "Computer program ... according to claim 1" (used for system claims).
- A "wherein"/"wobei" clause after the claim's final full stop is flagged (as in the reference claim 1).
- Antecedent check works in EN and DE. DE: ein/eine/einen/eines/einem/einer (+ mindestens ein, eine Mehrzahl von, mehrere) introduce; der/die/das/den/dem/des/dieser... refer back. Approach "flag it, you decide": false positives are acceptable (agreed).
- Back-references: to non-existent claims, to later claims (forward reference), to itself.
- Multiple dependencies, and multiple-dependent on multiple-dependent (US).
- Antecedent basis: a/an/at least one/one or more/a plurality of introduce a term; the/said refer back. The preamble counts as an introduction. Singular/plural mismatches are flagged. The check follows the dependency chain into parent claims.
- Reference signs: same term = same sign, same sign = same term, missing signs. Figures are NOT checked.
- Claim category detection (method, device, system, computer program, storage medium) and correspondence between independent claims.
- Blacklist of unclear terms: Claude supplies the seed list (EN + DE) in blacklist.json; Niklas amends it later.
- NOT wanted: fee or claim-count warnings, term-consistency or near-duplicate checks, a support check against the description, feature-analysis tables, a DE/EN glossary.
- Implementation decisions (Claude, 31.10.2026 – Niklas may overrule): a new check first removes the previous PatentCheck marks; plain multiple dependency is only flagged with `check.flagEveryMultipleDependency` (default off), multiple-dependent on multiple-dependent always; a signed term must carry its sign at its first mention in every claim; independent claim = no reference in the preamble (text before the first ":", "," or "comprising"/"wobei"/...); German phrases end at the capitalised noun; formatting findings are capped (`check.maxFormattingFindings`, default 20) plus one summary note.
- Findings output: Word comments under a fixed author name ("PatentCheck") plus coloured highlighting, one colour per check category. A ribbon button "Prüfmarkierungen entfernen" removes only these. An optional summary report opens as a new document (table: claim no., check, finding, text excerpt).

## 6. Amendments and prosecution
- Amendments are marked with Word Track Changes.
- One-click generation of the clean version and the marked-up version, each with the matching footer status. The marked-up version stays a Word document with Track Changes mark-up (no conversion to manual underline/strikethrough). The clean version = copy with all changes accepted and comments removed. Both are created as NEW documents; the original is never changed.
- Automatic renumbering with back-reference update; claim tree (dependency view).
- Merge dependent claim(s) into an independent claim, adjusting antecedents ("a" -> "the" where the term is already introduced) and renumbering the rest.
- Art. 123(2) basis table: NOT wanted.
- No translation features (no DE/EN glossary).

## 6a. Text building blocks (Textbausteine)
- Firm texts are confidential, so Claude ships NO firm content, only neutral example blocks as templates; Niklas fills in his own.
- Blocks are read from files, never hard-coded: folder `%APPDATA%\PatentTools\textblocks\` (path changeable in settings.json, e.g. later a shared network drive for colleagues).
- One .docx per block (keeps formatting). Subfolders become ribbon submenus: `EPA\`, `DPMA\`, `PCT_EESR\`, `Allgemein\`.
- File name = menu label; optional numeric prefix for sorting, e.g. `10_Einleitung Art. 94(3) EPÜ.docx` -> shown as "Einleitung Art. 94(3) EPÜ". Naming scheme confirmed by Niklas on 30.10.2026.
- The menu is built dynamically (getContent callback) every time it opens, so new files appear without restarting Word.
- Placeholders in square brackets ([DATUM], [D1], [MERKMAL], [ANSPRUCH]) are optionally asked for in an input dialog on insert; unfilled ones are highlighted.
- Claude provides example blocks (EPO: Art. 94(3) opening, main/auxiliary requests, oral proceedings as precaution, Art. 84, Art. 54, Art. 56 problem-solution skeleton, adaptation of description; DPMA: Eingabe zum Prüfungsbescheid, § 3/§ 4 PatG, Anhörung, Haupt-/Hilfsantrag; PCT/EESR: Rule 70a/161 EPC, Art. 19/34 PCT, WO-ISA reply).

## 7. Drafting helpers (wanted)
- Generate an apparatus claim from a method claim, and vice versa.
- Generate dependent-claim skeletons ("The method according to claim X, wherein" / "Verfahren nach Anspruch X, wobei").
- Strip or insert reference signs (e.g. for US filings).
- Turn claims into description paragraphs, with selectable intro phrases ("According to a first embodiment, ...", "According to an embodiment, ...", "In examples, ...", "Gemäß einer ersten Ausführungsform ...").
- Insertable building blocks / boilerplate.

## 8. Ribbon (decision)
Tab "Patent-Tools" with these groups:
1. Ansprüche: Neu nummerieren, Anspruchsbaum, Ansprüche zusammenführen, Abhängigen Anspruch anlegen
2. Prüfung: Vollprüfung, single checks (dropdown), Prüfmarkierungen entfernen, Prüfbericht
3. Format: Formatierung anwenden, Doppelte Leerzeichen, Zeichensetzung, Bezugszeichen-Format
4. Änderungen: Reinschrift erzeugen, Markierte Fassung erzeugen, Fußzeile/Status
5. Entwurf: Verfahren <-> Vorrichtung, Ansprüche -> Beschreibung, Bezugszeichen entfernen/einfügen
6. Textbausteine: EPA, DPMA, PCT/EESR (dynamic menu read from textblocks folder)
7. Einstellungen: Einstellungen öffnen, Blacklist bearbeiten, Info/Version

## 9. Priorities / roadmap
1. MVP-1: claim renumbering + back-reference update (including supplementary numbers).
2. MVP-2: full final claim check (numbering, punctuation, spacing, formatting, back-references, independent-claim check, antecedent basis, reference signs, blacklist, categories) with comments, highlighting, cleanup and report.
3. Formatting fixer (Arial 11 / 1.5 / no spacing / double spaces / punctuation / indent).
4. Clean and marked-up version + footer status.
5. Standard text sections (EPO/DPMA/PCT/EESR).
6. Claim tree, claim merge, dependent-claim skeletons.
7. Method <-> apparatus conversion, claims -> description, reference-sign strip/insert.

## 10. Testing
- Claude builds test documents with seeded errors (EN + DE, small and ~200 claims for performance).
- Niklas returns them with his own fixes as marked-up versions; these serve as expected results.

- The remaining pages of EP 20 187 485.6 (once uploaded) serve as the first real-world test case; Claude decides what the checks should catch there.

## 11. Open points [OPEN]
- Remaining pages of the reference application to be uploaded.

## Reference material in project
- claude/reference-claims-biometric-authorization.md (EP 20 187 485.6, page 1 of claims; remaining pages to be uploaded).
