# Patent-Tools (Word)

Ribbon-Tab „Patent-Tools“ mit VBA-Makros für Patentansprüche (DE/EN): Neunummerierung, Anspruchsprüfung, Formatierung, Reinschrift/markierte Fassung, Textbausteine, Entwurfshelfer.

- Anforderungen & Entscheidungen: `docs/project-context.md`
- Arbeitsanweisungen für Claude Code: `CLAUDE.md`
- Referenzmaterial: `docs/reference/` (EP 20 187 485.6)

## Mit Claude Code starten
```
cd patent-tools
git init && git add . && git commit -m "Initial project context"
claude
```
Erster Prompt z. B.: „Lies CLAUDE.md und docs/project-context.md und setze MVP-1 (Neunummerierung) um.“

## Installation (später, durch IT)
1. `build/PatentTools.dotm` in den Word-STARTUP-Ordner kopieren (`%APPDATA%\Microsoft\Word\STARTUP`).
2. Beim ersten Start legt das Makro `%APPDATA%\PatentTools\` mit `settings.json`, `blacklist.json` und `textblocks\` an (Vorlagen aus `config/`).

## Build
`src/` ist die Quelle. Ein Build-Skript erzeugt den .dotm-Container mit Ribbon-XML; die VBA-Module werden einmalig im VBA-Editor importiert (Alt+F11 → Datei importieren) und die .dotm gespeichert.
