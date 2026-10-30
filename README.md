# Word Macro (VBA) Project

Versioned source for Microsoft Word macros written in VBA.

## Layout

| Path            | Contents                                              |
|-----------------|-------------------------------------------------------|
| `src/modules/`  | Standard modules exported from the VBA editor (`.bas`) |
| `src/classes/`  | Class modules (`.cls`)                                 |
| `src/forms/`    | UserForms (`.frm` + binary `.frx`)                     |
| `docs/setup/`   | Setup instructions and project goals                  |

## Versioning workflow

Word stores macros inside binary `.docm` / `.dotm` files, which git cannot diff.
Keep the source of truth here as exported text files:

1. In Word, open the VBA editor (`Alt+F11`).
2. Edit your macros.
3. Export each changed module (right-click → *Export File…*) into the matching `src/` folder.
4. Commit the exported files.

To load macros into a document or template, import the files from `src/`
(right-click the project → *Import File…*) and save as `.docm` or `.dotm`.

VBA files must use Windows (CRLF) line endings for import to work.
`.gitattributes` enforces this.
