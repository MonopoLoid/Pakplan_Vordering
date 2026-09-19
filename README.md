# Pakplan Vordering

VBA-driven Excel tool that tracks packing-plan ("Pakplan") progress against
dispatch data pulled from the Paltrack SQL database, and produces the
"Vordering" progress sheet, "Opsomming" summary, and "Grafieke" charts.

## Layout

- `Pakplan Vordering Macro Template - v*.xlsm` — the actual working file.
- `src/` — the VBA source, exported as plain text so it can be diffed and
  reviewed in git. This is a **mirror** of what's inside the `.xlsm`'s VBA
  project, not an independently-built thing. After editing a module inside
  Excel, re-export it here before committing; after editing a file here,
  re-import it into the `.xlsm` before testing.
  - `Module1.bas` — main workbook logic (Setup/Update/Input/Export/Short/Chart).
  - `Module2.bas` — per-machine local config (server, save paths, email
    addresses) and the GitHub auto-update mechanism.
  - `SizeMapping.bas` — carton-size lookup engine.
  - `Sheet*.cls`, `ThisWorkbook.cls`, `UserForm1.frm` — object modules.

## Known in-progress work

See the working session — current focus is:
1. Securing the GitHub token used by `Module2`'s auto-update.
2. Consolidating the two `Workbook_Open` copies into a real event handler.
3. Moving ribbon/toggle state off the `Data` sheet cells and into
   `CustomDocumentProperties`.
