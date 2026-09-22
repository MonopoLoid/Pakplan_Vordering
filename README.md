# Pakplan Vordering

VBA-driven Excel tool that tracks packing-plan ("Pakplan") progress against
dispatch data pulled from the Paltrack SQL database, and produces the
"Vordering" progress sheet, "Opsomming" summary, and "Grafieke" charts.

## Layout

- `Pakplan Vordering Macro Template.xlsm` — the actual working file. No
  version number in the filename on purpose - that's what git is for. The
  human-readable version shown in the ribbon comes from
  `Module1.APP_VERSION`; bump it and `git tag` the matching commit
  (e.g. `git tag v1.9`) when you push an update.
- `src/` — the VBA source, exported as plain text so it can be diffed and
  reviewed in git. This is a **mirror** of what's inside the `.xlsm`'s VBA
  project, not an independently-built thing. After editing a module inside
  Excel, re-export it here before committing; after editing a file here,
  re-import it into the `.xlsm` before testing.
  - `Module1.bas` — main workbook logic (Setup/Update/Input/Export/Short/Chart).
  - `LocalConfig.bas` — per-machine local config (server, save paths, email
    addresses), ribbon settings storage (`GetAppSetting`/`SetAppSetting`,
    backed by `CustomDocumentProperties`), and the GitHub auto-update
    mechanism (`Git_Update`).
  - `SizeMapping.bas` — carton-size lookup engine.
  - `Sheet*.cls`, `ThisWorkbook.cls`, `UserForm1.frm` — object modules.
- `ribbon/customUI14.xml` — the custom ribbon tab's definition. This is a
  separate part of the `.xlsm`'s zip container, not part of the VBA
  project, so it can't be edited or exported via the VBA editor - it has
  to be edited as text and written back into the `.xlsm`'s zip directly.

## Known in-progress work

Auditing `Module1` for unqualified `Range`/`Cells`/`ActiveSheet`/`Selection`
references. `Update_Stuff`/`Setup_Stuff` show their progress bar with
`UserForm1.Show (False)` (modeless), which means Excel's UI isn't blocked
while they run - if the user clicks into a different open workbook mid-run,
any unqualified reference from that point resolves against whichever
workbook is now active, not this one. Paired with this: a friendly,
central error-handling layer so failures surface as a clear message
instead of a raw VBA runtime error.
