# Ready-to-import GridForge Macro Studio examples

`cleanup_contacts.json` processes columns B and C, then compacts duplicate rows based on the first column. Column indices in `keyColumns` are **zero-based**: A=0, B=1, C=2. Removing duplicate rows changes selected-cell contents and may alter relative formula references, so preview first. This is range compaction, not a worksheet-wide Delete Rows command.

`parameterized_tax.json` demonstrates parameters in formula text and conditional-format thresholds. `parameter_values.json` provides overrides. Replacing tokens does not execute JavaScript or VBA.

In the app, choose Automate → Import Macro JSON → Preview effects → Run selected macro. In the Windows installer, the Start menu also contains **GridForge Macro Studio**. To use Node's developer CLI, see `docs/MACRO_STUDIO_V4.md`.
