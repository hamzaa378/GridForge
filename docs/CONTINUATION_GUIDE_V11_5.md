# GridForge v11.5 — Continuation Guide

GridForge is a standalone spreadsheet application with its own worksheet grid and workbook engine. Excel 2019 is an interaction reference only. The desktop application does not embed or require Microsoft Excel.

## Implemented in v11.5

- Added a dedicated Formula ribbon with a searchable Insert Function dialog and categorized function menus backed by the registered formula catalog.
- Added SUM, AVERAGE, COUNT, MAX and MIN insertion for the selected numeric range or the contiguous numeric cells above the active cell.
- Added direct precedent/dependent highlighting, a clear-trace control, a worksheet Show Formulas toggle, and a formula inspection dialog with expression tree, references, input values and current result.
- Added focused Chromium regression coverage for the Formula ribbon interactions.

## Development and release

Run `npm run build`, `npm run regression`, and `python -m unittest discover -s tests -p "test_desktop*.py"`. On Windows, build the portable app and installer with `desktop/build_windows.ps1` or `BUILD_WINDOWS_INSTALLER.bat`.

## Remaining parity work

Named ranges and a full Name Manager, Excel calculation modes, stepwise formula evaluation, PivotTables, and full Excel compatibility are not included in v11.5. Keep new format behavior tied to import/export verification and document unsupported Excel features rather than implying full parity.
