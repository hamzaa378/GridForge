# GridForge v11.6 — Continuation Guide

GridForge is a standalone spreadsheet application with its own worksheet grid and workbook engine. Excel 2019 is an interaction reference only. The desktop application does not embed or require Microsoft Excel.

## Implemented in v11.6

- Added worksheet-scoped defined names for single cells and rectangular ranges.
- Added Define Name, Name Manager, Use in Formula, and Create from Selection controls to the Formula ribbon.
- Create from Selection reads non-empty headers in a selected row or column, generates unique legal names, and targets the data cells below or beside each header.
- Named references are resolved by the formula parser, included in dependency tracking, and recalculated when a named definition changes.
- Name changes are undoable/redoable and persist in GridForge JSON.
- XLSX import/export and macro-preserving XLSM edits read/write the supported sheet-scoped `definedNames` subset.

## Development and release

Run `npm install`, `npm run build`, `npm run regression`, and `python -m unittest discover -s tests -p "test_desktop*.py"`. On Windows, build the portable app and installer with `desktop/build_windows.ps1` or `BUILD_WINDOWS_INSTALLER.bat`.

## Remaining parity work

Workbook-scoped names, names that resolve to constants or arbitrary formulas, calculation mode controls, stepwise formula evaluation, PivotTables, and full Excel compatibility are not included. Keep new behavior tied to import/export verification and document unsupported Excel features rather than implying full parity.
