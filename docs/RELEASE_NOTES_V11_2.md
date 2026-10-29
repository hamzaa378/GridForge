# GridForge 11.2.0 Release Notes

## Excel 2019-style workspace

- Added the familiar File, Home, Insert, Page Layout, Formulas, Data, Review and View tab order, followed by Developer tools.
- Table Design appears while the active cell is inside a formatted table. Chart Design and Format appear while a chart is selected.
- Added standard ribbon group captions and Excel-style whole-row and whole-column selection.
- Selected rows and columns can be resized together. Page Layout includes direct row-height and column-width commands.

## Charts and formatting

- Added error bars per chart series: fixed value, percentage and standard deviation. They render in GridForge and round-trip through supported native Excel chart XML.
- Added Excel-inspired data-bar color presets alongside custom bar colors.
- Preserved the earlier additions: secondary value axes, chart bounds, value number formats, major/minor gridlines, markers, per-series colors, and six trendline types.
- Corrected chart theme handling when a theme is changed during chart creation, so the series color defaults follow the selected palette.

## Verification

- `npm run regression`: 8/8 stages passed; 914 Node tests passed; 31 general Chromium checks and focused chart/table workflows passed; formula and Data Studio stress benchmarks passed; XLSM and macro-bridge security checks passed.
- `python -m unittest discover -s tests -p "test_desktop*.py"`: 7 passed.
- `python scripts/smoke_migration_v110.py`: all 8 migration browser checks passed.
- `python desktop/gridforge_desktop.py --check`: passed.
- `python scripts/smoke_tables_charts_v109.py`: passed table/chart XLSX round-trip and theme checks.
- Windows installer compiled; isolated silent install, bundled frontend presence check and uninstall passed.
- Real Excel chart acceptance remains unperformed; browser and OOXML checks do not certify Excel rendering.
