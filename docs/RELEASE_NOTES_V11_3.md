# GridForge 11.3.0 Release Notes

## Standalone Excel-style workspace

- GridForge opens a new installation to a blank `Book1` with A1 selected. The demo workbook remains available from File → Example.
- Added Quick Access New, Undo, Redo and Save controls, moved command search into the title area, and updated formula-bar cancel/apply controls.
- Reorganized the Home ribbon into clipboard, font, alignment, number, styles and cells groups. The Help tab opens keyboard help.
- Tuned the ribbon, sheet tabs, grid headers and status bar to match common Excel 2019 layout proportions. GridForge remains fully standalone.

## Chart editing

- Added top-positioned chart legends alongside bottom, right and hidden positions.
- Added a Legend Position control to Chart Design, matching the chart editor setting.
- Added SVG rendering and supported native XLSX import/export for top legends.

## Verification

- Browser regression: 32 checks passed, including the blank-workbook start, Quick Access, ribbon, workbook interactions and chart workflows.
- Full regression passed 8/8 stages, 915 Node tests, 32 browser checks, the focused chart/table workflows, and 7 desktop tests. The Windows installer compiled, installed to an isolated folder, had its bundled assets verified, and uninstalled successfully.

## Known limitations

- GridForge implements its own grid, formulas, rendering and file behavior. It does not provide the full Microsoft Excel feature set or pixel-perfect UI.
- Drawing tools, PivotTables, PivotCharts and advanced Excel chart configurations remain outside this release.
- Native XLSX interoperability covers GridForge's documented subset; verify complex Office files in their originating application.
- The installer was verified on the current Windows build machine; clean-machine and real Excel acceptance checks remain outstanding.

---

The v11.2 notes are preserved in `RELEASE_NOTES_V11_2.md`.

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
