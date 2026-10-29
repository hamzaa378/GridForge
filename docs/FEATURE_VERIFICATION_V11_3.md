# GridForge v11.3 Feature Matrix and Verification

This matrix describes tested GridForge behavior. It does not claim full Excel compatibility or certify Microsoft Excel rendering.

| Area | Implemented and verified | Known limits |
|---|---|---|
| Excel-style workspace | Blank `Book1`/A1 first-run; title-area search; Quick Access New, Undo, Redo and Save; File/Home/Insert/Page Layout/Formulas/Data/Review/View/Help/Developer tabs; grouped Home commands; formula-bar cancel/apply; worksheet tabs and status line. Fresh-session UI regression passed. | GridForge uses its own worksheet grid and engine. Drawing tools, PivotTables, and some Excel ribbon commands are not implemented. |
| Row and column selection | Header click selects a full row or column; Shift extends selection; resizing selected adjacent rows/columns applies a shared drag delta; Page Layout offers direct selected height/width entry; existing AutoFit, persistence and XLSX sizing remain. | Dimension changes are not yet in workbook undo/redo. AutoFit is approximate and uses browser font measurement. |
| Data-bar presets | Eight Excel-inspired colors, custom color, solid/gradient style, negative color, optional hidden values and automatic/fixed bounds. | Presets are color choices, not the full Excel gallery. OOXML output preserves standard positive color, visibility and min/max; negative-color and all gradient variants are not portable. |
| Error bars | Per-series fixed, percentage and standard-deviation settings; SVG whiskers; validation; OOXML export/import; tested fixed-value rendering and percentage OOXML output. | GridForge's standard-deviation method and rendered caps may differ from Excel. Real Excel acceptance has not been performed. |
| Chart axes and series | Per-series primary/secondary axis, independent value ranges, axis titles, primary-axis bounds and number format, major/minor gridlines, column/line combo, custom colors and markers. | Imported complex/multi-group Excel charts may simplify. Primary and secondary axes share one displayed number-format choice. |
| Chart legends | Bottom, top, right and hidden placement; the top option is rendered in SVG and represented in supported native XLSX XML. | Left-side legends are not supported. Pie charts use a compact adapted legend layout. |
| Trendlines | Linear, exponential, logarithmic, polynomial, power and moving-average fits; forecasts, equation/R² flags, known-reference numerical tests and supported XLSX metadata. | Excel convention and equation display can differ; moving-average R² is unavailable. |
| Calculated table columns | Structured formula propagation, override tracking, new-row fill, recalculation, undo/redo and JSON/XLSX table metadata are covered by tests. | Structured references remain a supported subset; not all structural row insertion or mutation paths trigger growth. |
| Interoperability | OOXML structure and GridForge XLSX save/reopen tested; existing 13 chart types remain supported. | No Microsoft Excel open/save acceptance run is included. Unsupported Excel chart types/layouts, PivotCharts and some conditional-formatting details can simplify or be omitted. |

## Verification evidence (2026-10-24)

- `npm run regression` — 8/8 stages passed. TypeScript check passed; 915/915 Node tests passed; 10,000-formula and 10,000-cell Data Studio performance checks passed; 32 general Chromium checks and the focused v11.3 chart/table workflows passed; XLSM preservation passed; VBA bridge security passed 8 checks.
- `python -m unittest discover -s tests -p "test_desktop*.py"` — 7 tests passed.
- `python scripts/smoke_migration_v110.py` — 8 browser checks passed.
- `python scripts/smoke_tables_charts_v109.py` — native table/chart import/export and theme checks passed.
- `python desktop/gridforge_desktop.py --check` — passed.
- `desktop/build_windows.ps1` — PyInstaller desktop and VBA bridge built, followed by successful Inno Setup 11.3.0 compilation. The installer was silently installed to an isolated folder, its launcher, frontend assets, bridge and guide verified, then silently uninstalled.

The v11.3 browser regression used Chromium and inspected supported OOXML. It did not launch Microsoft Excel. The installer was checked on this Windows 11 build machine; clean-machine acceptance and real Excel open/save checks remain outstanding.
