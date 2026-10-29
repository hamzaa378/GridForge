# GridForge v11.2 Feature Matrix and Verification

This matrix describes tested GridForge behavior. It does not claim full Excel compatibility or certify Microsoft Excel rendering.

| Area | Implemented and verified | Known limits |
|---|---|---|
| Excel-style workspace | Standard tab order including Page Layout and Developer; contextual Table Design and Chart Design/Format; contextual visibility smoke-tested; ribbon controls retain existing keyboard navigation and collapse behavior. | GridForge still has its own workbook engine and drawing surface; this is an Excel-inspired interface, not Microsoft Excel itself. A few advanced settings remain in focused dialogs. |
| Row and column selection | Header click selects a full row or column; Shift extends selection; resizing selected adjacent rows/columns applies a shared drag delta; Page Layout offers direct selected height/width entry; existing AutoFit, persistence and XLSX sizing remain. | Dimension changes are not yet in workbook undo/redo. AutoFit is approximate and uses browser font measurement. |
| Data-bar presets | Eight Excel-inspired colors, custom color, solid/gradient style, negative color, optional hidden values and automatic/fixed bounds. | Presets are color choices, not the full Excel gallery. OOXML output preserves standard positive color, visibility and min/max; negative-color and all gradient variants are not portable. |
| Error bars | Per-series fixed, percentage and standard-deviation settings; SVG whiskers; validation; OOXML export/import; tested fixed-value rendering and percentage OOXML output. | GridForge's standard-deviation method and rendered caps may differ from Excel. Real Excel acceptance has not been performed. |
| Chart axes and series | Per-series primary/secondary axis, independent value ranges, axis titles, primary-axis bounds and number format, major/minor gridlines, column/line combo, custom colors and markers. | Imported complex/multi-group Excel charts may simplify. Primary and secondary axes share one displayed number-format choice. |
| Trendlines | Linear, exponential, logarithmic, polynomial, power and moving-average fits; forecasts, equation/R² flags, known-reference numerical tests and supported XLSX metadata. | Excel convention and equation display can differ; moving-average R² is unavailable. |
| Calculated table columns | Structured formula propagation, override tracking, new-row fill, recalculation, undo/redo and JSON/XLSX table metadata are covered by tests. | Structured references remain a supported subset; not all structural row insertion or mutation paths trigger growth. |
| Interoperability | OOXML structure and GridForge XLSX save/reopen tested; existing 13 chart types remain supported. | No Microsoft Excel open/save acceptance run is included. Unsupported Excel chart types/layouts, PivotCharts and some conditional-formatting details can simplify or be omitted. |

## Verification evidence (2026-10-24)

- `npm run regression` — 8/8 stages passed. TypeScript check passed; 914/914 Node tests passed; 10,000-formula and 10,000-cell Data Studio performance checks passed; 31 general Chromium checks and the focused v11.2 workflows passed; XLSM preservation passed; VBA bridge security passed 8 checks.
- `python -m unittest discover -s tests -p "test_desktop*.py"` — 7 tests passed.
- `python scripts/smoke_migration_v110.py` — 8 browser checks passed.
- `python scripts/smoke_tables_charts_v109.py` — native table/chart import/export and theme checks passed.
- `python desktop/gridforge_desktop.py --check` — passed.

The regression run used Chromium and inspected supported OOXML. It did not launch Microsoft Excel. The v11.2 installer compiled successfully with Inno Setup, was silently installed into an isolated workspace folder, its bundled frontend files were verified, and it was silently uninstalled. This does not replace clean-machine or Microsoft Excel acceptance testing.
