# GridForge v11.1 Feature Matrix and Verification

This matrix defines the tested scope for v11.1. It does not claim a percentage of Excel compatibility. Tests establish GridForge behavior and selected OOXML structure; they do not certify Microsoft Excel rendering.

| Area | Implemented and verified | Partial or not yet verified |
|---|---|---|
| Calculated table columns | Formula entered in a table data cell fills blank cells and formula cells; same-sheet [@Column] formulas evaluate row by row; literal cells remain manual overrides; source edits recalculate; new contiguous rows receive formulas; JSON and XLSX table formula metadata persist; formula edits undo/redo with the workbook. | Table formula grammar remains limited to GridForge's supported structured-reference subset. Structural row insertion and non-grid/bulk mutation pathways do not all trigger expansion. |
| Secondary chart axes | Per-series primary/secondary axis assignment; independent SVG value ranges; left/right value axes and right-axis title; column/line combo; chart model JSON persistence; tested XLSX secondary value-axis and series assignment output/import. | Imported Excel charts with complex/disconnected series, multiple axis groups, custom bounds, and unusual axis crossing may simplify or be skipped. |
| Trendlines | Linear, exponential, logarithmic, polynomial (orders 2–6), power, and moving average; per-series selection; equation/R² flags; forward/backward forecasting; domain and degenerate-data errors; known numerical reference fixtures; chart model persistence; tested OOXML trendline output/import. | Excel's moving-average display/forecast conventions and equation formatting may differ. R² is unavailable for moving averages and constant datasets. No external numerical library is used. |
| Conditional data bars | Live rules over selected ranges; automatic or fixed min/max; solid/gradient style; positive/negative color in GridForge; optional value hiding; multiple rules; rule validation; tested XLSX dataBar output/import and browser save/reopen. | There is no curated preset gallery. XLSX preserves the standard bar color, visibility and limits; separate negative-color and all gradient details are not fully portable. |
| Row/column dimensions | Drag resize, boundary double-click AutoFit, right-click pixel size entry, configured minimum/maximum, wrapped-text-aware approximate row fitting, JSON persistence, XLSX col/row height output and import. | Dimension changes do not join workbook undo/redo history. Resizing multiple selected rows/columns and AutoFit parity with Excel font metrics are not implemented. |
| Chart gallery and editor | Chart Design and Format contextual ribbon tabs appear when a chart is selected; type and style selectors, row/column switching, legend and data-label toggles, Select Data, title and size fields, and delete command are available in the ribbon. Chart selection clears when switching worksheets or deleting the selected chart. The existing detailed editor remains available for per-series axes, combo types, and trendlines. The contextual workflow and XLSX row/column orientation roundtrip are covered by Chromium smoke tests. | Error bars, axis bounds/number formats, minor gridlines, marker/point formatting, rich data-label layouts, and additional style presets remain incomplete. |
| XLSX tables | Native table calculated-column formula metadata is written/read; formulas remain worksheet formulas; tested in browser round-trip and OOXML checks. | Excel table totals, full structured formula syntax, and all calculated-column override edge cases are not certified in Excel. |

## Verification evidence

Completed in the v11.1 release workspace on Windows:

- `npm run regression` — passed all 8 stages after the contextual ribbon update. TypeScript check passed; Node tests passed 912/912; the 10,000-formula and 10,000-cell benchmarks passed; 31 legacy Chromium browser checks passed; all 6 focused v11.1 browser workflows passed; XLSM binary preservation passed; VBA bridge security passed 8 checks.
- `python -m unittest discover -s tests -p "test_desktop*.py"` — 7 tests passed.
- `python scripts/smoke_migration_v110.py` — all 8 migration/browser checks passed.
- `python desktop/gridforge_desktop.py --check` — passed launcher/site presence check.
- `npm run demo:v111` — generated the v11.1 JSON and XLSX examples.

Commands to repeat from the project root:

    npm run regression
    python -m unittest discover -s tests -p "test_desktop*.py"
    python scripts/smoke_migration_v110.py
    python desktop/gridforge_desktop.py --check

The regression runner performs TypeScript checking, Node unit tests, 10,000-formula dependency and 10,000-cell tool benchmarks, Chromium UI regressions, XLSM preservation checks, and VBA bridge protocol/security checks. The v11.1 browser script specifically exercises calculated-column propagation, row/column dragging and AutoFit, data-bar application, secondary-axis/trendline editing, contextual chart ribbon visibility and selection changes, and XLSX save/reopen.

## Microsoft Excel and installer status

The run environment has Chromium and Windows, but no real Microsoft Excel acceptance run is included in this release evidence. PyInstaller built the desktop app and VBA bridge, and Inno Setup 6.7.3 compiled `GridForge-Setup-11.1.0-win64.exe` from the updated package. The installer was silently installed into `%LOCALAPPDATA%\GridForge_v111_smoke`, the installed app's bundled-asset check exited successfully, and the app was uninstalled successfully. This verifies the local install/uninstall workflow, not a clean Windows machine or Microsoft Excel acceptance. The source ZIP contains build scripts rather than the compiled installer; the installer is a separate release file.
