# GridForge v10.8 — Excel-native chart engine

## What changed

GridForge no longer exports charts only as an SVG or GridForge JSON. The **XLSX exporter now writes Excel DrawingML chart objects** (`xl/charts/chartN.xml`), anchored DrawingML objects (`xl/drawings/drawingN.xml`), worksheet/drawing relationships and OOXML content-type entries. Chart values are linked to actual worksheet cell addresses, with a cached display value from the GridForge formula engine. The importer reads supported Excel chart parts back into native, editable GridForge chart definitions without contacting external services.

### Available chart types

Column, bar, line, area, pie, scatter (numeric XY), doughnut, stacked columns and stacked bars. Up to 24 charts per worksheet, up to 250 source rows, 12 columns and 8 series rendered per chart. First text column provides category labels; the first header row provides names. Scatter uses the first numeric series for X and the next numeric series for Y. For category charts, the first numeric column is not treated as the X axis when a textual category column exists.

### Editable controls

Insert > choose a chart; set source range, chart title, legend (bottom/right/off), horizontal and vertical axis captions, or data values. Editing a chart does not modify its source data. Drag/resize embedded charts; use the chart toolbar for editing, SVG export, and deletion. Chart values update after formula edits and save into GridForge JSON or XLSX. The Excel-native cache refreshes on export.

### Interoperability boundaries

**This version writes valid XML and has passed automated round-trip testing in Chromium, but the generated native Excel charts have not been opened and tested on real desktop Microsoft Excel in this environment.** Run `python scripts\verify_excel_charts_windows.py` on Windows with Excel and pywin32 to test the nine sample charts. Check series, axes, labeling and format in Excel manually before distributing a public release.

The importer handles common local, contiguous ranges and standard chart anchors. It does **not** promise lossless round trips for arbitrary Excel objects. Missing: combo charts, secondary axes, pivot charts, stock, waterfall, sunburst, histogram, treemap, trendline editing, external/cross-sheet chart series, theme-perfect colors, 3D charts, complex custom labels, named-table/chart references, advanced formatting, or charts inside `.xlsm` merge-preservation workflows. Some existing `.xlsm` features are handled separately; do not assume new XLSX chart support applies to `.xlsm` round-tripping.

Original worksheet source values remain intact when editing/deleting charts. Import is bounded; drawing relationships with external targets aren't followed. Imported macros/VBA never execute. Imported chart text is XML-escaped when displayed/exported. Chart XML is generated locally without third-party network requests.

## Evidence and useful files

- `src/engine/ooxml_charts.ts`: pure offline OOXML chart + drawing serialization and chart rels.
- `src/engine/xlsx.ts`: XLSX wiring, content types, relationship creation and supported DrawingML import.
- `src/engine/charts.ts`: linked SVG chart preview, new kinds, labels, legends and validation.
- `src/main.ts`, `public/styles.css`: Excel-like chart gallery and dialog with preview and settings.
- `tests/charts_native_v108.test.mjs`: serialization cases, multiple worksheets, invalid input, formula cache updates.
- `scripts/smoke_native_charts_v108.py`: in-memory Chromium creation and 9-chart XLSX import.
- `examples/native_charts_v108.xlsx`: nine native chart objects for real Excel GUI test.
- `scripts/verify_excel_charts_windows.py`: read-only real Excel COM verification. **Not executed** in Linux environment.

## Manual Windows checks

1. Start app (`START_HERE_WINDOWS.bat` / desktop app). Select a populated rectangle and create all nine chart types using Insert.
2. Edit a source formula and verify plotted values change; edit chart caption, legend, axes, labels. Save XLSX and GridForge JSON.
3. Open `examples\native_charts_v108.xlsx` in desktop Excel. Verify nine charts, open the Chart Design ribbon, edit a chart, ensure linked values and data labels are present, save and re-open.
4. Import that XLSX into GridForge; ensure nine charts appear; save it back to XLSX and re-open in Excel.
5. Run `python scripts\verify_excel_charts_windows.py` using your active `pywin32` environment and supply its JSON report for further compatibility fixes.
6. Run the Windows installer build and GUI checks before public distribution.
