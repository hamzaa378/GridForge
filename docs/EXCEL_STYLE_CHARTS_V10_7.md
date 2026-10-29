> **Historical v10.7 document.** The statement below that native XLSX charts are unsupported is superseded by GridForge v10.8. See `EXCEL_NATIVE_CHARTS_V10_8.md`.

# GridForge v10.7 — Excel-style charts and formatting

## What changed

The primary editing surface now resembles the familiar Office spreadsheet workflow: Home contains text formatting and number formatting; Insert contains a **six-type chart gallery**; File contains workbook import/export; Data contains sorting, filtering, conditional formatting and bulk tools; Automate contains the offline Macro Studio. Calculation Insights remains opt-in. Ctrl+F1 collapses the ribbon, and Excel familiar / Guided workspace choices remain available.

## Create and edit charts

1. Select a rectangular range containing text labels and numeric values (for example `A3:D11` on the built-in sample). Formula results are used, not raw formula text.
2. Select **Insert** and choose **Column**, **Bar**, **Line**, **Area**, **Pie**, or **Scatter**. A preview dialog opens with the selection and editable title. If only one cell was selected on the sample worksheet, the range is suggested automatically.
3. Insert. The chart appears over the worksheet and refreshes whenever its source cells change. Drag the **chart header** to reposition; drag the bottom-right grip to resize.
4. Use the chart header's edit, SVG download, or delete controls. **Manage charts** lists charts for the active sheet. Up to 24 charts per worksheet, 250 data rows, 12 columns and 8 rendered series per chart are supported.
5. Preserve charts by exporting a **GridForge JSON** workbook. They are automatically stored in the application's local draft too.

Data interpretation: text labels in the first column, header names in the first row, adjacent numerical columns as independent series. Missing/error values leave gaps; they are not silently treated as zeros. Pie charts use the first data series, accept positive slices only and limit the legend to 8 slices. Scatter plots use numeric X and Y columns. Charts handle negative and zero data. Data is not sent to any service.

## Home tab formatting

- Font families: Aptos, Calibri, Arial, Segoe UI, Georgia and Consolas.
- Font sizes 8–24, bold, italic, underline, strikethrough, text and fill colors.
- Text wrapping, horizontal alignment, General/Number/Integer/Currency/Percent formats.
- All or bottom cell borders, predefined Header/Title/Total/Accent/Attention styles, and clear formatting.
- Undo and redo for format changes; formatting is preserved in GridForge JSON. XLSX export supports these font, wrap and simple border settings. Unknown imported font families and nontrivial Excel borders are not fully round-tripped.

## Compatibility limits and honest release claims

This implementation is native and offline. It does **not** yet provide Excel's full Chart Design/Format capabilities such as trendlines, secondary axes, waterfall/stock charts, data labels, pivot charts, 3D, or all chart layout settings. **GridForge charts are not serialized as native DrawingML chart objects into exported XLSX/XLSM workbooks** and imported Excel chart objects are not preserved as editable GridForge chart definitions. Download individual charts as SVG, and save GridForge JSON to preserve editable charts. XLSM exports follow their existing limited-edit preservation rules; do not assume Excel workbook chart preservation or VBA execution in GridForge.

Cell dimensions remain a fixed virtualization grid (not free-width Excel columns). Wrapped text may be visually clipped by the standard row height; the formula bar retains complete content. The engine still has 506 locally implemented function names and 10 connection-dependent names. Full Excel formula, VBA and workbook parity is not claimed.

## Technical details

- `src/engine/charts.ts`: validated range extraction, SVG renderer and chart metadata validator; no network or third-party chart API.
- `src/engine/sheettools.ts`: per-sheet chart collections with strict limits and import validation.
- `src/main.ts`: gallery, dialogs, live worksheet charts, drag/resize, source-linked editing and download.
- `src/engine/format.ts` / `src/engine/workbook.ts`: validated style properties, undoable patches.
- `src/engine/xlsx.ts`: cell style serialization for font family/size, wrap, strikethrough, basic borders; NOT chart serialization.
- `desktop/gridforge_desktop.py`: native file-save dialog permits static SVG chart output.
- `tests/charts_v107.test.mjs`: correctness, sanitization, persistence, formatting and recalculation tests.

Security: chart range/title inputs are bounded; SVG labels and titles are XML-escaped; imported chart metadata is validated; no arbitrary scripting, imported VBA or external chart service is executed.
