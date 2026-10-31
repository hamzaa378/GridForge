# GridForge v11.29 Feature Matrix

| Area | Status | Current implementation and limitation |
|---|---|---|
| Draw tab | Implemented in GridForge | Pen/highlighter, line, arrow, rectangle, oval, text, selection/move, eraser, undo, replay, color, width, and shape fill persist in GridForge JSON. Native Office ink conversion and Excel drawing-object XLSX fidelity remain unavailable; export confirms before omitting GridForge drawings. |
| Page Layout | Improved | Ribbon quick controls and Page Setup include margin presets, paper/orientation, print area, fit scale, repeated rows/columns, gridlines/headings, headers/footers, printing, and page-break preview. Print sends a self-contained view to the operating system's printer dialog; any installed printer, including Microsoft Print to PDF when enabled, can be selected there. Header/footer settings round-trip through supported XLSX OOXML. Browser print repeats title rows; it does not reproduce repeated columns, full Excel pagination, or native page breaks. |
| Header/footer values | Implemented for supported subset | Left, center, and right positions are available for odd-page headers and footers. `&P`, `&N`, `&A`, and `&D` are rendered in browser printing and preserved as Excel header/footer tokens in XLSX. Even/first-page variants and rich-text font/format controls are not exposed. |
| PivotTables | Broad supported subset | Multi-field rows/columns, multiple calculations, grouping, filters, sorting, subtotals, totals, slicers, refresh, show-details sheets, and supported XLSX PivotTable metadata are present. Native timelines, full Excel field-list interaction, complete PivotChart controls, Power Query M, Data Model, DAX, and OLAP are unavailable. See `EXCEL_COMPATIBILITY_AUDIT.md`. |
| Excel compatibility | Partial | GridForge is a standalone spreadsheet application. This release improves Draw and print setup; it does not claim full Excel UI or file-format parity. |

## Code-only verification

- `npm run check` — passed.
- `npm run build` — passed; regenerated `dist/`.
- Windows packaging is compiled with `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI validation were run.
