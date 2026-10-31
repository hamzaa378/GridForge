# GridForge 11.29.1 Release Notes

## Draw and Page Layout

- Added persisted Draw tools for pen/highlighter strokes, lines, arrows, rectangles, ovals, text boxes, selection/move, erasing, undo, replay, and per-object color/width. Drawings are stored in GridForge workbook JSON; XLSX/XLSM export asks before omitting them because native Office ink and drawing-object compatibility is not implemented.
- Expanded Page Setup with six header/footer positions. `&P`, `&N`, `&A`, and `&D` render as page number, page count, sheet name, and date in the browser print view. The same values round-trip through supported XLSX odd-header/footer OOXML.
- Print now sends a self-contained worksheet view to the system print dialog without opening a popup. Users can choose any installed printer, including Microsoft Print to PDF when enabled in Windows.
- Kept the existing print-area, repeated-title, paper-size, orientation, margin, fit-to-page, gridline, heading, and page-break-preview controls.

## PivotTables

- Retained the existing native PivotTable subset: nested fields and measures, aggregations, sorting, filters, grouping, subtotals/totals, slicers, refresh, details sheets, and supported XLSX PivotTable/PivotChart metadata.
- Excel timelines, full Data Model/Power Pivot/DAX/OLAP, complete Excel PivotChart field-button behavior, and exhaustive native Excel acceptance remain unsupported. See the compatibility audit for boundaries.

## Verification scope

- TypeScript check and production build passed.
- Windows installer and release archives are built with the build script's test phases skipped.
- No test suites, Windows runtime checks, application launch, live UI, or Excel GUI verification were run.
