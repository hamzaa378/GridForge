# GridForge v12.0.149 — OpenDocument charts

- ODS export now writes spreadsheet-anchored chart objects and stores SVG previews for every GridForge chart type.
- Column, bar, stacked and percent-stacked, line, area, pie, doughnut, scatter, and radar charts are written as native ODF chart documents with an embedded data snapshot and standard chart series.
- GridForge chart settings are carried in a namespaced-purpose description field and restored when the file is reopened in GridForge. Native ODF charts with supported chart classes and a contiguous source range can also be imported as editable GridForge charts.
- Specialized chart types remain preview images in Calc. ODS chart styles, labels, axes, formulas, and linked-data behavior are not fully equivalent to Calc; the native embedded snapshot is separate from subsequent worksheet edits.
- Static TypeScript build passed. No automated tests, live UI, or LibreOffice GUI round-trip was run.
