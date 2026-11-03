# v12.0.170 — ODS combo charts

GridForge now exports combo charts as native ODF charts with per-series column/line types and primary/secondary value axes. The external ODS chart importer recognizes mixed bar/line series, axis assignments, secondary-axis bounds, and axis titles. The existing chart SVG fallback and GridForge metadata remain available for GridForge-authored charts.

ODF combo charts still have differences from Excel in chart source linkage, axis number formats, marker shape details, style variants, and advanced secondary-axis features. This implementation has only been code-checked; no automated tests or LibreOffice UI round-trip were run.
