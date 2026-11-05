# GridForge v12.0.203 — Column and bar gap width

Column, bar, stacked, 100% stacked, and combo charts now expose an Excel-style category gap-width setting from 0% to 500% in the chart editor. Clustered column, bar, and combo charts also expose series overlap from -100% to 100%. The SVG preview responds to both settings. XLSX exports and imports native DrawingML `c:gapWidth` and `c:overlap`; ODS writes and imports the corresponding chart style properties.

ODS chart import now recognizes `local-table` source ranges as references to the current worksheet. This lets the native chart importer recover chart settings without relying on GridForge's private chart metadata.

Regression coverage checks bounds, preview generation, native XLSX round-trip, ODS chart serialization, and ODS import after removing GridForge chart metadata. No live UI, Excel, or LibreOffice GUI test was run.
