# GridForge v12.0.204 — Chart line styles

Line, scatter, radar, area, and combo chart series now support solid, dash, dot, and dash-dot line styles. The editor exposes the choice for series that render as lines, and SVG previews show the configured dash pattern.

XLSX writes and reads native DrawingML preset dash styles. ODS writes and reads SVG-compatible dash arrays in chart series styles. Regression coverage validates settings and round-trips multiple series through native XLSX charts and through ODS after removing GridForge-only chart metadata.

No live UI, Excel, or LibreOffice GUI test was run.
