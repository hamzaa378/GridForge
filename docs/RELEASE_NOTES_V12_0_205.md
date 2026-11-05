# GridForge v12.0.205 — Chart marker sizes

Chart series can now set marker sizes from 2 to 72. The chart editor exposes the size next to marker visibility and shape; SVG previews use the configured size. XLSX uses native DrawingML marker-size settings, and ODS uses native chart symbol-width and symbol-height properties.

Regression coverage validates the size bounds and round-trips different sizes through XLSX and ODS. The ODS test removes GridForge-only chart metadata to verify native chart style import.

No live UI, Excel, or LibreOffice GUI test was run.
