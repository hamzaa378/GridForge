# GridForge v12.0.206 — Chart line widths

Line based chart series now support line widths from 0.25 to 10 points in quarter-point steps. The chart editor exposes the setting alongside line dash style. SVG previews scale the configured width.

XLSX stores widths in native DrawingML line geometry; ODS writes the SVG-compatible stroke width on chart series styles and reads it back. Tests cover width validation and different-width series through native XLSX and ODS imports after removing GridForge chart metadata.

No live UI, Excel, or LibreOffice GUI test was run.
