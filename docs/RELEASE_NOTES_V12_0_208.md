# GridForge v12.0.208 — Per-marker chart colors

Line, scatter, radar, and combo line series now support individual marker colors alongside existing category colors for pie, doughnut, column, bar, stacked, 100% stacked, and combo column charts. Up to 64 point-color controls are shown in the chart editor; formats can retain up to 10,000 source-point colors. SVG previews, native XLSX DrawingML markers, and ODS data-point styles retain the selected colors. For line and combo-line series, marker fill/line colors do not replace the series stroke color. Charts with complete, distinct point-color arrays per series now retain those arrays through native XLSX and ODS import/export, including metadata-free ODS chart recovery. The chart editor also keeps imported per-series point colors when the chart type and source range remain unchanged.

Regression coverage checks chart validation and SVG output, native XLSX marker XML and round-trip import, and metadata-free ODS import. New tests cover validation limits and independent point-color palettes across multiple series in XLSX and ODS. The tests are code-only; no live application, Excel, or Calc UI was opened.

The user guide now correctly documents existing cross-sheet formula support. An XLSX regression verifies scalar and range formulas continue to update after editing their source worksheet; ODS coverage verifies the same dependency refresh after import.


The 1,108-test code-only regression suite passes, including TypeScript checking, the 10,000-formula dependency benchmark, and the 10,000-cell Data Studio edit/undo benchmark. The Windows portable app and installer compiled successfully. The application was not launched and no live UI or LibreOffice GUI checks were run.
