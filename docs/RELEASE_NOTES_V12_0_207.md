# GridForge v12.0.207 — Chart category colors

Standard column, bar, stacked-column, stacked-bar, 100% stacked, and combo column series can now assign colors to individual categories. The chart editor offers up to 16 visible category color controls; the selected colors apply across the chart's column series. SVG previews use the point colors, and native XLSX DrawingML and ODS chart data-point styles preserve them.

Regression coverage checks renderer output, supported chart validation, XLSX round-trip, and ODS native import with GridForge-only metadata removed. No live UI, Excel, or LibreOffice GUI test was run.

XLOOKUP evaluation now preserves Excel's lazy `if_not_found` behavior while passing its optional match and search modes through to the function implementation. An error formula in the fallback no longer overrides a successful match; regressions cover fallback evaluation on misses, exact/wildcard/approximate matching, forward/reverse and binary search, invalid modes, nearest-value selection, and XMATCH search modes, including reverse-search tie handling when approximate matches contain duplicates. ODS and XLSX regressions also retain XLOOKUP's extended arguments on formula round-trip.
