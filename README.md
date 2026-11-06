# GridForge v12.0.210 — A flexible Excel-style spreadsheet workspace

> **Active working-tree development:** The current source adds a bounded Power Query M subset, a worksheet-backed local Data Model with a restricted DAX evaluator, a read-only local SQLite import, and an authenticated one-time HTTPS importer for the Windows desktop app. See [the Data Model/DAX limits](docs/LOCAL_DATA_MODEL_DAX_SUBSET.md), [the SQLite connector notes](docs/SQLITE_READONLY_CONNECTOR.md), [authenticated one-time web import](docs/AUTHENTICATED_WEB_IMPORT.md), and the [compatibility audit](docs/EXCEL_COMPATIBILITY_AUDIT.md). These additions are source changes in the working tree; they are not a numbered release, and Excel/Calc live-session acceptance has not been performed.

See [v12.0.210 formula compatibility](docs/RELEASE_NOTES_V12_0_210.md), [v12.0.209 scatter pairs and axis controls](docs/RELEASE_NOTES_V12_0_209.md), [v12.0.208 per-marker chart colors](docs/RELEASE_NOTES_V12_0_208.md), [v12.0.207 chart colors and lookup compatibility](docs/RELEASE_NOTES_V12_0_207.md), [v12.0.206 line widths](docs/RELEASE_NOTES_V12_0_206.md), [v12.0.205 chart markers](docs/RELEASE_NOTES_V12_0_205.md), [v12.0.204 chart line styles](docs/RELEASE_NOTES_V12_0_204.md), [v12.0.203 chart spacing](docs/RELEASE_NOTES_V12_0_203.md), and [v12.0.202 reversed ODS icon sets](docs/RELEASE_NOTES_V12_0_202.md).

See [v12.0.201 typed data-bar bounds](docs/RELEASE_NOTES_V12_0_201.md), [v12.0.200 native ODS color-scale thresholds](docs/RELEASE_NOTES_V12_0_200.md), [v12.0.199 conditional-formatting fidelity](docs/RELEASE_NOTES_V12_0_199.md), [v12.0.197 formula conditional formatting](docs/RELEASE_NOTES_V12_0_197.md), [v12.0.196 conditional data-bar controls](docs/RELEASE_NOTES_V12_0_196.md), [v12.0.195 ODS Calc conditional-format interoperability](docs/RELEASE_NOTES_V12_0_195.md), [v12.0.194 ODS horizontal chart axis fidelity](docs/RELEASE_NOTES_V12_0_194.md), [v12.0.193 ODS chart axis and formula import fidelity](docs/RELEASE_NOTES_V12_0_193.md), [v12.0.192 ODS chart and formula compatibility](docs/RELEASE_NOTES_V12_0_192.md), [v12.0.191 ODS border width fidelity](docs/RELEASE_NOTES_V12_0_191.md), [v12.0.190 Calc precise rounding formulas](docs/RELEASE_NOTES_V12_0_190.md), [v12.0.189 chart marker shapes](docs/RELEASE_NOTES_V12_0_189.md), [v12.0.188 ODS random-array compatibility](docs/RELEASE_NOTES_V12_0_188.md), [v12.0.187 ODS dynamic-array function interoperability](docs/RELEASE_NOTES_V12_0_187.md), [v12.0.186 ODS lookup and sequence formulas](docs/RELEASE_NOTES_V12_0_186.md), [v12.0.185 ODS chart trendlines and error bars](docs/RELEASE_NOTES_V12_0_185.md), [v12.0.184 formatted print output](docs/RELEASE_NOTES_V12_0_184.md), [v12.0.183 ODS chart and font fidelity](docs/RELEASE_NOTES_V12_0_183.md), [v12.0.182 Calc alignment fidelity](docs/RELEASE_NOTES_V12_0_182.md), [v12.0.181 ODS array formula round-trips](docs/RELEASE_NOTES_V12_0_181.md), [v12.0.180 Calc array formula mappings](docs/RELEASE_NOTES_V12_0_180.md), [v12.0.179 Calc text and criteria formula mappings](docs/RELEASE_NOTES_V12_0_179.md), [v12.0.178 Calc probability distribution compatibility](docs/RELEASE_NOTES_V12_0_178.md), [v12.0.177 Calc distribution formula interoperability](docs/RELEASE_NOTES_V12_0_177.md), [v12.0.176 Calc probability formula interoperability](docs/RELEASE_NOTES_V12_0_176.md), and [v12.0.175 Calc statistical formula interoperability](docs/RELEASE_NOTES_V12_0_175.md).

See [v12.0.174 Calc statistical formulas](docs/RELEASE_NOTES_V12_0_174.md), [v12.0.173 Calc-compatible ODS formula names](docs/RELEASE_NOTES_V12_0_173.md), [v12.0.172 larger ODS chart data caches](docs/RELEASE_NOTES_V12_0_172.md), [v12.0.171 ODS chart axis number formats](docs/RELEASE_NOTES_V12_0_171.md), [v12.0.170 ODS combo charts](docs/RELEASE_NOTES_V12_0_170.md), [v12.0.169 ODS chart categories and markers](docs/RELEASE_NOTES_V12_0_169.md), [v12.0.168 ODS chart axis fidelity](docs/RELEASE_NOTES_V12_0_168.md), [v12.0.167 ODS drawing import](docs/RELEASE_NOTES_V12_0_167.md), [v12.0.166 ODS fractional seconds](docs/RELEASE_NOTES_V12_0_166.md), [v12.0.165 ODS formula fallback](docs/RELEASE_NOTES_V12_0_165.md), [v12.0.164 ODS date/time formats](docs/RELEASE_NOTES_V12_0_164.md), [v12.0.163 ODS number-format colors](docs/RELEASE_NOTES_V12_0_163.md), [v12.0.162 ODS number-format maps](docs/RELEASE_NOTES_V12_0_162.md), [v12.0.161 ODS chart-axis bounds](docs/RELEASE_NOTES_V12_0_161.md), [v12.0.160 ODS chart-label fidelity](docs/RELEASE_NOTES_V12_0_160.md), [v12.0.159 ODS worksheet-local names](docs/RELEASE_NOTES_V12_0_159.md), [v12.0.158 ODS conditional style import](docs/RELEASE_NOTES_V12_0_158.md), [v12.0.157 ODS conditional fills](docs/RELEASE_NOTES_V12_0_157.md), [v12.0.156 ODS drawing and chart fidelity](docs/RELEASE_NOTES_V12_0_156.md), [v12.0.155 ODS drawing previews](docs/RELEASE_NOTES_V12_0_155.md), [v12.0.154 native ODS stock charts](docs/RELEASE_NOTES_V12_0_154.md), [v12.0.153 ODS worksheet object round-trips](docs/RELEASE_NOTES_V12_0_153.md), [v12.0.152 ODS formula preservation](docs/RELEASE_NOTES_V12_0_152.md), [v12.0.151 half-point font sizing](docs/RELEASE_NOTES_V12_0_151.md), [v12.0.150 locale formula parsing](docs/RELEASE_NOTES_V12_0_150.md), [v12.0.149 ODS charts](docs/RELEASE_NOTES_V12_0_149.md), [v12.0.148 ODS number formats](docs/RELEASE_NOTES_V12_0_148.md), [v12.0.147 ODS filters](docs/RELEASE_NOTES_V12_0_147.md), [v12.0.146 diagonal borders](docs/RELEASE_NOTES_V12_0_146.md), [v12.0.145 ODS hyperlinks](docs/RELEASE_NOTES_V12_0_145.md), [v12.0.144 ODS data validation](docs/RELEASE_NOTES_V12_0_144.md), [v12.0.143 mixed border edges](docs/RELEASE_NOTES_V12_0_143.md), [v12.0.142 chart and ODS compatibility](docs/RELEASE_NOTES_V12_0_142.md), [v12.0.141 ODS formula support](docs/RELEASE_NOTES_V12_0_141.md), [v12.0.140 border formatting and export](docs/RELEASE_NOTES_V12_0_140.md), [v12.0.139 LibreOffice-friendly font formatting](docs/RELEASE_NOTES_V12_0_139.md), [v12.0.138 OpenDocument custom number-format notes](docs/RELEASE_NOTES_V12_0_138.md), [v12.0.137 merged regions and performance notes](docs/RELEASE_NOTES_V12_0_137.md), [v12.0.136 OpenDocument formatting notes](docs/RELEASE_NOTES_V12_0_136.md), and [v12.0.135 OpenDocument support notes](docs/RELEASE_NOTES_V12_0_135.md).


See [v12.0.133 worksheet capacity and compatibility audit](docs/RELEASE_NOTES_V12_0_133.md).


See [v12.0.95 formula compatibility notes](docs/RELEASE_NOTES_V12_0_95.md), [v12.0.96 currency text notes](docs/RELEASE_NOTES_V12_0_96.md), [v12.0.97 time parsing notes](docs/RELEASE_NOTES_V12_0_97.md), [v12.0.98 TIME behavior notes](docs/RELEASE_NOTES_V12_0_98.md), [v12.0.99 time-component notes](docs/RELEASE_NOTES_V12_0_99.md), [v12.0.100 chart accessibility notes](docs/RELEASE_NOTES_V12_0_100.md), [v12.0.101 locale time parsing notes](docs/RELEASE_NOTES_V12_0_101.md), [v12.0.102 Draw selection notes](docs/RELEASE_NOTES_V12_0_102.md), [v12.0.103 compatibility improvements](docs/RELEASE_NOTES_V12_0_103.md), [v12.0.104 financial iteration notes](docs/RELEASE_NOTES_V12_0_104.md), [v12.0.105 NPV reference behavior](docs/RELEASE_NOTES_V12_0_105.md), [v12.0.106 counting and date parsing notes](docs/RELEASE_NOTES_V12_0_106.md), [v12.0.107 conditional aggregation updates](docs/RELEASE_NOTES_V12_0_107.md), [v12.0.108 date/time criteria notes](docs/RELEASE_NOTES_V12_0_108.md), [v12.0.109 MAXIFS/MINIFS notes](docs/RELEASE_NOTES_V12_0_109.md), [v12.0.110 named hyperlink navigation](docs/RELEASE_NOTES_V12_0_110.md), [v12.0.111 sparse count ranges](docs/RELEASE_NOTES_V12_0_111.md), [v12.0.112 sparse COUNTIF ranges](docs/RELEASE_NOTES_V12_0_112.md), [v12.0.113 sparse COUNTIFS ranges](docs/RELEASE_NOTES_V12_0_113.md), [v12.0.114 sparse SUMIFS/AVERAGEIFS ranges](docs/RELEASE_NOTES_V12_0_114.md), and [v12.0.115 sparse MAXIFS/MINIFS ranges](docs/RELEASE_NOTES_V12_0_115.md), and [v12.0.116 sparse criteria-range error handling](docs/RELEASE_NOTES_V12_0_116.md), and [v12.0.117 compact range dependencies](docs/RELEASE_NOTES_V12_0_117.md), and [v12.0.118 range-aware formula tracing](docs/RELEASE_NOTES_V12_0_118.md), and [v12.0.119 range dependency safeguards](docs/RELEASE_NOTES_V12_0_119.md), and [v12.0.120 larger sparse aggregation ranges](docs/RELEASE_NOTES_V12_0_120.md), and [v12.0.121 sparse basic aggregations](docs/RELEASE_NOTES_V12_0_121.md), and [v12.0.122 sparse range index](docs/RELEASE_NOTES_V12_0_122.md), and [v12.0.123 sparse order statistics](docs/RELEASE_NOTES_V12_0_123.md), and [v12.0.124 sparse aggregate family](docs/RELEASE_NOTES_V12_0_124.md), [v12.0.125 Excel variance aliases](docs/RELEASE_NOTES_V12_0_125.md), and [v12.0.126 safer ChartEx imports](docs/RELEASE_NOTES_V12_0_126.md), and [v12.0.127 compact ribbon sizing](docs/RELEASE_NOTES_V12_0_127.md), and [v12.0.128 manual page breaks](docs/RELEASE_NOTES_V12_0_128.md), and [v12.0.129 formula compatibility](docs/RELEASE_NOTES_V12_0_129.md), and [v12.0.130 locale number formatting](docs/RELEASE_NOTES_V12_0_130.md), and [v12.0.131 localized chart axes](docs/RELEASE_NOTES_V12_0_131.md), and [v12.0.132 localized date names](docs/RELEASE_NOTES_V12_0_132.md).

GridForge is an offline-first spreadsheet application with its own worksheet grid, formula engine and Excel-inspired ribbon; Microsoft Excel is not required. Version 12.0 makes the interface easier to adapt to different work styles and accessibility needs while carrying forward the cloud workbook connector. Version 12.0.55 renders numbers to 15 significant digits for Excel-like general text output; v12.0.49 corrects the formula reference and compatibility counts to match the 506 local and 10 connected names, documents scoped LAMBDA and dynamic spilling accurately, and rejects oversized spill outputs before populating neighboring cells; v12.0.48 keeps Treemap and Sunburst hierarchy tuples distinct even when category labels contain the display separator; v12.0.47 flattens nested JSON records into bounded worksheet columns; v12.0.46 moves CUBESETCOUNT from a misleading offline text counter to the connected-function list; v12.0.45 indexes PIVOTBY row and column prefix groups to reduce repeated full-group scans; v12.0.44 lets public JSON refresh queries extract a nested array through a JSON Pointer before loading and transformation; v12.0.43 adds nested GROUPBY row subtotals with top or bottom totals, hierarchy-safe ordering, and an output-size preflight; v12.0.42 adds nested PIVOTBY row and column subtotals with matching safeguards; v12.0.41 separates PIVOTBY field captions from hierarchy values in multi-level headers; v12.0.40 moves routine workbook autosaves to IndexedDB, with migration and fallback for existing local-storage saves; v12.0.39 supports vector sort-order arrays for multi-field PIVOTBY; v12.0.38 adds PIVOTBY parent-row and parent-column aggregation bases; v12.0.37 extends formula-based PIVOTBY to multiple row and column field levels with hierarchy output; v12.0.36 lets the Name Box navigate cross-sheet cells and ranges, including quoted sheet names; v12.0.35 updates the Help panel with the added navigation, selection, sheet, filter, and fill shortcuts; v12.0.34 adds Ctrl+Arrow navigation through populated regions and Home row navigation; v12.0.33 adds Ctrl+R fill-right with formula reference adjustment; v12.0.32 expands the Name Box to select ranges and supported defined names; v12.0.31 adds keyboard selection for worksheet rows, columns, and the full sheet; v12.0.30 adds Ctrl+Home/Ctrl+End navigation to the workbook used area; v12.0.29 adds Excel-style worksheet navigation shortcuts and remembers each sheet’s selection and scroll position; v12.0.28 adds multi-measure PIVOTBY output and defers formula recalculation during workbook restore to avoid repeating worker-side calculations on the interface thread; v12.0.27 adds worksheet-local named LAMBDA formulas; v12.0.26 expands formula-based GROUPBY and PIVOTBY summaries with filtering, sorting, grand totals, multiple grouped fields in GROUPBY, and LAMBDA aggregators; v12.0.25 adds Excel-style LAMBDA evaluation to MAP, BYROW, BYCOL, MAKEARRAY, REDUCE, and SCAN; v12.0.24 keeps the Draw and Page Layout ribbons compact; v12.0.23 adds `_xlfn.` serialization for PERCENTOF, ISOMITTED, and service-backed DETECTLANGUAGE; v12.0.22 added FIELDVALUE and STOCKHISTORY; v12.0.21 expands REGEXEXTRACT to first-match, all-match, and capture-group modes; v12.0.20 fixes XLSX `_xlfn.` prefixes for newer functions; v12.0.19 fixes Excel 1900 date serial conversion in date validation and Pivot dates; v12.0.18 improves explicit @ implicit intersection for computed arrays; v12.0.17 fixes the Draw and Page Layout ribbon height; v12.0.16 reduces peak memory during CSV import; v12.0.15 moves GridForge JSON import work off the interface thread; v12.0.14 does the same for CSV import; v12.0.13 adds export size safeguards; v12.0.12 moves CSV and JSON export work into a background worker, complementing v12.0.11's XLSX export worker; v12.0.10 added Box workbook access.

## v12.0 highlights

- Import and export bounded OpenDocument spreadsheets (`.ods`). Supported formulas round-trip through OpenFormula, including relative-reference adjustment for repeated formula cells; formulas outside the translatable subset retain cached values. ODS also retains supported values, sheet names and visibility, common cell formatting and print options, validations, safe hyperlinks, AutoFilters, and chart objects. Common chart types export as editable ODF charts with SVG previews, and every GridForge chart type exports a preview. GridForge-authored chart settings round-trip; external native ODF chart import covers supported types with a contiguous source range. Specialized external charts, advanced chart styling, PivotTables, macros, and unrelated embedded objects are not preserved.
- Choose among four persistent visual presets: Excel 2019, Modern, Compact, and High Contrast. The choice changes presentation only and does not alter workbooks.
- Combine a visual preset with independent light/dark color mode and Excel-style or Guided workspace. Guided mode keeps its beginner actions available regardless of the visual preset.
- Added a high contrast palette with stronger focus outlines and clearer active ribbon and cell selections, plus a compact layout for smaller displays and reduced scrolling.
- Existing Excel-style command groups, worksheet tools, cloud open/save actions, printing, charts, and PivotTable subset remain available from the same app.
- GridForge remains a separate spreadsheet engine; the redesign does not establish complete Excel feature or file-format parity.

See [v12 feature matrix](docs/FEATURE_VERIFICATION_V12.md), [v12 release notes](docs/RELEASE_NOTES_V12.md), [12.0.1 responsiveness notes](docs/RELEASE_NOTES_V12_0_1.md), [12.0.2 ChartEx notes](docs/RELEASE_NOTES_V12_0_2.md), and [12.0.3 notes](docs/RELEASE_NOTES_V12_0_3.md) and [12.0.4 import responsiveness notes](docs/RELEASE_NOTES_V12_0_4.md) and [12.0.5 chart error-bar notes](docs/RELEASE_NOTES_V12_0_5.md) and [12.0.6 ribbon and locale notes](docs/RELEASE_NOTES_V12_0_6.md), [12.0.9 chart notes](docs/RELEASE_NOTES_V12_0_9.md), [12.0.10 Box cloud notes](docs/RELEASE_NOTES_V12_0_10.md), [12.0.11 XLSX export responsiveness notes](docs/RELEASE_NOTES_V12_0_11.md), [12.0.12 CSV and JSON export responsiveness notes](docs/RELEASE_NOTES_V12_0_12.md), [12.0.13 export safety notes](docs/RELEASE_NOTES_V12_0_13.md), [12.0.14 CSV import responsiveness notes](docs/RELEASE_NOTES_V12_0_14.md), [12.0.15 JSON import responsiveness notes](docs/RELEASE_NOTES_V12_0_15.md), [12.0.16 CSV memory notes](docs/RELEASE_NOTES_V12_0_16.md), and [12.0.17 ribbon layout notes](docs/RELEASE_NOTES_V12_0_17.md), [12.0.18 formula compatibility notes](docs/RELEASE_NOTES_V12_0_18.md), [12.0.19 date compatibility notes](docs/RELEASE_NOTES_V12_0_19.md), [12.0.20 XLSX formula compatibility notes](docs/RELEASE_NOTES_V12_0_20.md), [12.0.21 REGEXEXTRACT notes](docs/RELEASE_NOTES_V12_0_21.md), and [12.0.22 connected formula serialization notes](docs/RELEASE_NOTES_V12_0_22.md), and [12.0.23 function serialization notes](docs/RELEASE_NOTES_V12_0_23.md), and [12.0.24 ribbon sizing notes](docs/RELEASE_NOTES_V12_0_24.md), [12.0.25 LAMBDA notes](docs/RELEASE_NOTES_V12_0_25.md), and [12.0.26 GROUPBY/PIVOTBY notes](docs/RELEASE_NOTES_V12_0_26.md), plus [12.0.27 named LAMBDA notes](docs/RELEASE_NOTES_V12_0_27.md), [12.0.28 responsiveness notes](docs/RELEASE_NOTES_V12_0_28.md), [12.0.49 function/spill limits](docs/RELEASE_NOTES_V12_0_49.md), and [12.0.55 numeric precision](docs/RELEASE_NOTES_V12_0_50.md).

## v11.30 highlights (incremental release)

- Connect user-configured OneDrive, Google Drive, Dropbox, or Box desktop OAuth apps to browse, open, save a workbook, or save a new copy. Refresh credentials are protected with Windows DPAPI; Box also stores its app secret using DPAPI. Cloud save is explicit and checks cloud revisions before overwriting. See [cloud storage setup](docs/CLOUD_STORAGE_SETUP.md).
- Local spreadsheet use remains offline. Provider app registration is required; cloud listing is root-only, Google Drive uses `drive.file`, Box simple uploads are limited to 50 MB, and native Google Sheets import and continuous sync are not supported.

See [v11.30 feature matrix](docs/FEATURE_VERIFICATION_V11_30.md) and [release notes](docs/RELEASE_NOTES_V11_30.md).

## v11.29 highlights

- Added persistent Draw tools for pen/highlighter, shapes, arrows, text boxes, selection/move, erasing, undo, and replay. GridForge-only drawings remain stored in its JSON format; XLSX export warns before omitting them.
- Added six header/footer positions to Page Setup with `&P`, `&N`, `&A`, and `&D` tokens, browser-print rendering, and supported XLSX OOXML round-tripping.
- Retained page setup, print area, repeated titles, margins, fit-to-page, print options, and page-break preview.
- Preserved the broad existing PivotTable subset and native metadata support; Excel timelines, Data Model, DAX, OLAP, and complete PivotChart controls remain unavailable.

See [v11.29 feature matrix](docs/FEATURE_VERIFICATION_V11_29.md) and [release notes](docs/RELEASE_NOTES_V11_29.md).

## v11.27 highlights

- Polynomial trendlines now fit in a centered, scaled domain with QR factorization, reducing instability when X values are large (such as date serials). The displayed equation identifies the normalized X variable used by the fit.
- Moving-average chart trendlines retain the actual X positions for irregularly spaced scatter data instead of drawing averages at inferred evenly spaced positions.
- Multi-criteria formulas now reject ranges with mismatched row/column shapes, even when their flattened cell counts match. Chart series can use calculated Standard Error bars, with supported native XLSX metadata round-tripping.
- Added 12 Excel Engineering conversions between binary, octal, decimal, and hexadecimal, including Excel's signed two's-complement ranges, optional zero-padding, and input/output limits.
- `FLOOR` rejects mixed-sign number/significance inputs with Excel's `#NUM!` result and uses a small floating-point tolerance around exact multiples to prevent values such as `FLOOR(1.1,0.1)` from falling a full increment early. `CEILING` uses the same boundary handling.
- These are targeted compatibility improvements. They do not establish full formula or chart parity with Excel.

## v11.26 highlights
- `DAYS360` now distinguishes Excel's U.S. (NASD) and European methods, including month-end adjustments. `YEARFRAC` basis 0 and 4 use their matching conventions; basis 1 accumulates actual days against each year's serial-day length.
- Date serial validation now follows Excel's supported 1900-system range through 9999-12-31 in shared date helpers and working-day functions, including `WORKDAY.INTL` / `NETWORKDAYS.INTL`, while retaining serial 60 for Excel's fictitious 1900-02-29. Invalid all-weekend masks are rejected by `WORKDAY.INTL`.
- Formula references are stronger: `INDIRECT` handles A1/R1C1 references, local names and cross-sheet ranges; `OFFSET` now returns tracked rectangular references; `ADDRESS` supports A1/R1C1 and relative/absolute modes.
- `CELL` and `ISFORMULA`/`FORMULATEXT` inspect supported references across worksheets, while filter-aware `SUBTOTAL`/`AGGREGATE` refresh when row visibility changes.
- Scatter charts now plot multiple Y series against numeric X values and support horizontal axis bounds, units and formats that round-trip in supported XLSX charts.
- Pie and doughnut charts add first-slice rotation, hole sizing, per-slice colors, label placement and supported XLSX round-tripping.

- Data tab includes clear entry points for Power Query, the Data Model, Power Pivot/DAX, OLAP, native Timelines, and PivotChart field buttons. Selecting one opens a notice that explains what is missing and what GridForge supports instead.
- Help → Feature availability lists major compatibility gaps, partial implementations, and alternatives, including formula compatibility, specialized chart interchange, PivotCharts, and external refresh.
- Pivot Summary field counts follow source dimensions and available worksheet/output space instead of the previous fixed eight-field cap. The source picker accepts the engine's full 104,000-cell source bound. Each selected row or column field can now use its own Values calculation for ascending or descending value sorting, and those sort settings round-trip through native PivotTable AutoSort metadata.
- Native XLSX PivotTable import, summary calculation, and export support a separate Top/Bottom value filter on each selected row or column field. The Pivot Summary dialog can add, edit, and remove multiple filters, each with its own field, Values calculation, direction, and item count.
- Pivot item filters distinguish no filter from an explicit empty selection, preserving imported Excel states where no items are visible. Slicer items toggle independently while multi-selecting, and clearing a slicer restores all items.
- XLSX PivotTable import resolves workbook-scoped names that point directly to worksheet cell ranges, in addition to supported tables and local names. Formula-based and external-workbook names remain unsupported.
- Compatibility follow-up: supported Excel date/time number formats now round-trip through XLSX, and 1904-system date literals normalize on import. Histogram bin width, underflow, and overflow settings round-trip through native ChartEx; ChartEx legends preserve placement; Waterfall can round-trip a final total point. Formula adjustments cover XLOOKUP row/column arrays, selected date/time and blank/count behavior, VALUE and NUMBERVALUE current-locale separators and localized digits, NUMBERVALUE percent signs, ASCII-space TRIM, DATEDIF boundaries, and XNPV validation; these remain targeted fixes rather than full Excel equivalence.
- Lookup compatibility now includes `MATCH` exact and sorted approximate modes, `XMATCH` and `XLOOKUP` forward/reverse and sorted binary searches, plus XLOOKUP wildcard and next-smaller/next-larger modes. Binary search validates ordering; broad Excel comparison across coercion and locale edge cases remains outstanding.
- Waterfall charts accept multiple total/subtotal point indexes in their editor and persist supported indexes using native ChartEx subtotal metadata. Page Setup now imports/exports Print Titles; repeated rows also repeat in browser print, while repeated columns are retained for Excel.
- Combo chart XLSX import/export now retains mixed column/line series and their primary/secondary axis assignments. Validation rules now preserve Excel input prompts and stop/warning/information alerts, with matching controls in the editor.
- `TEXTSPLIT` accepts delimiter arrays, case-insensitive matching, ignore-empty handling for row/column delimiters, and custom padding for uneven rows. Excel’s default error padding and some locale-sensitive casing remain incomplete.
- VLOOKUP/HLOOKUP and LOOKUP approximate modes now validate ascending order; TEXTBEFORE/TEXTAFTER support directional instances, case mode, match-at-end, fallback, and empty-input rules. View → Gridlines now controls persisted worksheet gridline visibility while keeping formatted cell borders.
- Standard chart label positions are editable and round-trip through supported DrawingML, and series names matching headers export as linked cell references. Best Fit rendering is approximate; ChartEx label positions remain unsupported.
- The cell context menu now inserts, edits, or removes worksheet hyperlinks. Imported, created, and direct `HYPERLINK` formula targets render with hyperlink styling and open supported external targets or select in-workbook cell ranges on visible worksheets. Nested formula links and named destinations remain unsupported.
- Standard chart labels can show values, categories, series names, and pie/doughnut percentages. `SHEET()` and `SHEETS()` use workbook worksheet order and count for supported single-sheet references; selected `AGGREGATE` statistical operations now include median, mode, and percentile/quartile variants.
- Cell formatting now includes vertical top/middle alignment and indent controls, with supported settings preserved in XLSX.
- Custom Excel number-format codes can be entered in the formatting toolbar and preserved through XLSX round-trips; the grid renders a supported subset and safely falls back for unsupported tokens.
- Conditional formatting now supports optional three-color scales with a 50th-percentile midpoint, preserved in native XLSX color-scale rules.
- Conditional numeric rules include inclusive comparisons, not-equal, and between/not-between, with supported operators preserved in XLSX.
- Formula conditional formatting evaluates relative and absolute references in the target-cell context. XLSX expression rules and ODS formula style maps round-trip; XLSX retains conditional-rule priorities and stop-if-true flags. ODS blank-cell coverage and full rule precedence remain partial.
- Dynamic-array `FILTER` can include rows or columns and return an array for an empty result; `SORT` supports sorting across columns, and `SORTBY` supports multiple stable keys with independent directions.
- Dynamic-array `TAKE` and `DROP` distinguish omitted arguments from explicit zero counts and return `#CALC!` for an empty result.
- Imported PivotTables retain manually ordered field items during refresh and XLSX export. Choose “Manual” in the Pivot Summary dialog to reorder items with Move Up/Move Down controls; manually grouped fields list their group captions, and date or numeric grouping lists the displayed group labels instead of raw source values. Opening an existing Pivot after source data changes, or changing its source range, reconciles added and removed categories while retaining the relative order of existing items; new values follow source order. “Manual (starts in source order)” retains the original encounter order until you reorder items.
- Source release includes the TypeScript source, compiled browser app (`dist/`), desktop wrapper, installer recipe, GitHub Actions workflow, and complete dependency/build instructions.
- Compiled release artifacts: `build/release/GridForge-Setup-12.0.210-win64.exe` (installer), `GridForge-Portable-win64.zip` (portable app), `GridForge-v12.0.210-source.zip` (source with compiled `dist/`), and `GridForge-v12.0.210-GitHub-Repository.zip` (source plus installer). SHA-256 values are in `build/release/SHA256SUMS.txt`.

See [v11.28 feature matrix](docs/FEATURE_VERIFICATION_V11_28.md), [release notes](docs/RELEASE_NOTES_V11_28.md), [v11.27 feature matrix](docs/FEATURE_VERIFICATION_V11_27.md), and [desktop build guide](docs/DESKTOP_BUILD_AND_RELEASE.md).

## v11.22 highlights

- Define manual PivotTable groups with `Group = item | item`, one group per line; unassigned values remain visible as independent categories.
- Manual groups feed the summary, filters and native XLSX PivotTable discrete grouping metadata. Supported imports restore group captions and item mappings.
- TypeScript check/build and installer compilation only; no automated tests, Windows runtime checks, live UI checks or Excel GUI validation.
- See the [v11.22 feature matrix](docs/FEATURE_VERIFICATION_V11_22.md) and [release notes](docs/RELEASE_NOTES_V11_22.md).
## v11.15 highlights

- Create stock charts from Open, High, Low and Close columns in any order, with an optional date/label column.
- Invalid high/low bounds show a chart message, and incomplete rows are skipped.
- The Inno Setup installer and desktop metadata are updated to version 12.0.0; the installer was compiled without running Windows test phases.
- Stock chart XLSX export currently uses a clustered-column fallback.
- See the [v11.15 feature matrix](docs/FEATURE_VERIFICATION_V11_15.md) and [release notes](docs/RELEASE_NOTES_V11_15.md).

## v11.14 highlights

- Add **Funnel** charts for sequential stages and **Treemap** charts using one or more category columns plus a numeric value column.
- Create a **Descriptive Statistics** report with mean, standard error, median, mode, sample variance and standard deviation, kurtosis, skewness, range, min, max, sum and count.
- Statistics are written to a separate output range; text, blanks and formula errors are ignored, and overlapping the source is rejected.
- Specialized charts still export to XLSX as clustered columns. See the [v11.14 feature matrix](docs/FEATURE_VERIFICATION_V11_14.md) and [release notes](docs/RELEASE_NOTES_V11_14.md).

## v11.13 highlights

- Create Pareto charts with categories ranked by contribution, an accumulated percentage line, and a combined “Other” category for long tails.
- Create box-and-whisker charts with interpolated quartiles, Tukey whiskers and outlier markers.
- Set histogram bin count from 1 to 50, or leave it at 0 for automatic bins.
- Use Goal Seek to preview a numeric input that drives a formula toward a target. Calculations run on workbook copies; apply checks that the workbook has not changed since preview.
- Histogram now exports and imports as native ChartEx with a source range and either bin count or bin width for one numeric series, plus title, data-label, and series-color settings. Treemap and Sunburst preserve multi-level category columns; other advanced hierarchy/style behavior remains limited.
- Refresh public HTTPS CSV and JSON query sources from the Data tab. Endpoints must allow browser CORS; requests send no credentials, reject redirects, and are capped at 3 MB. XLSX exports include the last fetched result as cell values, not as an Excel connection.
- Goal Seek is bounded to one changing cell on the active worksheet and a finite numeric target; it is not a general Excel Solver replacement.

See the [v11.13 feature matrix](docs/FEATURE_VERIFICATION_V11_13.md), [release notes](docs/RELEASE_NOTES_V11_13.md) and [continuation guide](docs/CONTINUATION_GUIDE_V11_13.md).

## v11.12 highlights

- Use direct references across worksheets, including quoted names and ranges, such as `=Sheet2!A1` or `=SUM('Quarter 1'!B2:B20)`.
- Create Pivot Summary reports, including Values-only summaries with no row fields, with up to eight nested fields in each of Rows and Columns with independently reorderable Rows, Columns, and primary Values fields, up to eight value calculations (including repeated source fields and custom captions), independent Sum, Count, Count Numbers, Average, Minimum, Maximum, Product, StdDev, StdDevp, Var and VarP aggregations, with Excel-like numeric-versus-text defaults, Show Values As percentages of grand, row, column, parent-row, parent-column, % Of the Previous or Next item, running-total, percent-running-total, Difference From, % Difference From, and rank ascending/descending within a selected Base field, and Index, hierarchical row subtotals, tabular, outline, or compact report layout, subtotals above or below row groups, with optional blank lines after parent groups, manual categorical grouping for selected row/column items, plus one-level date grouping by year, quarter, month, or day and numeric interval grouping with data-derived or explicitly set bounds, for a selected row/column field, a global sort default with independent ascending/descending/source-order settings for each row and column field, value sorting by a selected calculation, numeric/date-aware field sorting, and Top/Bottom item filters ranked within each parent group by a selected Values calculation, optional row/column grand totals, row/column item filters and caption filters (including inclusive Between/Not between), and up to eight searchable multi-item report filters. Inclusive date-range filters can be applied to ungrouped row or column fields and are stored with the Pivot Summary; XLSX export represents them as PivotField item visibility. Large item lists are searched in bounded chunks. Blank row and column items are retained, and ordinary edits refresh only summaries whose sources changed, including pivots stored on one worksheet and sourced from another; appending rows to a source table expands matching PivotTable sources automatically; Refresh Summaries refreshes every saved report across the workbook and protects occupied output cells, and charts created from a summary track its refreshed range. Native XLSX PivotTable metadata, including supported percentage modes, including native x14 parent-total and percent-running-total calculations plus standard % Of, running totals, Difference From, rank, and Index calculations, per-field value-sort scope, caption filters, and Top/Bottom AutoShow filters, is written for this bounded layout.
- Refresh saved summaries manually or after ordinary edits, paste, undo and redo. Reports are generated as standard worksheet cells and can use existing chart features.
- GridForge provides persistent on-sheet slicer cards for ungrouped PivotTable row and column fields; slicer selections refresh the summary and export as PivotField item filters. Native Excel slicer objects and timelines, multi-level date hierarchies, calculated fields, Power Query M, Data Model, Power Pivot, DAX, OLAP and full Excel PivotTable editing remain unsupported. Summary-linked standard DrawingML charts export with native PivotChart links, although GridForge edits them as regular charts; specialized ChartEx charts do not carry native PivotChart links. A selected date field can be grouped by one interval (year, quarter, month, or day), and a numeric field can be grouped into custom intervals; both round-trip through supported native XLSX PivotTable cache metadata. GridForge query refresh is limited to worksheet transformations and public HTTPS CSV/JSON sources.

See the [v11.12 feature matrix](docs/FEATURE_VERIFICATION_V11_12.md), [release notes](docs/RELEASE_NOTES_V11_12.md) and [continuation guide](docs/CONTINUATION_GUIDE_V11_12.md). The 506 local function names do not establish exact Excel behavior, and 10 service-dependent names are not fully available offline.

## v11.11 highlights

- Move worksheet tabs left or right and set a tab color; colors persist in GridForge JSON and native XLSX.
- Hide or unhide selected rows and columns; visibility persists in GridForge JSON and native XLSX.
- Sort a selected table range by up to three ordered keys, with independent ascending or descending order per key.
- Find text or formulas across all visible worksheets and jump from results to the matching sheet and cell.
- Restrict cell entries by date or text length with the supported comparison operators; validation rules export to native XLSX.

Find All spans visible worksheets; Replace All remains limited to the active worksheet. Date input accepts ISO dates such as `2025-12-31` or Excel date serials. Text-length validation counts the entered text.

See the [v11.11 feature matrix](docs/FEATURE_VERIFICATION_V11_11.md) and [release notes](docs/RELEASE_NOTES_V11_11.md).

## v11.10 highlights

- Apply dropdown lists from typed options or a live range on the current worksheet.
- Restrict whole numbers and decimals with between, not-between, equality and comparison operators.
- Configure input prompts and stop, warning or information behavior for invalid entries.
- Save supported validation rules in GridForge JSON and native XLSX Data Validation records.
- Split one-column text into adjacent fields using comma, semicolon, tab, space or a custom delimiter, with a preview and quoted-field support.
- Hide and unhide worksheets; keep at least one worksheet visible and persist visibility in GridForge JSON and XLSX.

Inline lists are limited to 50 values / 255 total characters. Source ranges must be on the same worksheet; imported list validations can resolve direct defined names on that sheet and retain multiple target areas; conditional-format rules also preserve multiple areas. Custom-formula validation is not supported yet.

## v11.6 highlights

- Define, edit, delete and reuse worksheet-scoped names for cell references and ranges.
- Create names from selected heading rows or columns, and use names directly in formulas such as `=SUM(NetSales)`.
- Named range changes recalculate dependent formulas and support undo/redo.
- Names persist in GridForge JSON and the supported worksheet-scoped Excel `definedNames` subset in XLSX/XLSM.
- Retains the v11.5 Excel-style Formula ribbon: searchable function insertion, categorized function menus, SUM/AVERAGE/COUNT/MAX/MIN shortcuts, trace precedents/dependents, Show Formulas and formula inspection.
- Retains the existing chart, table, formatting, import/export, macro preservation, and desktop startup improvements from earlier releases.

Named ranges are worksheet-scoped and refer to cells or rectangular ranges. Workbook-scoped names and names defined as constants or arbitrary formulas are not included. GridForge remains a standalone application and does not claim full Excel feature parity.

GridForge uses Excel 2019 as an interaction reference while keeping its own UI and calculation engine. It does not embed, automate, or require Microsoft Excel.

## Download and run the compiled application

The compiled browser application is checked into `dist/`. To run it without building the TypeScript source, install Python 3.10 or newer and from the project folder run:

```powershell
python run_gridforge.py
```

Then open the local address printed in the terminal. On Windows, `START_HERE_WINDOWS.bat` starts the same precompiled app. The built native Windows installer is `build/release/GridForge-Setup-12.0.210-win64.exe` after the Windows build finishes; it contains the compiled desktop app and optional VBA bridge. End users need Windows 10/11 x64 and Microsoft Edge WebView2 Runtime. They do not need Node.js, npm, or Python when using the installer.

### Install using the included Windows installer

If the GitHub repository ZIP includes a current-version installer, open its `release/` folder and run `GridForge-Setup-12.0.210-win64.exe`. Follow the setup prompts and launch GridForge from the Start menu or its desktop shortcut. Windows 10/11 x64 and Microsoft Edge WebView2 Runtime are required. The installer includes the compiled application and optional Excel VBA bridge; Node.js, npm, and Python are not needed to install or run it. The installer does not install Microsoft Excel. If no installer is present, run the portable build from `build/portable/GridForge/GridForge.exe` instead.

## Software dependencies

| Task | Required software |
|---|---|
| Run the precompiled browser app | Python 3.10+ (standard library only) and a current browser; no packages are installed by `run_gridforge.py`. |
| Edit and compile the frontend | Node.js 20+ with npm; dependencies are pinned through `package-lock.json` (`typescript`, `esbuild`, and `@xmldom/xmldom`). |
| Build Windows desktop binaries | Windows 10/11 x64, Node.js 20+, Python 3.12 x64 recommended, pip/network access for the pinned requirements in `desktop/requirements-windows.txt`, and Inno Setup 6. |
| Run the installed desktop app | Windows 10/11 x64 and Microsoft Edge WebView2 Runtime. |
| Run optional VBA bridge | Desktop Microsoft Excel and the bridge executable. GridForge itself does not require Excel. |

## Build and develop from source

From the project root in PowerShell:

```powershell
npm ci
npm run check
npm run build
python run_gridforge.py
```

`npm run dev` performs a frontend build and serves it locally. `npm run start` serves the already compiled `dist/`. The default source tree includes `dist/`, so Python-only use does not need Node. Do not use `npm run dev` if you want to avoid rebuilding the compiled output.

### Build the Windows installer

Install the builder dependencies above, then run:

```powershell
./desktop/build_windows.ps1 -SkipTests
```

The `-SkipTests` option builds the frontend, portable desktop app, VBA bridge, and installer without running the Windows test phases. Outputs go to `build/portable/` and `build/release/`. For a portable build without Inno Setup, use `./desktop/build_windows.ps1 -SkipTests -PortableOnly`.

### GitHub source and releases

The GitHub repository ZIP is `build/release/GridForge-v12.0.210-GitHub-Repository.zip`. It contains project source, compiled browser assets (`dist/`), the portable Windows application and bridge under `build/portable/`, PyInstaller spec files, build instructions and workflows, plus the current portable archive under both `release/` and `build/release/`. When a v12.0.210 Windows installer has been built, the packager also places it in `release/`. Extract it to start from a complete repository snapshot, or copy its root contents into a GitHub repository. The `Windows desktop release` workflow can build a Windows installer and portable ZIP from a `desktop-v*` tag or manual Actions run. The workflow performs its configured Windows checks; the local build command below does not run tests. The installer and archive SHA-256 checksums are listed in `build/release/SHA256SUMS.txt` next to the download files. No GitHub repository URL is configured in this working copy.

After the Windows desktop builder has created the portable app folder, regenerate the portable ZIP, clean source archive, GitHub repository ZIP, and checksums with:

```powershell
python scripts/package_source_release.py
```

The script packages the current portable app and bridge, plus source and compiled `dist/`; the GitHub repository ZIP also includes `build/portable/` and the PyInstaller spec files. Temporary work directories, dependencies, and Python caches are excluded. It includes the installer only when the current-version installer exists, and writes SHA-256 checksums for available release artifacts.

## Example workbook

- `examples/GridForge_v11_2_Excel_UI_Chart_Tools_Demo.json` — calculated table column, data bars, dimensions and a combo chart with per-series error bars.
- `examples/GridForge_v11_2_Excel_UI_Chart_Tools_Demo.xlsx` — the example using native OOXML table, chart, conditional-formatting and dimension records.

Rebuild the v11.2 examples with `npm run demo:v112`.

## Verification

```powershell
npm run regression
python -m unittest discover -s tests -p "test_desktop*.py"
python scripts/smoke_formula_ribbon_v115.py
python scripts/smoke_defined_names_v116.py
python scripts/smoke_migration_v110.py
python desktop/gridforge_desktop.py --check
```

Previous versions: [v11.10 feature matrix](docs/FEATURE_VERIFICATION_V11_10.md), [v11.10 release notes](docs/RELEASE_NOTES_V11_10.md), [v11.9 feature matrix](docs/FEATURE_VERIFICATION_V11_9.md), and [v11.9 release notes](docs/RELEASE_NOTES_V11_9.md).



























