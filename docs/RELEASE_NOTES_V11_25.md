# GridForge 11.25.0 Release Notes

## Formula and calculation compatibility

- `INDIRECT` now resolves A1 and R1C1 cell/range text, quoted sheet names, and supported worksheet-local defined names. Dynamic dependencies are tracked for local and cross-sheet ranges so edits refresh dependent formulas.
- `OFFSET` now returns one cell or a rectangular range, supports cross-sheet references and omitted height/width arguments, and refreshes formulas when any cell in the returned area changes.
- `ADDRESS` now supports all four absolute/relative modes, A1 or R1C1 output, and an optional worksheet prefix.
- `CELL` accepts an omitted reference and supports cross-sheet contents/type/address information, common format codes, prefix and protection information. `ISFORMULA` and `FORMULATEXT` now inspect references on other worksheets.
- `SUBTOTAL` and supported `AGGREGATE` options recalculate after filter visibility changes and correctly distinguish filter-hidden rows from manually hidden rows. The visibility lookup uses a bounded cache.
- Date helpers now recognize Excel's fictitious 1900-02-29 serial in additional weekday, workday, and date conversion paths; `WEEKDAY` supports return types 11–17 and `WEEKNUM` supports Excel's documented week-start types and ISO week numbering.

## Chart improvements

- Scatter charts now plot all Y series against the first numeric X series, use numeric scales for both axes, and expose horizontal-axis bounds, major unit, and number format controls. Series colors, marker visibility, point labels, and legend colors are applied per series.
- Scatter axis bounds, units, titles, and number formats import/export through standard XLSX DrawingML chart axes.
- The SVG chart preview now applies supported Excel custom number formats to axis tick labels, including custom X-axis formats for scatter charts.
- Pie/doughnut charts support first-slice rotation, editable doughnut-hole percentage, per-slice colors and configurable label positions; supported settings and RGB point colors round-trip through XLSX.
- Chart editor changes preserve advanced axis settings that are edited in the chart ribbon.

## Compatibility boundaries

GridForge remains an independent spreadsheet implementation. This release does not implement full Excel formula, locale, numeric precision, chart-style, PivotTable, Power Query M, Data Model, DAX, OLAP, or connected-service parity. Custom axis formats are rendered for the subset understood by GridForge's format renderer; unsupported number-format syntax is preserved for XLSX export but falls back in the preview. R1C1 `INDIRECT` supports worksheet cell/range addresses, not external-workbook references. See [the v11.25 feature matrix](FEATURE_VERIFICATION_V11_25.md) and [the Excel compatibility audit](EXCEL_COMPATIBILITY_AUDIT.md).

## Verification scope

- `npm run check` passed (TypeScript no-emit).
- Automated tests, Windows tests, application launch, live UI checks, and Excel GUI import/export checks were not run.
- The Windows installer compiled with `-SkipTests`; source and portable package checksums are recorded in `build/release/SHA256SUMS.txt`. The installer was not launched.
