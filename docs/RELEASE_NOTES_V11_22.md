# GridForge 11.22.0 Release Notes

## Manual PivotTable grouping

The Pivot Summary editor now accepts manual field groups using one line per group: `Group caption = item | item`. Prefix literal `=`, `|`, or `\` with `\`, and write `\\` for one literal backslash. Prefix an edge space with `\` to preserve it. Use `(blank)` to assign empty source cells. Items not listed in a group remain visible as their original captions. Captions and item assignments are validated before generating the summary. Group mappings are indexed while filter options are built, and grouped item lists are cached by source and grouping definition. Workbook replacement, source-sheet changes, edits, paste, sorting, query refresh, and undo/redo clear the caches; sorting also refreshes Pivot Summaries so source-order sorting stays current.

Manual groups flow through nested row or column summaries, supported field filtering, GETPIVOTDATA criteria, and native XLSX PivotTable cache output using discrete grouping properties. The importer restores group captions and source-item assignments for supported local worksheet PivotTables; PivotTable cache sources that name a worksheet table or a local defined range are resolved to that range. Excel grouping cache variants have not been exhaustively verified.

## PivotTable compatibility and export fixes

Legacy saved PivotTables that use the single `filterField` / `filterValue` form now validate and refresh correctly, including after being normalized to the newer filter list. Caption matching for report filters is case-insensitive. Native XLSX export now builds label-filter item visibility without referencing a value before initialization, fixing export of PivotTables with caption filters.

When a PivotTable source contains formulas linked to another worksheet, edits to that upstream worksheet now expose the dependent recalculated cells to the workbook layer. GridForge refreshes only PivotTables whose source ranges overlap those cells, and serializes these background refresh passes so later edits are not left showing an older result.

Charts created from a Pivot Summary now have a code-only regression check for their native XLSX `pivotSource` and PivotTable `chartFormats` metadata. Excel field-button behavior still requires Excel-side verification.

Top and Bottom value filters now remain authoritative when “Show items with no data” is enabled. Labels rejected by the value filter are no longer reintroduced by the empty-item visibility pass, including within nested row and column groups.

The native XLSX round-trip check exposed and fixed a saved-filter normalization defect: a single multi-select report filter no longer gets an incomplete legacy `filterField` alias. Exported PivotTables with report filters now import back as editable GridForge Pivot Summaries and recalculate after source edits in the code-only round-trip check.

Manual-group imports now omit pass-through cache groups whose caption is identical to their only source item. The mapped groups still import intact, while ungrouped items remain ungrouped for features such as slicer eligibility. Code-only coverage now includes manual-group and Top N round-trips.

Workbook validation now rejects slicers connected to manually grouped fields, matching the Pivot Summary editor and the documented requirement that GridForge slicers use ungrouped row or column fields.

Pivot date grouping and date-range controls now use the Excel 1900 serial boundary correctly for early dates, including the special serial 60 label (1900-02-29). This prevents January/February 1900 dates from shifting and keeps March 1900 serials aligned.

Creating, updating, single-report refresh, workbook-wide Refresh All, refreshes triggered by cell edits, slicer filtering/connections, row and column group toggles, collapse/expand-all, and PivotTable detail extraction now calculate from serialized workbook snapshots in module workers. Edit, undo/redo, paste, import, and query-driven refresh requests are debounced and serialized; stale workbook or PivotTable-definition results are discarded and queued again against the latest state. Slicer and group changes preserve rollback behavior if output validation fails.

Workbook-wide refresh now calculates up to four summaries per worker snapshot, deserializing the workbook once per batch instead of once for every summary. The worker applies each result to its private snapshot before calculating the next report, preserving dependent-summary order. If a summary cannot be applied, later reports in that batch are recalculated from the live workbook before applying.

Pivot source validation now accepts ranges up to the supported worksheet capacity of 1,000 rows by 104 columns (104,000 cells). The previous 12,000-cell ceiling unnecessarily rejected larger valid sources. A worker-job regression check covers a 13,000-cell source with 999 records and verifies its grouped result.

The XLSX importer now identifies native ChartEx by its series layout, so Waterfall, Pareto, Box-and-whisker, Funnel, Treemap, and Sunburst exports round-trip back to editable GridForge charts. This also updates the chart regression checks to cover the native export and import path.

## Verification scope

- `npm run check`
- `npm run build`
- `npm test` (all Node code tests)
- Windows installer compilation with test phases skipped
- Focused code-only PivotTable checks (`features_v112` and `pivots_native_v118`)
- 29 focused assertions, including larger-than-12,000-cell worker source coverage
- Background Pivot Summary job checks against serialized workbook snapshots
- Node XLSX PivotTable export/import/refresh round-trip using a test-only XML parser
- No Windows runtime tests, live UI checks, installer launch, or Excel GUI validation
