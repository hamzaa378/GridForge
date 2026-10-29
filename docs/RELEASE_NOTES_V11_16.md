# GridForge 11.16.0 Release Notes

## Sunburst chart

- Added a hierarchical Sunburst chart with proportional rings for each category level.
- Supply one or more label columns followed by the numeric values column. Duplicate paths are combined, values must be non-negative, and invalid input produces an explanation in the chart preview.
- Sunburst is available in the Insert gallery, chart editor and chart-design ribbon.

## Pivot formulas

- `% Of` and `Difference From` calculations can now use a specific item from the selected Base field, alongside Previous and Next. The choice is editable in the Pivot Summary dialog, participates in `GETPIVOTDATA`, and is stored in supported native XLSX PivotTable metadata.
- Pivot source fields now get stable, unique captions: blank headers become `Column n`, and repeated headers receive suffixes such as `(2)`. This keeps field controls and `GETPIVOTDATA` criteria distinct.
- Pivot refresh now reads each needed source cell once per row and uses indexed multi-select filters, avoiding repeated scans of large filter lists during report generation.
- Native PivotTables with no Row fields are now accepted on import, so Values-only summaries can be refreshed and edited in GridForge.
- A source field can now appear on a Row or Column axis and also be summarized in Values, matching common Excel PivotTable layouts; native XLSX import/export keeps the field in both roles.
- Added workbook-aware `GETPIVOTDATA` for generated Pivot Summaries. It resolves value captions, field/item criteria, configured base aggregations, active report/item/caption/date filters, and Top/Bottom value filters.
- Formula dependencies follow the Pivot Summary output region, so refreshes and changes that grow the report invalidate dependent formulas. Saved workbooks rebind and recalculate these formulas when their sheet and Pivot Summary context is restored.
- `GETPIVOTDATA` evaluates the saved Pivot Summary's configured base aggregation and Show Values As mode, including percentages, Index, % Of, running totals, differences, and ranks. It only recognizes saved generated GridForge Pivot Summaries; Excel cache and subtotal combinations are not exhaustively verified.
- Nested row and column groups now have saved expand/collapse controls plus editor actions to collapse or expand all groups. XLSX writes and reads native PivotItem hide-details metadata when labels can be matched; repeated labels under different parents share the native field-item state.

## Installer

- Updated the Windows desktop and Inno Setup versions to 11.16.0.
- Installer compilation skips Windows regression and desktop tests as requested. The build does not install or launch the application.

## Verification and limits

- TypeScript check/build passed for this continuation. The Windows installer was rebuilt with test phases skipped, as requested. The existing automated test count below is historical and does not include this change.
- The earlier 978-test result belongs to a prior source revision.
- Sunburst and other specialized charts export to XLSX as clustered columns until their native chart serializers are implemented.
- No live UI or Windows tests were run.
