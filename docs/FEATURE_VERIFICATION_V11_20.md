# GridForge v11.20 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| PivotTable “Show items with no data” | Optional setting retains observed row and column category paths when report filters remove their records. Item, date, and label filters still limit the category set. Empty value intersections use the configured missing-value caption. | TypeScript check/build only |
| XLSX PivotTable interoperability | Writes `showAll="1"` for row/column PivotFields and imports the setting when present. | Code-only; Excel GUI acceptance was not run |
| Refresh All | Data tab action and Ctrl+Alt+F5 refresh saved queries across worksheets and all Pivot Summaries, including when the workbook has no saved queries. | TypeScript check/build only |
| Installer | Desktop and Inno Setup metadata updated to 11.20.0. | Installer compilation without Windows test phases |

## Limits

- No Windows tests, live UI checks, or Excel acceptance tests were run.
- Show items with no data is a workbook-backed subset: one setting applies across row and column fields and preserves observed source paths. It does not synthesize every possible parent/child combination or restore members absent from the source cache.
- This remains GridForge’s worksheet-backed Pivot Summary model, with a supported native XLSX PivotTable subset. Native slicer/timeline parts, Data Model, DAX/OLAP, calculated fields/items, and full Excel pivot-cache editing remain unsupported.
