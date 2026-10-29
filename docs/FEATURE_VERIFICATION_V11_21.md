# GridForge v11.21 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| Per-field PivotTable “Show items with no data” | Each selected row/column field has an independent control; visible observed hierarchy paths survive report filtering when enabled. Item, date, and label filters continue to constrain paths. Empty intersections use the configured missing-value caption. | TypeScript check/build only |
| XLSX PivotTable interoperability | Writes `showAll="1"` only for enabled row/column PivotFields and imports each field setting independently. Legacy v11.20 global settings normalize to all currently selected axis fields. | Code-only; Excel GUI acceptance was not run |
| Installer | Desktop and Inno Setup metadata updated to 11.21.0. | Installer compile with Windows test phases skipped |

## Limits

- No Windows tests, live UI checks, or Excel acceptance tests were run.
- The report engine preserves observed source paths; it does not synthesize every possible nested parent/child combination or restore members absent from the source cache.
- Native Excel slicers/timelines, Data Model, DAX/OLAP, calculated fields/items, multi-level date hierarchies, manual grouping, and full Excel pivot-cache editing remain unsupported.
