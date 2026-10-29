# GridForge v11.17 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| Shared Pivot Summary slicers | One worksheet slicer can filter up to 16 compatible saved summaries using the same source sheet, source range, and ungrouped row/column field. | TypeScript check/build only; no live UI or Excel acceptance test |
| Slicer update safety | Filter changes and slicer deletion update connected reports together; deletion restores each report’s prior field filter. Failed refreshes roll back to the previous slicer selection. | TypeScript check/build only |
| Connection validation | Rejects missing, incompatible, duplicate, grouped-field, and cross-source links; removes or promotes connections when a linked report or its source worksheet is deleted, or when the report is edited. | TypeScript check/build only |
| Installer | Desktop and Inno Setup metadata set to 11.17.1. | Installer compile with test phases skipped |

## Limits

- GridForge slicer controls are worksheet objects, not native Excel slicer parts. They are not editable as Excel slicers after XLSX export.
- Native timelines, Data Model-backed PivotTables, Power Pivot/DAX, OLAP, and fully interactive Excel PivotCharts remain unsupported.
- No live application, Windows test suite, or Excel UI test was run for this change.
