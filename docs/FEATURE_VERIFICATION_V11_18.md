# GridForge v11.18 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| Editable slicer connections | Change the compatible Pivot Summaries controlled by a saved slicer from its worksheet card. | TypeScript check/build only; live UI not run |
| Selection propagation | Newly connected summaries inherit the slicer’s current selection; disconnected summaries restore their pre-slicer field filter. | TypeScript check/build only |
| Transaction safety | Invalid connections or failed summary refreshes roll back all affected filters and the connection list. | TypeScript check/build only |
| Compatibility rules | Connections require the same source worksheet/range and the same ungrouped row/column field; a slicer requires its primary summary. | TypeScript check/build only |
| Installer | Desktop and Inno Setup metadata set to 11.18.0. | Installer compile with test phases skipped |

## Remaining gaps

- Slicers remain GridForge worksheet controls; native Excel slicer parts and timelines are not written.
- The Data Model, Power Pivot/DAX, OLAP, calculated fields/items, and full native Excel PivotTable editing remain unsupported.
- No Windows tests, live UI checks, or Excel acceptance checks were run.
