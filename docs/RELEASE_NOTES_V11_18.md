# GridForge 11.18.0 Release Notes

## Pivot slicer connections

- Edit a saved slicer’s Pivot Summary connections from the chain button on its worksheet card.
- The connection picker offers only summaries with a matching source sheet, source range, and ungrouped row/column field.
- Newly connected summaries inherit the slicer’s active selection. Disconnecting a summary restores its saved pre-slicer filter.
- Deleting the slicer restores saved filters; connection changes roll back if validation or any affected summary refresh fails.

## Verification and limits

- TypeScript check and production build passed.
- The Windows installer is compiled with test phases skipped; it is not installed or launched.
- No Windows regression tests, live UI tests, or Excel acceptance tests were run.
- Native Excel slicer controls and timelines, Data Model/Power Pivot/DAX/OLAP, and full Excel PivotTable behavior remain gaps.
