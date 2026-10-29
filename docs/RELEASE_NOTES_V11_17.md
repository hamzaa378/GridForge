# GridForge 11.17.1 Release Notes

## Pivot Summary slicers

- A slicer can now control multiple compatible Pivot Summaries on the same worksheet.
- Compatibility requires the same source sheet and exact source range, plus the same ungrouped row/column field.
- Filtering and clearing a slicer updates each connected report; failures restore the previous filter selections. Deleting a slicer restores the field filters that existed before the slicer was added.
- Editing or deleting a connected Pivot Summary, or deleting its source worksheet, removes or promotes its slicer connection so saved references remain valid.
- The Pivot Summary editor lists compatible reports when creating a slicer.

## Verification and limits

- TypeScript check and production build passed.
- The 11.17.1 Windows installer is compiled with test phases skipped; it is not installed or launched.
- No Windows regression tests, live UI tests, or Excel acceptance tests were run.
- GridForge slicers remain custom worksheet controls. Native Excel slicer objects, timelines, Data Model, Power Pivot/DAX, OLAP, and the full native PivotTable field panel are still gaps.

