# Native PivotTable slicer support

GridForge now imports and exports a bounded subset of native Excel worksheet PivotTable slicers.

## Included

- Native slicer cache and view parts, workbook and worksheet relationships, content types, and drawing anchors.
- Slicer captions, placement, field binding, and selected item state.
- Native slicer round-trip back into GridForge's editable slicer cards.
- Shared native PivotCaches for PivotTables with matching source sheet/range, group definitions, and grouped-field sort order.
- Connections to all referenced PivotTables that share the slicer's cache.

## Limits

- A slicer connected in GridForge to PivotTables with different native cache definitions is written against its primary PivotTable cache. Other GridForge connections remain in GridForge's own model and are not claimed as native Excel connections.
- Timeline parts, OLAP/Data Model slicers, full Excel cache editing, and live Excel acceptance are not included.
- Generated OOXML parts are parsed and covered by code-level round-trip tests. No live Excel or Windows UI test was run for this change.
