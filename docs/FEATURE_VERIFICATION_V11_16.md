# GridForge v11.16 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| Sunburst chart | Aggregates parent totals from hierarchy paths; allocates nested annular sectors by relative value; supports up to the worksheet range limit | SVG tests for hierarchy, level aggregation, rendering and negative-input message |
| Chart integration | Sunburst appears in the chart gallery, editor and chart-design ribbon | TypeScript build and source tests |
| XLSX compatibility | Sunburst exports as a clustered-column fallback | OOXML serialization test; no Excel GUI test |
| Installer | Inno Setup and desktop metadata set to 11.16.0 | Compiled with desktop test phases skipped |
| Source verification | TypeScript checks/build and 978 automated tests | No live UI or Windows tests |
| Pivot detail drill-through | Double-click a generated value, subtotal, or grand-total cell, or use its cell context menu, to create a Details worksheet from the matching filtered source rows with source formatting | TypeScript checks/build only for this change; no runtime or Windows tests |
| Pivot nested row/column expand/collapse | Saved group paths hide/show descendants and retain selected state across refresh/import; editor provides Collapse all/Expand all; XLSX writes/reads native item hide-details flags where labels resolve | TypeScript check/build only; no live UI or Windows tests |
| Pivot slicer responsiveness | Slicer value lists reuse a workbook-revision cache and are recomputed after source value changes or formula invalidation | TypeScript check/build only; no runtime or UI timing measurement |

## Limits

- Excel encodes collapsed state on a field item, so duplicate labels under different parents may share the expanded/collapsed state. Native state writing/import is not validated in a live Excel session.

- Input is a headered rectangle: one or more leading hierarchy columns and a final non-negative numeric value column. Duplicate paths are aggregated.
- Labels may be omitted from small sectors; the SVG title remains available for the chart.
- Sunburst exports to XLSX as a clustered-column chart; native chart semantics are not preserved.
- The installer was compiled but not installed or launched.
