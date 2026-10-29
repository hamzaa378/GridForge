# GridForge 11.19.0 Release Notes

## PivotTable empty cells

- Missing row/column intersections now display blank by default instead of a misleading numeric zero.
- The Pivot Summary dialog includes **For empty data cells, show** for a replacement such as `0` or `N/A`.
- The custom replacement is written to native PivotTable `showMissing`/`missingCaption` metadata and restored on import.

## Verification

- TypeScript check and production build passed.
- The Windows installer is compiled with test phases skipped and is not launched.
- No Windows tests, live UI checks, or Excel GUI validation were run.
