# GridForge 11.21.0 Release Notes

## Per-field PivotTable item visibility

Show items with no data is now configured independently on each selected row or column field in the Pivot Summary editor. Enabling it preserves observed source paths when report filters exclude records; item, date, and label filters still constrain the paths. Empty intersections use the existing missing-value caption behavior.

XLSX export writes the native `showAll` flag on enabled PivotFields, and import restores each field separately. Older GridForge global settings migrate to all selected row and column fields. Excel desktop acceptance was not performed.

## Verification

- `npm run check`
- `npm run build`
- No automated tests, Windows tests, live UI checks, installer launch, or Excel GUI validation were run.
