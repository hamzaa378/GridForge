# GridForge 11.20.0 Release Notes

## PivotTable category visibility

Pivot Summary settings now include **Show items with no data**. When enabled, row and column categories observed in the source remain in the report if report filters exclude their records. Date, item, and label filters continue to constrain the categories. Values at empty intersections use the existing missing-value caption behavior.

Supported XLSX export writes the native PivotField `showAll` attribute for row and column fields, and import restores the setting. Excel desktop acceptance was not performed. The Data tab Refresh All command also refreshes saved queries across worksheets and Pivot Summaries throughout the workbook, including when there are no saved queries. `Ctrl+Alt+F5` invokes the same workbook-wide refresh from the worksheet.

## Verification

- `npm run check`
- No automated tests, Windows tests, live UI checks, installer launch, or Excel GUI validation were run.
