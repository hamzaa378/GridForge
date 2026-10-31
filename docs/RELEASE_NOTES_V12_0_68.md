# GridForge v12.0.68

## Page Layout and Excel workbook compatibility

- Page Setup adds **Center horizontally** and **Center vertically** options.
- The settings are validated, saved with GridForge documents, and preserved through XLSX `printOptions` import/export.
- Older documents without these options default to both off.
- The browser print engine does not reproduce Excel's exact page centering behavior; the options are retained for Excel workbook use.

## Verification

- TypeScript check and Windows installer/portable compilation are recorded in the release manifest.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
