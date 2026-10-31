# GridForge v12.0.71

## Defined-name metadata

- XLSX import and export preserve Excel's `hidden` flag for workbook-scoped and worksheet-scoped defined names.
- Name Manager displays the hidden status and retains it when editing an existing name.
- Macro-preserving XLSM worksheet-name edits retain the hidden flag for changed names.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
