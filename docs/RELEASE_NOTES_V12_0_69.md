# GridForge v12.0.69

## Workbook defined names

- Name Manager can create and edit names scoped to the active worksheet or the entire workbook.
- Worksheet-scoped names take precedence when both scopes use the same name.
- Workbook names can hold formulas, LAMBDAs, constants, or sheet-qualified cell/range references and are available to formulas throughout the document.
- Workbook names persist in GridForge JSON and round-trip through XLSX workbook defined-name metadata.
- Renaming a worksheet updates workbook-scoped references; formulas still use GridForge's supported evaluator behavior.

## Verification

- TypeScript check and Windows installer/portable compilation are recorded in the release manifest.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
