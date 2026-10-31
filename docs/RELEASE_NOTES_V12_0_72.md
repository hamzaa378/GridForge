# GridForge v12.0.72

## XLSM workbook names

- Macro-preserving XLSM export now merges workbook-scoped defined-name additions, edits, and deletions into the original workbook metadata.
- Unchanged or unsupported workbook names remain untouched, and VBA project payloads are still preserved without execution.
- Changed defined names retain comments and hidden flags.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
