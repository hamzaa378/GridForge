# GridForge v12.0.76

## Expanded worksheet canvas

- Worksheet bounds now extend to 10,000 rows × 256 columns, while the default blank sheet stays 400 × 52.
- Cell data and styles remain sparse. Each worksheet is capped at 500,000 populated cells and 500,000 styled cells; XLSX import is capped at 250,000 cell records per sheet.
- CSV import and PivotTable source validation use the expanded bounds while retaining their own import and range limits.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
