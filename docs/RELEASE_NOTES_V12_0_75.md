# GridForge v12.0.75

## Resilient XLSX defined-name import

- Worksheet-scoped defined names are validated one at a time during import.
- A malformed or unsupported name now generates its own warning while supported names on that worksheet are retained.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
