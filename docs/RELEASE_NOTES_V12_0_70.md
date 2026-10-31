# GridForge v12.0.70

## Workbook name references

- Deleting a worksheet now rewrites references to that sheet inside workbook-scoped defined names to `#REF!`.
- This matches the existing handling for ordinary cell formulas and keeps the broken reference visible and editable in Name Manager.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
