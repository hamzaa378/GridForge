# GridForge v12.0.73

## Worksheet name references

- Renaming a worksheet now updates qualified references inside worksheet-scoped defined names throughout the workbook.
- Deleting a worksheet changes matching references inside worksheet-scoped names to `#REF!`, matching formula behavior.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
