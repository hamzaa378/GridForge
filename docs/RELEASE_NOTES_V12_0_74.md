# GridForge v12.0.74

## Worksheet-scoped defined names

- Names can refer to supported formulas and constants in addition to cell/range references and LAMBDAs.
- XLSX import retains formulas that refer to other worksheets instead of skipping them as unsupported ranges.
- Formula dependency tracking uses the resolved references in these names so dependent cells can update when source cells change.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
