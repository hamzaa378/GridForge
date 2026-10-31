# GridForge v12.0.107

## Conditional aggregation improvements

- `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS` now treat a criterion argument sourced from an empty cell as zero. A typed empty string remains a text criterion.
- `SUMIF` and `AVERAGEIF` now align an explicit result range to the criteria range's row and column dimensions, starting at the result range's top-left reference. This covers result ranges that are smaller or larger than the criteria range.
- Invalid references and referenced formula errors continue to propagate through the workbook evaluator.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
