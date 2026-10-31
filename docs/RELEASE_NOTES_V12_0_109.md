# GridForge v12.0.109

## Conditional maximum and minimum

- `MAXIFS` and `MINIFS` now treat an empty-cell criterion reference as zero, consistent with the other conditional aggregation functions.
- Result and criteria references must now have the same row and column dimensions. Ranges with equal cell counts but different shapes are rejected with `#VALUE!`.
- `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS` now enforce Excel's 127 criteria-pair limit before processing ranges.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
