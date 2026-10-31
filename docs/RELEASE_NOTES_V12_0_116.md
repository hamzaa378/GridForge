# GridForge v12.0.116

## Sparse criteria-range error behavior

- Sparse `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS`, and `MINIFS` pre-scan populated cells in their local reference ranges and preserve the evaluator's established error propagation.
- The scan stays proportional to stored cells rather than range area, so mostly empty worksheet ranges remain sparse.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
