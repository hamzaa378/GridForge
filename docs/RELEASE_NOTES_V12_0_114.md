# GridForge v12.0.114

## Sparse SUMIFS and AVERAGEIFS ranges

- Local-sheet `SUMIFS` and `AVERAGEIFS` now evaluate populated result cells and active spill values only, avoiding a full traversal of large blank result ranges.
- Criteria references must have the same dimensions as the result range, and each criterion is evaluated as a scalar with the established blank-to-zero rule.
- Text and logical result values remain excluded from the numeric aggregation; matching referenced errors retain error propagation.
- Cross-sheet references keep the established evaluator path.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
