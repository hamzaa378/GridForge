# GridForge v12.0.115

## Sparse MAXIFS and MINIFS ranges

- Local-sheet `MAXIFS` and `MINIFS` now evaluate populated result cells and active spill values only, avoiding dense traversal of large mostly blank result ranges.
- Criteria references must match the result range dimensions; matching and scalar criterion coercion use the existing workbook evaluator behavior.
- Text and logical values in the result range are excluded, and the functions return zero when no numeric value matches.
- Cross-sheet references keep the established evaluator path.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
