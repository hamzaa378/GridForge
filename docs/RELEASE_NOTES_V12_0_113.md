# GridForge v12.0.113

## Sparse COUNTIFS ranges

- Local-sheet `COUNTIFS` references now visit populated cells and active spill results, then account for matching unpopulated cells using the shared reference dimensions.
- Preserves the existing criterion matcher for blank, comparison, wildcard, and case-insensitive criteria; validates the 127-pair limit and matching range shapes.
- Cross-sheet references retain the established evaluation path. Error-valued cells do not match criteria in the sparse local path.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
