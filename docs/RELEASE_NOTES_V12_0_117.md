# GridForge v12.0.117

## Compact local range dependencies

- Local worksheet ranges in formulas are stored as rectangular dependencies instead of expanding into one dependency entry per cell.
- Column indexes find formulas affected by edits inside a referenced range; transitive impact previews use the same range-aware dependency lookup.
- The 20,000-cell formula-range limit remains. Cross-sheet range dependency handling is unchanged.
- Formula inspection lists compact range references and returns a bounded sample of referenced cell values.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
