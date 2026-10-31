# GridForge v12.0.119

## Compact dependency safeguards

- Dynamic references avoid adding redundant per-cell edges when the formula already depends on that cell through a compact range.
- Array spill collision checks still detect a spill into a cell referenced by the same formula's compact range, preserving circular-reference reporting.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
