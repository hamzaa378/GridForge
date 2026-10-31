# GridForge v12.0.112

## Sparse COUNTIF ranges

- `COUNTIF` on local references now checks stored cells and active spill values, then accounts for unpopulated cells using the range area and the same criterion matcher.
- This avoids traversing every blank coordinate for common sparse ranges while preserving the existing wildcard, comparison, blank, and case-insensitive matching logic.
- Cross-sheet references retain their existing evaluation path. Referenced errors continue through the current error propagation behavior.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
