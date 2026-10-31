# GridForge v12.0.111

## Sparse range counting

- Local-sheet references passed to `COUNT`, `COUNTA`, and `COUNTBLANK` now use sparse workbook storage plus active spill cells rather than iterating every coordinate in the reference.
- `COUNTBLANK` computes truly unpopulated positions from the range area, then inspects stored cells to account for empty-string formulas and error cells. `COUNTA` continues counting formula results and errors as occupied; `COUNT` counts only numeric results and skips errors.
- Cross-sheet references retain the existing reference-resolution path.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
