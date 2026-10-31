# GridForge v12.0.88

## Freeze panes from the active cell

- Freeze every worksheet row above and column left of the active cell.
- Unfreeze all currently fixed rows and columns from the View tab.
- Retain the existing Freeze row and Freeze col shortcuts. Existing XLSX pane metadata import/export is reused.

## Verification

- TypeScript static validation and Windows release compilation use `-SkipTests`.
- No app launch, live UI, Windows runtime, or Excel GUI tests were run.
