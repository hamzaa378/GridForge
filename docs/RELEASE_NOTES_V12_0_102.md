# GridForge v12.0.102

## Drawing selection

- Draw > Select Objects supports Ctrl-click/Shift-click multi-selection, moving the selected objects together, deleting the selection, and applying color/width changes to all selected strokes and shapes.
- Ctrl+D or Draw > Duplicate creates offset copies of the selected drawings.
- Duplicate operations respect the worksheet's drawing-count and point-count limits; interrupted group moves restore the original points.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
