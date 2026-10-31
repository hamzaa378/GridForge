# GridForge 11.28.0 Release Notes

## Excel-style frontend redesign

- Reworked the main workspace framing around the workbook: compact title and quick-access area, Excel-like tab order, a compact ribbon, formula bar, sheet tabs, and familiar status strip.
- Restored the Draw tab position. Selecting it explains that ink and drawing tools are not available and opens the feature-availability panel.
- Gave the worksheet grid more prominence by hiding diagnostic metrics in the default workbook view; formula insights remain optional.
- Preserved the guided workspace and light/dark themes.

## Compatibility boundary

GridForge is an independent spreadsheet application with an Excel-inspired interface. It does not embed Microsoft Excel. The Draw tab's ink tools are unavailable; Insert → Shapes provides the closest supported alternative. This visual redesign does not change calculation or file-format compatibility.

## Verification scope

- TypeScript type-check and production build completed.
- Windows installer and source/repository archives were compiled with Windows test phases skipped.
- No automated tests, Windows runtime tests, installer launch, live UI checks, or Excel GUI checks were run.
