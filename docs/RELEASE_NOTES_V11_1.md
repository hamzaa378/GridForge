# GridForge 11.1.0 Release Notes

## Added

- Excel-style calculated table columns with structured current-row formulas, literal-value overrides, future-row fill, JSON metadata, native table formula metadata, and cell formula undo/redo.
- Per-series secondary chart axes with a separate right-side value scale and title.
- Six selectable trendline types, polynomial order and moving-average period, forecasts, equation/R² display options, source-linked redraw, and supported native chart serialization.
- Live conditional data bars with solid/gradient fills, positive and negative GridForge colors, automatic/fixed value thresholds, and optional cell-value display.
- Draggable row/column boundaries, double-click AutoFit, manual heading context-menu dimensions, and JSON/XLSX dimension persistence.
- Contextual Chart Design and Format tabs use the existing Excel-style ribbon for type/style, row/column orientation, legend, labels, Select Data, title, size, and delete actions.
- A combined example workbook and v11.1-specific browser, mathematical, metadata, and OOXML tests.

## Changed

- Package and installer sources now identify as version 11.1.0.
- Desktop source-asset tests now explicitly read UTF-8, independent of the host Windows code page.
- Browser regression opens the File ribbon before JSON export, matching the visible Excel-style command flow.

## Verification boundaries

The feature matrix and actual run results are in docs/FEATURE_VERIFICATION_V11_1.md. The source ZIP contains installer build scripts; the Windows installer is delivered separately. The installer was built and smoke-installed/uninstalled on this Windows machine. Real Microsoft Excel acceptance has not been run.
