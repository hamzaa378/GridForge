# GridForge v12.0.137 — Merged regions and faster sheet navigation

- Added Merge & Center and Unmerge controls for selected cell ranges. Non-anchor values are cleared only after confirmation; overlapping merges and merges exceeding the worksheet coverage limit are rejected. Merge and unmerge operations participate in undo/redo.
- XLSX and ODS import/export now preserve supported merged regions. ODS continues to save calculated values rather than formulas and retains supported styles and sheet dimensions.
- Rendering caches cumulative row and column positions and uses binary search to locate cells during scrolling, avoiding repeated full-prefix scans on large sheets.
- `npm run build` completed. The Windows installer and release archives were rebuilt with test phases skipped. Automated tests, live UI checks, app launch, and Excel/LibreOffice GUI checks were not run.
