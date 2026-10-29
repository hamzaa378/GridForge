# GridForge v10.7 — Release verification

## Verified locally in this build environment

- TypeScript `tsc -p tsconfig.json`: PASS.
- Browser static asset copy `node scripts/copy-static.mjs`: PASS.
- Full Node test suite (`node --test tests/*.test.mjs`): 861 of 861 passed.
- New chart tests: source range limits, two-series formula data, six chart kinds, safe escaped title, negative/missing values, per-sheet JSON persistence, duplicate sheet deep copy, source recalculation, XY scatter, and cell styles with undo/XLSX serialization.
- Desktop wrapper Python unittest: 7 tests passed, including SVG native-save allowlist/safety checks.
- Source-linked benchmark: 10,000 formulas, 26 dependency-affected cells after single source edit; benchmark completed without error (local timings not portable).
- Chromium in-memory test: real browser scripting executed using the compiled JS modules and stylesheet loaded from local files via browser Blob URLs to avoid hosted navigation restrictions. Script opened the spreadsheet with **zero page errors**, rendered Home, switched to Insert, previewed and created a Column chart on `A3:D11`, then switched Home and applied Georgia 18px. Screenshots: `docs/v10_7_home_preview.png`, `docs/v10_7_insert_preview.png`, `docs/v10_7_worksheet_chart_preview.png`.

The environment blocked normal browser `page.goto` navigation by administrative policy. This in-memory UI smoke is useful evidence but does not replace full Windows WebView2/browser acceptance testing. No Windows `.exe` or Inno Setup installer was built in this Linux environment. No live Microsoft Excel run was conducted for charts or formatting.

## Before distribution on Windows

1. Extract to a clean directory; run `python run_gridforge.py --check` and `python run_gridforge.py`.
2. Check Home tab formatting, including preset styles, text/fill colors, wrap and borders; test Ctrl+Z / Ctrl+Y.
3. Select `A3:D11` on the sample, insert each chart type (Scatter is best tested with numeric X/Y columns); drag, resize, edit title/type/source, delete, export SVG.
4. Change sample numeric input `B4`; verify chart re-renders; save JSON, restart and import JSON, verify chart metadata and position persist.
5. Export XLSX and verify numeric values, formulas, font size/family, bold/strike/wrap, and simple borders against desktop Excel. Native Excel chart objects are **not exported**; retain JSON and SVG files as needed.
6. Check workbook import, filters, formatting undo/redo, existing macro templates and the optional VBA bridge opt-in policy.
7. Run `BUILD_WINDOWS_INSTALLER.bat`, launch the portable EXE, install the setup, verify Start menu GridForge/Macro Studio shortcuts, Save As (including SVG), restart, and uninstall.
8. Test on a clean Windows 10/11 x64 machine with WebView2, standard user permissions, offline. Check installed application is code-signed before broad distribution if possible.

## Known limitations

There is no native Excel DrawingML chart export/import. Chart features such as trendlines, 3D, secondary axes, data label placement and waterfall charts remain unsupported. Spreadsheet row and column dimensions are fixed; wrapping does not automatically grow row height. Excel XLSM macro/VBA interoperability limitations from earlier releases remain. 506 function-name parity was measured earlier under one fixture per name; this release does not claim further Excel parity.
