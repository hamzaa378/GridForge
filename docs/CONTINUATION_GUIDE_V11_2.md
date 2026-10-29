# GridForge v11.2 — Continuation Guide

This is the primary handoff guide for the v11.2 source package. It continues the v11.1 workbook engine and focuses on an Excel 2019-style workspace, chart tools and worksheet selection.

## Main files

| File | Role |
|---|---|
| `src/main.ts` | Ribbon, virtualized grid, dialogs, chart/table/conditional-formatting UI and resize gestures |
| `src/engine/workbook.ts` | Sparse cell storage, formulas, dependencies, recalculation and cell/style history |
| `src/engine/tables.ts`, `table_calculations.ts`, `table_growth.ts` | Table metadata, structured-reference subset, calculated columns and row growth |
| `src/engine/charts.ts`, `trendlines.ts` | Chart model, SVG renderer, axis scaling, error bars and trendline calculations |
| `src/engine/ooxml_charts.ts` | Native chart XML output |
| `src/engine/sheettools.ts` | Per-sheet tables, charts, filters, conditional rules and dimension metadata |
| `src/engine/xlsx.ts` | XLSX import/export for cells, tables, chart parts, conditional formatting and dimensions |
| `src/engine/document.ts` | Multi-sheet document JSON and compatibility defaults |
| `desktop/gridforge_desktop.py` | Local desktop wrapper and file handling |
| `desktop/build_windows.ps1`, `desktop/installer/GridForge.iss` | Windows executable and Inno Setup configuration |
| `scripts/regression.mjs` | TypeScript, unit, performance, browser, XLSM and VBA security stages |
| `scripts/smoke_v111_browser.py` | Table, chart, resize and XLSX browser workflows |
| `tests/v111_features.test.mjs` | Chart math, series options and supported OOXML feature tests |

## Implemented in v11.2

- Excel-style tab order and contextual Table Design, Chart Design and Format tabs.
- Header-based whole row/column selection, multi-selection resizing, and selected row-height/column-width commands in Page Layout.
- Data-bar color presets in addition to custom colors and existing solid/gradient options.
- Per-series error bars: fixed value, percentage and standard deviation; SVG display and supported OOXML persistence.
- Existing calculated columns, secondary axes, trendlines, axis bounds/number format, major/minor gridlines, markers and table-linked charts continue from v11.1.

## Build and checks

From the project root with Node.js 20+ and Python 3:

```powershell
npm install
npm run build
npm run regression
python -m unittest discover -s tests -p "test_desktop*.py"
python scripts/smoke_migration_v110.py
python desktop/gridforge_desktop.py --check
```

The Windows installer build additionally requires Inno Setup 6 and the pinned Python packaging dependencies. Run `BUILD_WINDOWS_INSTALLER.bat` on a Windows x64 build machine. For this v11.2 source, the installer compiled and passed an isolated silent install/uninstall check; real Excel acceptance remains open.

## Limitations and next work

- Worksheet dimension changes are not part of workbook undo/redo yet; AutoFit uses approximate browser font metrics.
- Excel's full chart gallery and advanced chart types are not implemented. Imported complex Excel chart layouts may be simplified.
- Error-bar standard deviation behavior, native Excel rendering and installation must be accepted in real Excel/Windows environments.
- Table formulas implement a documented structured-reference subset, not the full Excel formula grammar.
- XLSX tests check GridForge behavior and selected OOXML parts, not compatibility certification.

See `docs/FEATURE_VERIFICATION_V11_2.md` for results and exact coverage, and `docs/RELEASE_NOTES_V11_2.md` for the user-facing change list.
