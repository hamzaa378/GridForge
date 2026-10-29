# GridForge v11.1 — Continuation and Handover Guide

This is the primary continuation guide for the v11.1 source package. It starts from v11.0 and preserves the existing browser, formula engine, desktop launcher, and Windows installer build kit.

## Architecture and important files

| File | Responsibility |
|---|---|
| src/main.ts | Ribbon, worksheet grid, dialogs, cell interaction, chart/table editors, conditional formatting and dimension gestures |
| src/engine/workbook.ts | Sparse cell storage, formula evaluation, dependency graph, recalculation, undo/redo |
| src/engine/parser.ts | Formula lexer/parser, A1 references, and relative formula shifting |
| src/engine/tables.ts | Table validation, styling, header names, and supported structured-reference resolution |
| src/engine/table_calculations.ts | Calculated-column template propagation, manual override tracking, future-row fill, and metadata reconciliation around cell undo/redo |
| src/engine/table_growth.ts | Contiguous row growth, table overlap checks, filters, and table-linked chart range extension |
| src/engine/charts.ts | Chart model validation, independent axis scales, SVG rendering, and trendline overlays |
| src/engine/trendlines.ts | Numerical trendline fits, forecasts, equation summaries, R², and invalid-domain handling |
| src/engine/sheettools.ts | Per-sheet tables, charts, conditional rules, filters, and row/column dimension metadata |
| src/engine/ooxml_charts.ts | Native DrawingML chart output and chart drawing relationships |
| src/engine/xlsx.ts | OOXML package export/import for worksheets, tables, charts, conditional formatting, and dimensions |
| src/engine/document.ts | Multi-sheet GridForge JSON persistence and compatibility defaults |
| desktop/gridforge_desktop.py | Local desktop wrapper, static file serving, save/open integration |
| desktop/build_windows.ps1 and desktop/installer/GridForge.iss | Windows executable and installer build configuration |
| scripts/regression.mjs | TypeScript, unit, performance, browser, XLSM, and VBA bridge verification stages |
| scripts/smoke_v111_browser.py | Focused Chromium flows for the v11.1 priorities |
| tests/table_calculations_v111.test.mjs | Calculated-column propagation, overrides, row fill, undo/redo |
| tests/v111_features.test.mjs | Trendline references, axis scales, data-bar validation, and OOXML feature records |

## Implemented in v11.1

- Calculated columns fill blank cells and formula cells when a table-body formula is entered. Occupied literal values remain intact and are tracked as manual overrides. New contiguous table rows get the formula unless that row is an override.
- Trendlines are represented in chart series metadata, not only in rendered SVG. The six requested types are calculated in the local engine and serialized in the supported native chart subset.
- Secondary axis assignments are saved per chart series. The SVG renderer calculates separate ranges and renders a right value axis. Native chart output writes a second value axis and the importer restores tested assignments.
- Data bars are live conditional rules. Value ranges are recalculated from current numeric cells; these are not static fills. XLSX uses native conditional-formatting dataBar rules.
- Row and column dimensions are stored as sheet metadata, used by the virtualized grid, and mapped to worksheet column/row dimension records on XLSX export/import.
- Chart Design and Format are contextual ribbon tabs. Selecting a chart exposes type/style, row/column orientation, legend, labels, source-data, title, size, and delete controls in the existing ribbon. Switching worksheets and deleting the selected chart clear the contextual selection.
- Example JSON and XLSX files are under examples/ and can be rebuilt with npm run demo:v111.

## Setup and test commands

Requires Node.js 20+ and Python 3. From the project root:

    npm install
    npm run build
    npm run regression
    python -m unittest discover -s tests -p "test_desktop*.py"
    python scripts/smoke_migration_v110.py
    python desktop/gridforge_desktop.py --check
    npm run demo:v111

The Windows installer script is BUILD_WINDOWS_INSTALLER.bat. Building the installer requires a compatible Windows environment, PyInstaller, Inno Setup 6, and WebView2 testing. The ZIP does not include a compiled installer.

## Verification results and status

See docs/FEATURE_VERIFICATION_V11_1.md for the feature-by-feature matrix, actual regression counts, performance figures, and test status. The current test run includes Chromium-based browser interaction; it does not include real Microsoft Excel automation.

## Known limitations

- The supported structured-reference syntax remains a same-sheet subset. Table expansion happens through normal grid edits/pastes directly below the table.
- Data bars do not have a curated color preset gallery. Native XLSX preserves standard bar color, visibility and fixed/automatic thresholds, but negative-color and gradient parity are limited.
- Advanced chart editing remains incomplete: axis min/max controls, number formats, error bars, minor gridlines, rich labels, and individual point formatting are not complete. The detailed series dialog is still used for per-series axis and trendline settings.
- Row and column dimension updates persist but are not included in the undo/redo history. Multi-selected row/column sizing is not implemented.
- Browser and OOXML tests do not prove how every chart or conditional rule renders in Microsoft Excel.
- This package contains source and installer scripts, not a compiled/installed Windows executable.

## Recommended next development steps

1. Finish worksheet dimension undo/redo and multi-row/multi-column resize, then add keyboard and browser tests.
2. Add data-bar color presets, zero-axis controls, and import/export for supported negative-bar OOXML extensions.
3. Add chart axis bounds/number formats, major/minor gridline settings, custom series colors, markers, and better labels.
4. Run Excel COM acceptance tests for the example workbook, inspect Excel repair warnings, and compare chart axes/trendlines/data bars after save/reopen.
5. Build and install the Windows package on a clean Windows 10/11 x64 machine, then record launcher, save/open, install, upgrade, and uninstall results.
6. Update this feature matrix with new verified scope before making any future compatibility percentage claim.
