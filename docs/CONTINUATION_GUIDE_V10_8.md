# GridForge v10.8 — Developer continuation guide

**Read this first** when resuming work. Project root: the folder containing `package.json`, `src/`, `dist/`, `desktop/`, `tests/`, `scripts/`. This is the continuation of the v10.7 Excel-style UI and charts release, not a separate spreadsheet app.

## What works now

- Browser and Windows-desktop packaging source; precompiled browser assets for Python-only launch.
- Workbook with 506 offline formula functions, 10 connection-dependent recognized names, Excel-like ribbon, formatting, worksheet virtualization, offline Macro Studio.
- Nine SVG chart types from worksheet data; editable selection range, legend, titles, value labels; chart metadata lives in per-sheet `tools.charts`.
- XLSX export writes editable DrawingML chart objects (chart XML + drawing XML + relationships + content-type entries). Formula values and series references are cached at export.
- XLSX import recognizes supported chart types with simple contiguous local data ranges and drawing anchors, converting them back to GridForge chart metadata.
- Chart JSON metadata validation, size limits, XML escaping, external relationship exclusion, no automatic macro execution.

## File responsibilities

- `src/engine/charts.ts`: chart type union, data extraction, SVG rendering, metadata validation.
- `src/engine/ooxml_charts.ts`: OOXML chart-space serialization, drawing anchors and relationships.
- `src/engine/xlsx.ts`: chart OOXML parts, Xlsx chart import, normal workbook cell import/export.
- `src/engine/sheettools.ts`: per-sheet chart persistence, validation and caps.
- `src/main.ts`, `public/styles.css`: Insert gallery, chart editor, worksheet overlays.
- `tests/charts_native_v108.test.mjs`: serialization and data linkage regression.
- `scripts/smoke_native_charts_v108.py`: Chromium UI and import/export round-trip.
- `scripts/verify_excel_charts_windows.py`: live Excel COM test (Windows-only; pending).
- `desktop/installer/GridForge.iss`, `desktop/build_windows.ps1`, `BUILD_WINDOWS_INSTALLER.bat`: installer builder.

## Recommended next technical work

1. **Run live Excel tests and review the report first.** Correct schema or drawing behavior found in real Excel before adding more chart types.
2. Implement Excel Chart Design tab with series-by-series selection, row/column orientation, titles, styling and theme palette controls.
3. Add native Excel combo charts and secondary axes; extend data model to noncontiguous and cross-sheet series references. Keep limits and validation.
4. Add 100% stacked, histogram, waterfall and stock charts with substantive tests. Avoid name-only support.
5. Consider native DrawingML chart preservation when modifying `.xlsm` packages, without executing or damaging VBA binaries.
6. Test install/uninstall, WebView2, workbook autosave, XLSX chart save/open and native macro replay on clean Windows.

## Reproducibility

```bash
npm run build
node --test tests/*.test.mjs
python scripts/smoke_native_charts_v108.py
python scripts/desktop_browser_smoke.py
python run_gridforge.py --check
```

Windows: `python scripts\\verify_excel_charts_windows.py` followed by `BUILD_WINDOWS_INSTALLER.bat`. **Do not claim genuine EXE built or actual Excel parity until those run successfully.**
