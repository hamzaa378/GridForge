# GridForge v10.8 — Release verification record

## Baseline

Based on the GitHub-ready GridForge v10.7 source (Excel-style formatting + SVG worksheet charts + offline Macro Studio). Existing macros, 506 local functions, 10 recognized connection-dependent functions, and Windows installer build scripts were retained.

## Completed in this environment

- **876 / 876 Node automated tests PASS**, including all existing 861 tests and 15 new native chart test subcases.
- **7 / 7 Python desktop unit tests PASS.**
- TypeScript `npm run build` PASS; precompiled `dist/` copied.
- `python run_gridforge.py --check` PASS with Python-only launcher.
- `python scripts/desktop_browser_smoke.py` PASS (native Save As XLSX handoff).
- `python scripts/smoke_native_charts_v108.py` PASS (new chart UI, live worksheet insertion, importing nine chart objects and re-exporting nine native chart parts; no Chromium JS errors).
- All nine chart types serialized into OOXML parts with valid XML, referenced DrawingML relationships and correct ZIP checksums. Independent `xml.etree.ElementTree` parse PASS.
- 10,000-formula selective recalculation benchmark PASS; small source edits recalculated 26 dependent cells.

## Limitations not passed or not yet executed

- **The Windows installer EXE has NOT been built on this Linux build machine.** Build via `BUILD_WINDOWS_INSTALLER.bat` on Windows; then verify installation and launch.
- **Native Excel acceptance is NOT yet run.** On Windows, open `examples/native_charts_v108.xlsx` and execute `python scripts/verify_excel_charts_windows.py` with `pywin32` installed. Report will be `docs/V10_8_REAL_EXCEL_CHART_RESULTS.json`.
- Browser smoke is an in-memory Chromium evaluation; it is NOT equivalent to WebView2 runtime testing on Windows.
- Unsupported chart types, PivotCharts, combo/secondary axes, external chart data, full formatting fidelity, and chart updates inside macro-preserving `.xlsm` are unimplemented.

## Source and test map

`src/engine/ooxml_charts.ts` creates native chart XML, drawing anchors and chart relationship parts. `src/engine/xlsx.ts` wires these into worksheet import/export. `src/engine/charts.ts` validates chart metadata, extracts formula results and renders nine offline SVG previews. `src/main.ts` and `public/styles.css` define chart gallery and editor. Tests: `tests/charts_native_v108.test.mjs`, `scripts/smoke_native_charts_v108.py`, Windows-only `scripts/verify_excel_charts_windows.py`.

## Windows commands

```powershell
npm install
npm run build
npm test
python -m unittest discover -s tests -p 'test_desktop*.py' -v
python scripts\verify_excel_charts_windows.py
.\BUILD_WINDOWS_INSTALLER.bat
```

Keep v10.8 chart compatibility evidence separate from earlier v10.7 report; never claim complete Microsoft Excel compatibility based solely on XML structural checks.
