# GridForge v10.7 — Single continuation and handover guide

## Project snapshot

- Base: v10.6 macro studio and Windows installer source; no prior macro features intentionally removed.
- Platform: TypeScript ES modules compiled to `dist`, HTML/CSS app; Python standard-library launcher for browser use; `desktop/gridforge_desktop.py` pywebview + PyInstaller and Inno Setup for Windows packaging.
- Spreadsheet: 400 rows × 52 columns (20,800 virtual canvas cells), with 506 local formula names and 10 connection-dependent names. Incremental dependency engine, XLSX import/export with limits, XLSM content-preserving edits with limits, CSV/JSON, multiple sheets, undo/redo, data tools, formatting, offline macro studio.
- NEW: Six offline chart types with SVG output, per-sheet persisted chart specs in GridForge JSON, chart preview, edit/move/resize/delete, downloadable SVG; font family/size, borders, text wrap and strike, cell-style presets; Excel-inspired Insert tab.

## Key code responsibilities

- `src/main.ts`: ribbon tabs, spreadsheet selection, file operations, editing, charts UI, formatting, macros and save. Chart cards are DOM overlays created by `renderCharts()` after virtual grid cells render.
- `src/engine/charts.ts`: range/data extraction, metadata validation, six deterministic chart SVG renderers. Avoid untrusted SVG/HTML interpolation: always escape dynamic labels.
- `src/engine/sheettools.ts`: `SheetTools.charts` persisted per sheet, max 24 with bounds validation. Existing JSON files lacking `charts` are supported via defaults.
- `src/engine/document.ts`: per-sheet JSON and duplicate sheet; no format version bump required because `tools` is version-tolerant.
- `src/engine/format.ts`, `src/engine/workbook.ts`: style validation, formats, history/undo and transactional application.
- `src/engine/xlsx.ts`: XLSX import/export for workbook cells, formatting (font family/size, strike/wrap and basic borders); **does not export GridForge chart drawings**.
- `src/engine/macros.ts`, `scripts/run_macro_offline.mjs`: offline macro engine. v1–v4 JSON recordings supported.
- `desktop/gridforge_desktop.py`: app wrapper + native Save As (now allows static SVG files); no unrestricted filesystem API.
- `desktop/installer/GridForge.iss`, `desktop/build_windows.ps1`, `.github/workflows/windows-desktop-release.yml`: packaging workflows. Build version 10.7.0.
- `tests/charts_v107.test.mjs`, `tests/test_desktop.py`: new calculations, persist/format roundtrip and native saving security.

## Fresh install and tests

```
npm install --no-audit --no-fund
npm run build
npm test
python -m unittest discover -s tests -p 'test_desktop*.py' -v
npm run benchmark
python run_gridforge.py --check
```

User launch without npm: `python run_gridforge.py`, or `START_HERE_WINDOWS.bat`. Installer build (Windows x64 only, Python + Node + Inno Setup): `BUILD_WINDOWS_INSTALLER.bat`. App users do not need developer runtimes; Edge WebView2 is required.

## Architecture/integrity restrictions

- Do not change `src/engine/functions*` or parser semantics just for UI redesign.
- Do not execute VBA or scripts on import; original VBA only through explicit, isolated Windows bridge on trusted workbooks.
- Charts are stored in application JSON, and save to localStorage during normal edits. XLSX export currently omits editable charts; app notifies user and docs disclose this.
- All chart titles, ranges and numeric data must be validated, especially on JSON import. No network requests for chart rendering.
- Resize and drag updates are bounded. Avoid breaking grid virtualization and anchor positions by dynamically varying column widths without redesigning position math.
- Avoid claiming Windows installer or Excel interoperability was tested unless done on Windows/Excel.

## Outstanding work (prioritized)

1. Windows build + WebView2 GUI, installer, shortcut, native SVG file Save As, install/uninstall acceptance.
2. Expand workbook formatting compatibility: dynamic row heights / column widths, merged cells, richer XLSX imported borders/fonts.
3. Optionally implement validated DrawingML chart export to Excel XLSX and import of native charts. This needs Excel/LibreOffice-compatible workbook tests, not just XML checks.
4. Chart improvement: axis labels, legend selection, editing individual series, accessible keyboard manipulation, data-label toggles, PNG export.
5. Extended PivotTable engine and external connections require explicit infrastructure and policy review; do not fake supported functions.
6. Full end-to-end browser tests in the Windows environment (browser navigation is restricted in the hosted runner).

## Verified test evidence

See `docs/V10_7_VERIFICATION.md` for exact counts/results; `docs/v10_7_*_preview.png` for captured in-memory Chromium UI screenshots. Historical Excel 506-function parity results and Macro Studio handover remain in `docs/`. Do not replace them with optimistic coverage claims.
