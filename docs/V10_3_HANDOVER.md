# GridForge v10.3 handover / continuation guide

## Release status

- Based on the official `GridForge_v10_2_506_Function_Excel_Parity_Test_GitHub_Ready.zip` source and saved actual Excel results supplied by the user.
- Maintains the **506-function registry**, 400×52 virtualized canvas, multi-sheet document, formula parser, undo/redo, CSV/JSON/XLSX/XLSM, action macros, optional desktop-Excel VBA bridge, 100 Data Studio actions, and dual light/dark themes.
- Repairs 32 of 34 *formerly* unexpected Excel fixture mismatches, with 0 regressions in the 418 original matches. The remaining two cases and four known semantic gaps are documented rather than suppressed.
- Clean Excel-like home interface: light theme by default, Home/File/Data/Tools/View/Formulas/More navigation, Open and Save Excel shortcuts, optional Insights sidebar and metrics.

## Main source responsibilities

- `src/engine/functions_v9.ts`: bond prices, coupons, yields, durations, complex number formatting.
- `src/engine/functions_v10.ts`: chi-square/statistical routines, accurate ERF, Bessel Y, odd-coupon argument guards, some date/financial calculations.
- `src/engine/functions_v6.ts`: normal confidence and 30/360 discount-security calculations, AVERAGEA coercion.
- `src/main.ts`: spreadsheet UI, ribbon, clean/advanced mode, theme and workflow commands.
- `public/styles.css`, `public/index.html`: light-first compact workspace and progressive disclosure.
- `tests/excel_parity_v10_3_snapshot.test.mjs`: permanent Excel baseline regression checks.
- `scripts/verify_excel_snapshot.mjs`: reproduces count comparison against actual user Excel output.
- `scripts/excel_focus_v10_3.py`: optional desktop-Excel targeted recheck for the two remaining cases.

## Commands

Python-only app: `python run_gridforge.py`. No pip modules required for viewing and editing the precompiled application.

Developer commands:

```powershell
npm install
npm run check
npm test
npm run parity:snapshot
npm run benchmark
npm run benchmark:tools
npm run regression
```

For browser testing: `python -m pip install playwright`, install Chromium (`python -m playwright install chromium`), or provide `GRIDFORGE_CHROMIUM` to a local Chromium executable. Tests use an isolated browser profile/port.

## Verification and limits

Read `docs/V10_3_EXCEL_REPAIR_AUDIT.md` for exact numeric claims and unresolved cases. `npm test` and browser regressions are functional checks, not proof of full Excel compatibility. The hosted browser may not support page reload persistence checks. VBA run-through requires a Windows machine with desktop Excel and pywin32; the local bridge security checks do not establish real macro execution.

## Next recommended steps

1. Run `scripts/excel_focus_v10_3.py` and examine the XNPV/IMPOWER Excel outputs.
2. Re-run the complete Excel comparison on Windows, store new evidence separately, and update the audit from measured data.
3. Expand finance fixtures: negative yields, leap years, coupon day-count bases, odd coupon valid and invalid shapes; keep tests based on documented conventions, not expected fixture values alone.
4. Improve workbook fidelity and Excel-native UX incrementally without removing the custom parser or hiding calculation limitations.
