# Compare GridForge directly against Microsoft Excel on Windows

## Requirements

- Windows 10/11
- Installed, licensed **Microsoft Excel desktop** (prefer Microsoft 365 with the current formula set)
- Python 3
- `pywin32` (`py -3 -m pip install pywin32`) to automate desktop Excel
- Optional Node.js 20+ to rebuild and regenerate GridForge results. For the provided release, compiled code and GridForge results are already bundled.

**The tests do not open your personal spreadsheets or run your VBA macros.** The runner creates a new unsaved workbook containing only the 140 bundled deterministic formula probes, disables Excel events and automatic macro execution, recalculates, records results, and closes the workbook without saving.

## Simple method: Python only

From the extracted GridForge project folder, in PowerShell:

```powershell
py -3 -m pip install pywin32
py -3 scripts\excel_parity_desktop.py
py -3 scripts\excel_parity_compare.py
```

Or double-click `RUN_EXCEL_PARITY_WINDOWS.bat` after installing pywin32.

The first script **uses actual desktop Excel** and creates `docs/PARITY_EXCEL_RESULTS.json`. The second compares those results with the bundled GridForge results (`docs/PARITY_GRIDFORGE_RESULTS.json`) and writes:

- `docs/PARITY_COMPARISON.md` — human-readable matches and mismatches
- `docs/PARITY_COMPARISON.json` — machine-readable case-by-case results

The comparison command exits with status 1 if there are confirmed scalar differences. Status 2 means required Excel evidence is missing.

## Development workflow

After changing the TypeScript engine:

```powershell
npm install
npm run parity:gridforge
npm test
npm run benchmark
py -3 scripts\excel_parity_desktop.py
py -3 scripts\excel_parity_compare.py
```

`npm run parity:compare` is an alternative Node.js comparison command.

Recompute GridForge results **after every formula-engine change**, so a stale snapshot is not compared to a newer build. Never manually invent `PARITY_EXCEL_RESULTS.json`.

## How agreement is calculated

- Numbers: absolute/relative tolerance of `1e-8`.
- Text, booleans and blanks: exact typed equality.
- Errors: compare Excel-style error codes; meaningful custom error types may differ.
- Excel versions that do not recognize a function are shown separately as unsupported.
- The four recorded dynamic-array limitations are **excluded** from the scalar agreement denominator and shown as known gaps.
- Only comparable, deterministic scalar cases contribute to the quoted case-match percentage. The result is **not** a percentage of Excel functions or all workbook features implemented.

## Interpreting results

If the comparison reports differences, look first at rounding, text-versus-number coercion, wildcard criteria, dates, and error types. Record the exact Excel version and reproducible formula. For any new Excel feature to be called compatible, create more cases around missing values, invalid arguments, locale, references, nested expressions, and large ranges.

## Known limits

This is a first 140-case cross-engine harness, not a certification suite for all 506 registered functions. It does not cover real VBA execution, pivot tables, charts, formatting equivalence, external connections, performance equivalence, or dynamic-array spill footprints. For these use a separate acceptance matrix and representative real workbooks.

Do not use a `.xlsm` from an unknown source to test macros. VBA can perform arbitrary desktop actions.
