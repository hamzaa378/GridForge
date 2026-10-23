# GridForge Excel Compatibility Audit — v10.1

## Executive assessment

**GridForge is a local, browser-based, Excel-inspired spreadsheet editor, not a complete Microsoft Excel replacement.** It has 506 registered Excel-style function names, independently implemented formula parsing, XLSX/XLSM interoperability and spreadsheet interactions. Function registration alone does not establish Excel behavior equivalence.

There are **two distinct levels of evidence**:

1. **GridForge-only verification (performed):** 140 deterministic formula probes, of which 136 have scalar expected values and four document known dynamic-array limitations. After three targeted classes of fixes, the 136 scalar probes match their manually specified expected values. This is *not* a result of running Microsoft Excel.
2. **Live Excel differential verification (not performed here):** The package contains a Windows Excel COM test runner. When executed on a machine with desktop Excel, it computes the very same formulas and compares result types, numbers, booleans, strings and errors, producing `docs/PARITY_COMPARISON.md` and JSON evidence. Do not claim the two engines agree until this runner is executed.

## Repeatable tests available now

| Check | Current evidence | Meaning |
|---|---|---|
| TypeScript | Pass | Builds and type-checks |
| Automated tests | 816 / 816 pass | Existing engine/document tests + 140 regression probes + fixture integrity |
| GridForge scalar reference probes | 136 / 136 match | Expected examples, *not* live Excel |
| Known dynamic-array gaps | 4 explicitly recorded | Current engine returns `#VALUE!` where Excel may spill a matrix |
| Browser UI regression | 30 / 30 pass | In an isolated Chromium browser; storage reload skipped |
| 10,000-formula benchmark | Pass | Selective recalculation; not a throughput guarantee |
| 10,000-cell Data Studio benchmark | Pass | Preview/apply/undo and unrelated-cell checks |
| Python-only launch | Pass | Compiled frontend served using Python standard library |
| XLSM package preservation | Pass | Synthetic VBA bytes unchanged during supported cell edits |
| VBA bridge security validation | Pass | Authentication, origin, names, XLSM preflight |
| Actual Excel formula comparison | **Not measured** | Needs Windows + Microsoft Excel desktop + pywin32 |
| Actual VBA execution | **Not measured** | Needs Excel desktop, trusted test macro and opt-in |
| Excel workbook object-model compatibility | **Not measured** | Includes pivot tables, print setup, protected ranges and chart parts |

## Confirmed defects repaired during the audit

1. **Negative half rounding:** `ROUND(-1.5,0)` previously returned `-1`, but Excel returns `-2`. Core rounding now uses symmetric, away-from-zero tie behavior.
2. **Typed literal coercion in aggregations:** `SUM("3",4)` previously returned `4`, and `SUM(TRUE,2)` returned `2`. The parser evaluator now distinguishes directly typed literal values from referenced cells for `SUM`, `COUNT`, `AVERAGE`, `PRODUCT`, `MIN` and `MAX`.
3. **COUNTIF and related criteria wildcards:** `COUNTIF(range,"A*")` previously returned zero for `Apple` and `Apricot`. Criteria matching now supports Excel-style `*` and `?`, with `~` escaping, for text equality/inequality matching.

Each repair is exercised through the public formula evaluator in the regression suite. Additional edge cases may remain.

## Known behavior gaps relative to Excel

- **Dynamic spilling:** `SEQUENCE`, `FILTER`, `SORT`, `TEXTSPLIT` and similar functions can generate matrices internally; GridForge does not automatically fill adjacent cells or report the full `#SPILL!` experience.
- **Excel date model:** Date serials use a simplified 1900-based convention; the 1900 leap-year anomaly, locale parsing and all timezone cases have not been exhaustively compared.
- **Full-function fidelity:** 506 registered names do not imply 506 identical implementations. `FORECAST.ETS` is a simplified trend estimate; `TRANSLATE`, `IMAGE`, `LAMBDA`, `LET`, `GROUPBY`, higher-order functions and specialized financial functions have restricted semantics.
- **Grid size:** The default worksheet has 400 rows × 52 columns (20,800 cells). Excel supports much larger worksheets.
- **Cell presentation:** Not every Excel format, conditional-format rule, hyperlink, validation rule, chart, pivot table, print configuration or workbook-level feature is preserved.
- **Cross-sheet references and workbook links:** Do not claim full `Sheet2!A1`, structured-table references, external links, named ranges or arbitrary Excel object model compatibility.
- **VBA:** VBA does not execute in the browser. On Windows, the optional Excel bridge can explicitly ask installed desktop Excel to run a trusted macro. This does not constitute an in-browser VBA runtime.
- **XLSM:** Supported cell edits are merged into the original macro-enabled package. Structural edits, advanced formula structures, signatures and unsupported package parts are not guaranteed.
- **Localization and volatility:** Locale-specific separators, coercions, dates, 15-digit precision, and time-dependent functions require targeted desktop-Excel comparison.

## Suggested positioning for reviewers

> GridForge is a **local, browser-based spreadsheet application with an Excel-familiar interface**. It implements a custom TypeScript formula engine with 506 registered Excel-style functions, supports selected XLSX/XLSM workflows, and provides optional desktop Excel integration for VBA. Compatibility with Microsoft Excel is being validated with a reproducible cross-engine test suite; it is not a drop-in replacement.

Do not call it "Microsoft Excel for the browser", claim that all 506 functions are Excel-identical, or publish an Excel compatibility percentage based only on static expected values.

## How to verify real Excel parity

See [EXCEL_PARITY_WINDOWS_GUIDE.md](EXCEL_PARITY_WINDOWS_GUIDE.md). Run a matched test set on your Windows machine, inspect every discrepancy, fix and re-run before publishing a numerical compatibility claim. Add golden workbooks and real user scenarios in later audit rounds.

## Documentation sources

- [Microsoft — ROUND function](https://support.microsoft.com/en-us/excel/functions/round-function)
- [Microsoft — Dynamic arrays and spilled behavior](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior)
- [Microsoft — Excel CalculateFullRebuild](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild)
