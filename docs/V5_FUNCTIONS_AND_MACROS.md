# GridForge v5 — 80 New Functions and Macro-Safe Compatibility

## Overview

- Original custom registry: **56 functions**.
- v4 expansion: **80** additional functions.
- v5 expansion: **80** additional functions.
- **Total: 216 registered Excel-style functions** in `src/engine/functions.ts`.
- Browser function sidebar now provides **Excel-like category navigation**, search, and a Recently added filter.

| Category added in v5 | Count | Implemented functions |
| --- | ---: | --- |
| Math & Engineering | 18 | ACOSH, ASINH, ATANH, ACOT, ACOTH, COT, COTH, CSC, CSCH, SEC, SECH, MROUND, FACTDOUBLE, MULTINOMIAL, ARABIC, ROMAN, BASE, DECIMAL |
| Statistical | 12 | SKEW, KURT, COVARIANCEP, COVARIANCES, PEARSON, STEYX, STANDARDIZE, TRIMMEAN, NORMDIST, NORMSDIST, NORMINV, NORMSINV |
| Text | 14 | T, N, DOLLAR, FIXED, TEXT, LENB, LEFTB, RIGHTB, MIDB, FINDB, SEARCHB, REPLACEB, REGEXTEST, REGEXREPLACE |
| Date & Time | 8 | DATEDIF, YEARFRAC, DAYS360, WEEKNUM, ISOWEEKNUM, DATEVALUE, TIMEVALUE, NETWORKDAYSINTL |
| Financial | 15 | RATE, NPER, IPMT, PPMT, EFFECT, NOMINAL, SLN, SYD, DDB, MIRR, IRR, XNPV, XIRR, FVSCHEDULE, RRI |
| Information | 8 | ISERR, ISERROR, ISNA, ISEVEN, ISODD, ISLOGICAL, ISNONTEXT, TYPE |
| Lookup & Reference | 5 | LOOKUP, CHOOSE, ROWS, COLUMNS, ADDRESS |
| **Total** | **80** | Each new function has an individual automated calculation test |

Excel stores certain modern names with `_xlfn.` prefixes, and some names contain dots. Supported import/export aliases include `NETWORKDAYS.INTL`, `NORM.DIST`, `NORM.S.DIST`, `NORM.INV`, `NORM.S.INV`, `COVARIANCE.P`, and `COVARIANCE.S`. The GridForge formula editor uses the unpunctuated names (e.g., `NETWORKDAYSINTL`).

## Working examples

```excel
=MROUND(14,5)
=ROMAN(1987)
=NORMDIST(0,0,1,TRUE)
=ISERROR(1/0)
=CHOOSE(2,1/0,"Used branch only")
=NETWORKDAYSINTL(DATE(2026,10,19),DATE(2026,10,23),1)
=IPMT(0.05/12,1,60,10000)
=REGEXREPLACE("INV-1234","[0-9]+","9999")
```

`CHOOSE`, `ISERROR`, `ISERR`, `ISNA` are special-cased in `workbook.ts` to evaluate only the necessary expressions. An unselected branch containing an error does not cause a failure.

## Compatibility boundaries

- All functions are implemented **without an external spreadsheet calculation engine**, but they are not guaranteed to have identical Excel optional arguments, locale handling, or numerical precision.
- Date functions use Excel's **1900 serial date model** with UTC arithmetic. Some 1900 leap-day and locale/timezone cases differ. `DATEVALUE` currently accepts ISO `YYYY-MM-DD`; `TIMEVALUE` accepts 24-hour `HH:MM[:SS]`.
- `NETWORKDAYSINTL` supports standard weekend codes 1–7 and 11–17 or a 7-digit weekend mask. GridForge names omit Excel's dot.
- `REGEXTEST` and `REGEXREPLACE` use a **restricted JavaScript regular-expression subset**, not the whole Microsoft Excel regex flavor. Backreferences, lookaround and nested quantified groups are blocked as potentially expensive.
- `TEXT` supports common general/decimal/currency/percent patterns, not arbitrary Excel format strings.
- `LENB`/`LEFTB`/`MIDB` etc. measure UTF-8 bytes; Excel byte functions historically use a different locale-sensitive encoding, so international text results can differ.
- `ROMAN` supports traditional Roman formatting (not all Excel forms); `LOOKUP` expects sorted inputs for approximate matching.
- `IRR`, `RATE`, `XIRR` are numerical solvers with bounded search/convergence and may fail on pathological cash flows or cases with multiple possible roots. `YEARFRAC`/`DATEDIF`/`DAYS360` implement useful subsets of Excel conventions. Validate financial outputs against Excel or another independent source before relying on them.
- No array spills, native dynamic-array functions, external links, cross-sheet formula evaluation, pivot-table engine, VBA runtime, or arbitrary Excel add-ins.

## What “macro support” means in v5

### 1. Safe GridForge action macros (actually runnable)

GridForge records **cell content changes** and **cell formatting** initiated through its normal editor and formatting controls. It stores a JSON recipe on the current browser origin and lets users replay it on the originally named sheet.

Workflow:

1. Select **Automate**, enter a name and click **Start recording**.
2. Edit a cell or use formatting toolbar controls. Supported changes are recorded; sorting, arbitrary right-click operations, workbook imports and formula results are not recorded.
3. Reopen **Automate** and click **Save recording**.
4. Choose the name from the list and click **Run selected macro**.
5. To share it, choose **Export macro JSON**. Another user can select **Import macro JSON** after inspecting its contents.

Security/limits: only `cells` and `styles` operations, maximum 200 steps and 5,000 target operations per macro. Incoming addresses, strings, styles, version and sheet names are validated before replay. A replay is an undoable grouped transaction. There is **no script interpreter** and no VBA/JavaScript/shell execution. Recipes are stored in localStorage, not written into XLSX.

### 2. Opening a legacy `.xlsm` (VBA is NEVER executed)

1. Choose **Import** and select an `.xlsm` workbook.
2. Supported worksheets, cell data, formulas and basic formatting are loaded by the existing OOXML reader. Read the import report; unsupported formulas or spreadsheet objects can be missing.
3. An **Original XLSM** control becomes available. It downloads the exact file originally uploaded, byte for byte, even if cells were subsequently edited in GridForge.
4. To save the **edited** content, choose **XLSX** or **JSON**. XLSX is a fresh GridForge OOXML export and does not include any VBA project.

**Crucial warning:** The original `.xlsm` download contains **none of the edits made in GridForge**. Conversely the edited `.xlsx` contains **none of the original VBA**. GridForge **does not merge workbook modifications with VBA macros, preserve Excel form controls, or execute imported macro procedures.** To combine edited content and executable VBA, use Excel and test it there. Excel may display its own macro-security warnings depending on document trust settings.

`.xls` (legacy binary) is not an OOXML ZIP and cannot be imported by this loader. Some protected/encrypted `.xlsm` files, external workbook formulas, charts, names and advanced Excel features may not load. Import limit: 25 MB; sheet/cell count limits still apply.

## Code responsibilities

- `src/engine/functions_v5.ts` — 80 functions, argument checks and category metadata.
- `src/engine/categories.ts` — Excel-like category mapping for all 216 functions.
- `src/engine/functions.ts` — central registry, combining 56 + 80 + 80.
- `src/engine/workbook.ts` — expression evaluation and lazy branches.
- `src/engine/macros.ts` — action-macro JSON validation and replay.
- `src/main.ts` — function category UI, recording controls, XLSM original download.
- `src/engine/xlsx.ts` — supported OOXML formula name normalization, macro-safe import warnings.
- `tests/functions_v5.test.mjs` — all 80 new calculations, errors, categories and macros.
- `scripts/regression_browser.py` — realistic browser macro recorder, original XLSM and edited XLSX tests.

## Testing and demonstration

```powershell
npm run regression:core
npm run regression
npm run gallery:build
```

In GridForge, **Import → `examples/v5_80_function_gallery.xlsx`** to see all 80 new functions producing live values. The matching JSON is the most faithful GridForge backup. To compare Excel compatibility, try a copy of a real spreadsheet but retain the original and independently review outputs.
