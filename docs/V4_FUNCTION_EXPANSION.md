# GridForge v4 — 80 Additional Excel-Style Functions

**What changed:** GridForge expands from **56 to 136** independently registered functions. All are implemented in TypeScript without an external spreadsheet formula engine. The previous 56 functions and all existing workbook features remain available.

## Function inventory by category

### Mathematics and trigonometry (24)

`SUMSQ`, `SUMPRODUCT`, `QUOTIENT`, `TRUNC`, `EVEN`, `ODD`, `GCD`, `LCM`, `FACT`, `COMBIN`, `PERMUT`, `RADIANS`, `DEGREES`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `SINH`, `COSH`, `TANH`, `LOG10`

### Statistics and regression (16)

`GEOMEAN`, `HARMEAN`, `STDEVP`, `VARP`, `DEVSQ`, `AVEDEV`, `MODE`, `PERCENTILE`, `QUARTILE`, `RANK`, `CORREL`, `RSQ`, `SLOPE`, `INTERCEPT`, `FORECAST`, `PERCENTRANK`

### Text processing (12)

`PROPER`, `REPT`, `CLEAN`, `TEXTJOIN`, `CONCATENATE`, `CODE`, `CHAR`, `UNICODE`, `UNICHAR`, `NUMBERVALUE`, `TEXTBEFORE`, `TEXTAFTER`

### Dates and times (16)

`DATE`, `YEAR`, `MONTH`, `DAY`, `TODAY`, `NOW`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `DAYS`, `EDATE`, `EOMONTH`, `WEEKDAY`, `NETWORKDAYS`, `WORKDAY`

### Logic, criteria and lookup (8)

`XOR`, `IFS`, `SWITCH`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `XLOOKUP`, `HLOOKUP`

### Finance (4)

`PMT`, `FV`, `PV`, `NPV`

## Try the new functions

| Formula to enter | Expected result |
| --- | --- |
| `=SUMPRODUCT(A1:A3,B1:B3)` | Sum of pairwise products of the equally shaped ranges |
| `=SUMIFS(C2:C10,A2:A10,"North",B2:B10,">10")` | Sum where both conditions match |
| `=XLOOKUP("SKU-2",A2:A20,C2:C20,"Not found")` | Matching return value or fallback |
| `=IFS(A1>90,"A",A1>80,"B",TRUE,"C")` | Grade for the first true condition |
| `=LOG10(1000)` | `3` |
| `=TEXTJOIN(", ",TRUE,"one","two")` | `one, two` |
| `=YEAR(DATE(2026,10,20))` | `2026` |
| `=NETWORKDAYS(DATE(2026,10,19),DATE(2026,10,23))` | `5` |
| `=PMT(0,10,1000)` | `-100` |

## Engineering changes

1. `src/engine/functions_extra.ts`: 80 independently implemented function definitions and argument/domain checks.
2. `src/engine/functions.ts`: registers the new definitions beside the earlier 56.
3. `src/engine/workbook.ts`: adds lazy branches for `IFS`, `SWITCH` and the optional `XLOOKUP` not-found value.
4. `src/engine/parser.ts`: resolves `LOG10(...)` versus `LOG10` address tokenization.
5. `src/engine/xlsx.ts`: writes and normalizes modern Excel `_xlfn.` prefixes without rewriting quoted strings.
6. `tests/functions_v4.test.mjs`: covers each new function, edge cases, dependency behavior and OOXML output.
7. `scripts/regression_browser.py`: new end-to-end cases for formulas inside the grid.

## Compatibility boundaries

- GridForge is **not completely compatible with Microsoft Excel**. The new functions have bounded signatures and supported subsets.
- No dynamic array-spilling, cross-sheet formula references, named ranges, structured references, implicit intersections or lambda expressions.
- `XLOOKUP`: exact match with one-dimensional lookup/return ranges; no approximate matches, reverse search, wildcard mode or spilled return.
- `HLOOKUP`: exact matching only. `PERCENTILE` and `QUARTILE` use inclusive interpolation.
- Date values use Excel’s 1900 serial system. GridForge uses UTC for date arithmetic, and `TODAY` and `NOW` update only when their formula recalculates. The results appear as numeric serials by default, not human-readable dates.
- The 1900 fictitious day 60 is handled for calendar components. Historical-weekday and pre-1900 edge cases differ from Excel.
- Conditional criteria do not implement Excel wildcards or locale-specific matching. Paired regression skips incomplete numeric pairs.
- Financial functions cover periodic, constant-rate cash flows; no advanced day-count conventions or IRR solver.
- XLSX import/export remains a supported subset. Keep the original Excel file and a GridForge JSON backup when preserving exact GridForge state matters.

## Ready-to-import function gallery

The `examples/v4_80_function_gallery.json` workbook includes all 80 additional formulas in a dedicated table, with source inputs in columns `AA:AJ` to avoid accidentally changing formula results. Import the JSON into GridForge. A corresponding `examples/v4_80_function_gallery.xlsx` is provided for interoperability experiments; keep the JSON version as the complete GridForge-compatible backup.

Rebuild the examples from tested formulas using `npm run gallery:build`.

## Test commands

```powershell
npm install
npm run regression:core
```

For real browser tests, install Playwright Chromium as described in `docs/SETUP_FROM_ZERO.md` and run `npm run regression`. Python-only users can start the precompiled app with no npm install; source changes and rebuilding require Node.js.
