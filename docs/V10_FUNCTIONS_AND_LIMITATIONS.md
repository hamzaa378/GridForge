# GridForge v10 — 100 new Excel-style function implementations

The registry grew from **406 to 506 unique names**. Code: `src/engine/functions_v10.ts`, integrated in `src/engine/functions.ts`. Each name is independently registered with callable behavior, with a direct fixture test in `tests/functions_v10.test.mjs`.

## Added function families

| Category | Names added | Main examples |
| --- | ---: | --- |
| Statistical | 32 | BETA.DIST, CHISQ.DIST, F.DIST, GAMMA.DIST, T.DIST, CONFIDENCE.T |
| Lookup & Reference | 28 | FILTER, SORT, UNIQUE, SEQUENCE, TEXTSPLIT's companion array tools, BYROW, MAP, LET |
| Financial | 16 | COUPNUM, COUPNCD, ACCRINT, AMORLINC, EUROCONVERT |
| Text | 7 | TEXTSPLIT, ASC, JIS, BAHTTEXT, FILTERXML, TRANSLATE |
| Information | 7 | INFO, SHEET, SHEETS, FIELDVALUE, IMAGE, DETECTLANGUAGE |
| Math & Engineering | 5 | BESSELI, BESSELJ, BESSELK, BESSELY, PERCENTOF |
| Database | 3 | GROUPBY, PIVOTBY, CUBESETCOUNT |
| Logical | 1 | ISBETWEEN |
| Other | 1 | AGGREGATE |
| **Total** | **100** | |

## Useful examples

```excel
=BETA.DIST(0.5,2,2,TRUE)
=CHISQ.DIST(2,2,TRUE)
=INDEX(SEQUENCE(3,2),2,2)
=INDEX(TEXTSPLIT("red,blue",","),1,2)
=INDEX(FILTER(A2:B4,C2:C4),1,2)
=INDEX(SORT(A2:B4,1,-1),1,1)
=BYROW(A2:B4,"SUM")
=REDUCE(0,A2:A5,"SUM")
=ISBETWEEN(5,1,10)
=EUROCONVERT(100,"EUR","DEM")
```

`BYROW` returns an array; use `INDEX(BYROW(A2:B4,"SUM"),1,1)` for a scalar in a cell.

## Compatibility tiers

**Implemented with bounded numeric behavior:** distribution calculations, covariance, permutations from existing functions, matrix transformation, filtering, sorting, date serial coupon helpers and currency conversion. Test results cover a representative input for each name, but numerical edge cases and date conventions are not fully Excel-certified.

**Implemented as intentional subsets:**

- `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `TEXTSPLIT`, `VSTACK` and other matrix-returning functions produce in-memory arrays. **No spilled ranges** or `#SPILL!` handling. Wrap with `INDEX` to read a scalar result.
- `LET` and `LAMBDA` do not support Excel lexical bindings, closures or invocation syntax. They accept limited scalar placeholder forms. `MAP`, `BYROW`, `BYCOL`, `SCAN`, `REDUCE`, and `MAKEARRAY` use explicitly selected, safe operation labels (e.g. `"SUM"`), not arbitrary executable lambdas.
- `FORECAST.ETS*` uses a simplified linear-trend forecast, not Excel's exponential-triple-smoothing engine.
- `ACCRINT`, `COUP*`, `AMOR*` and irregular bond yield/price formulas use approximations; do **not** use them for real financial pricing without comparing to Excel and verifying calendars/basis.
- `SHEET` and `SHEETS` report limited current-sheet context, not the full workbook; Excel cross-sheet reference expressions remain unsupported.
- `IMAGE` displays a safe URL/alt-text placeholder rather than an image inside the spreadsheet grid. `TRANSLATE`, `DETECTLANGUAGE`, `CUBESETCOUNT`, `FIELDVALUE`, `FILTERXML`, `PHONETIC`, and `BAHTTEXT` provide restricted, offline-only behavior.
- `AGGREGATE` currently implements selected codes and options, not every Excel calculation and hidden-row behavior.

**Not implemented:** live Cube / OLAP engine, stock-history provider, general WEBSERVICE requests, fully compatible VBA in-browser, dynamic spilling, multi-sheet reference evaluation.

## Testing and traceability

- `tests/functions_v10.test.mjs` invokes **all 100 functions** with example arguments, checks 20 independent known results, verifies formula integration, and checks structured failure paths.
- `npm test` runs the full inherited formula/workbook tests as well as the v10 set.
- `scripts/regression_browser.py` verifies the new formulas through a real Chromium grid.
- `examples/v10_100_function_reference.xlsx` and `.json` document all 100 names with 30 evaluated live formulas.
- `docs/FUNCTION_REFERENCE.md` is generated from the actual registry, not a manually maintained list.
- `docs/V10_EXCEL_FUNCTION_GAP.md` uses a frozen, explicitly bounded list. It is **not** an up-to-date total of all Microsoft Excel functions.

For the authoritative function names and evolving version constraints consult [Microsoft's Excel function list](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical).
