# GridForge v9 — Code audit and compatibility boundaries

## Audit performed

- Registry uniqueness: 306 existing + 100 **new unique** names; the first candidate set contained a duplicate (`QUARTILE`) and was corrected by implementing `BINOM.INV` instead.
- Parser: modified boolean-token precedence so both `=TRUE` and `=TRUE()` work, while retaining ordinary literals.
- Engineering numeric bounds: invalid signed radix conversions such as `DEC2BIN(512)` return an error instead of silently producing an out-of-range result.
- SUBTOTAL: rejects unsupported operation codes, rather than treating arbitrary codes as valid modulo 100.
- Function-category classification: 406 names assigned to exactly one visible category, including Database and Compatibility.
- UI: all old toolbar controls remain reachable in All tools; new command grouping and optional learning guide are regression-tested.

## Tested but not certified as Excel-identical

- Statistical numerical integration/inversion uses iterative approximations; extremes and distribution tails can differ from Excel.
- Legacy Treasury/fixed-income functions use simplified accrual/coupon conventions. Do not use as authoritative securities valuation.
- Matrix/frequency/regression functions return internal arrays; GridForge requires `INDEX` instead of spilling to surrounding cells.
- Excel conditional database criteria, external data, date systems, and number display settings are bounded subsets.
- Legacy XLSM edit/export preserves VBA and unsupported package parts but cannot update all charts/formatting/shapes or worksheet structure.
- Desktop VBA execution requires a trusted workbook and Microsoft Excel on Windows; Linux tests validate the protocol and file preservation, not the Excel VBA runtime.

## Regression commands

- `npm run check`, `npm test` for compiler and engine
- `npm run benchmark` and `npm run benchmark:tools` for recalculation and large edits
- `npm run regression` for full 7-stage suite, including Chromium, XLSM and VBA bridge validation
- `npm run functions:gaps` regenerates the frozen-catalog gap report.

## Scope of the missing-function count

The reference in `excel_comparison_catalog.json` intentionally combines a fixed historical compatibility list with a hand-curated newer set. This allows an exact reproducible tracked count but is not an exhaustive list of Excel for Microsoft 365 functions.
