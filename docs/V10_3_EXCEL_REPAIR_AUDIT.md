# GridForge v10.3 — Excel calculation repair audit

## Measured comparison

The baseline results were produced by the user's **actual Windows Microsoft Excel 16.0, build 20430.0**, running the frozen one-case-per-function 506-fixture suite. The v10.3 formula outputs were re-evaluated locally against that **saved Excel result**, not against another live Excel session.

| Same frozen 506-case fixture | v10.2 | v10.3 |
|---|---:|---:|
| Matching comparable cases | 418 | **450** |
| Unexpected mismatches | 34 | **2** |
| Known non-Excel-equivalent functions | 4 | **4** |
| Comparable denominator | 456 | **456** |
| Agreement on this comparable set | 91.67% | **98.68%** |
| Originally matching cases that regressed | — | **0** |

**Scope:** exactly one fixture per registered function. This is not complete behavioral equivalence to Excel. The 36 unavailable/rejected, 6 COM failures, 3 deliberately skipped online, and 5 volatile cases are not silently counted as matches. Some unsupported modern Excel functions may be available on other Excel versions.

## The 32 repaired cases

- **Numerical and statistical:** `CHISQ.INV`, `CHISQ.INV.RT`, `CONFIDENCE`, `CONFIDENCE.NORM`, `ERF.PRECISE`, `ERFC.PRECISE`, `T.INV`.
- **Engineering and text:** `BESSELY`, `ASC`, `AVERAGEA`, `IMCOS`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMSIN`.
- **Finance and bonds:** `AMORDEGRC`, `DISC`, `DURATION`, `INTRATE`, `MDURATION`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `RECEIVED`, `YIELD`, `YIELDDISC`, `YIELDMAT`.

### Root causes and changes

1. **Chi-square inverse instability.** The large-argument gamma CDF overflowed and produced non-monotonic values. Replaced with series/continued-fraction regularized gamma and robust inversion.
2. **Precision.** Improved normal integral/error functions, including inverse-normal confidence calculations and the `T.INV(0.5,df)` symmetry case.
3. **Bessel Y.** Added the missing second integral term rather than treating Bessel Y as only a signed sine integral.
4. **Financial date conventions.** Introduced 30/360 day counts, a maturity-anchored coupon schedule, fractional-coupon exponent discounting, clean-price accrued-interest subtraction, and corresponding yield/duration calculations.
5. **Odd-coupon input validation.** Rejects the invalid dates/rates in the saved Excel fixture instead of returning a spurious numeric result. Valid odd-coupon calculations remain approximate and are **not** promised to match Excel.
6. **Text and complex formatting.** Added Excel-like 15-significant-digit presentation, locale-aware handling of the English-Excel `ASC` fixture, and erroring on directly supplied nonnumeric strings in `AVERAGEA`.
7. **Depreciation.** Corrected `AMORDEGRC` rate coefficients and remaining-book-value method. Kept `AMORLINC`'s original straight-line behavior.

### Two cases still open

- **`IMPOWER("1+i","2")`**: GridForge currently returns `1.22464679914735E-16+2i`; Excel produced `1.22514845490862E-16+2i`. The difference in the real component is about `5.0e-20`, but the outputs are strings and so fail strict textual comparison. Do not hardcode the reference string to fake a pass.
- **`XNPV(0,F1:F5,G1:G5)`**: GridForge calculates `400` for five chronological dates and flows summing to 400. The saved Excel run reports `#NUM!`. This may be an Excel input-fixture or automation issue. Do not alter a mathematically plausible calculation to reproduce the error without a focused, real-Excel recheck.

### Four known semantic differences

`CUBESETCOUNT` (no OLAP engine), `FIELDVALUE` (text-only substitute rather than Excel linked data), `FORECAST.ETS.SEASONALITY` (linear placeholder instead of ETS), and `INFO` (GridForge environment, not Excel version). They remain non-equivalent and are reported as such.

## How to repeat the audit

```powershell
npm install
npm run parity:snapshot
npm test
```

The frozen evidence is in `docs/PARITY_506_EXCEL_RESULTS.json` and `docs/EXCEL_V10_2_BASELINE_COMPARISON.json`, and the re-evaluated cells are in `docs/EXCEL_SNAPSHOT_V10_3.json`. The source has `tests/excel_parity_v10_3_snapshot.test.mjs` and `scripts/verify_excel_snapshot.mjs` to fail the build if original matches regress or fewer than 32 mismatches are fixed.

**For an updated real Excel comparison on Windows:**

```powershell
python scripts/excel_focus_v10_3.py
python scripts/excel_parity_506_desktop.py
python scripts/excel_parity_506_compare.py
```

Requires `pywin32` and installed desktop Excel. The focused runner writes separately to `docs/EXCEL_FOCUSED_RECHECK.json` and never overwrites the original 506-result file. The full runner re-evaluates all cases and overwrites the 506-result file, so back it up if you need to retain the original evidence.
