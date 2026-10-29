# GridForge 11.26.0 Release Notes

## Excel-compatible date calculations

- `DAYS360` now uses the U.S. (NASD) 30/360 convention by default and the European convention when the method argument is TRUE. U.S. calculation adjusts month-end dates and preserves negative results when the start follows the end.
- `YEARFRAC` basis 0 now follows the U.S. 30/360 convention; basis 4 uses European 30/360. Basis 1 accumulates actual serial days against each calendar year's length, including the additional serial day in Excel's 1900 date system.
- Shared date helpers and working-day functions (`NETWORKDAYS`, `WORKDAY`, `NETWORKDAYS.INTL`, and `WORKDAY.INTL`), as well as `DAYS360` and `YEARFRAC`, validate dates against the Excel 1900-system range, ending at serial 2,958,465. Excel serial 60 continues to represent its fictitious 1900-02-29. The international functions now reject out-of-range holidays; `WORKDAY.INTL` also rejects an all-weekend mask.

## Compatibility boundaries

This improves a targeted part of date compatibility; it does not make formulas fully equivalent to Excel. Locale-aware date parsing, the 1904 date system in formulas, all error/coercion rules, floating-point edge cases, and comprehensive comparisons against Excel remain incomplete. Excel documents a known YEARFRAC caveat for U.S. basis with a start date at February month-end; no live Excel comparison was performed for this release. See the [v11.26 feature matrix](FEATURE_VERIFICATION_V11_26.md) and the [Excel compatibility audit](EXCEL_COMPATIBILITY_AUDIT.md).

## Verification scope

- `npm run check` passed (TypeScript no-emit).
- The TypeScript check and build passed; function documentation was generated.
- The Windows installer and portable app compiled with `-SkipTests`; the GitHub-ready source and portable archives were generated. SHA-256 values are in `build/release/SHA256SUMS.txt`.
- No automated tests, Windows tests, installer launch, live UI checks, or Excel GUI import/export checks were run.
