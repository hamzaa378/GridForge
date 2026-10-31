# GridForge 11.27.0 Release Notes

## Formula compatibility

- Added `BIN2DEC`, `BIN2HEX`, `BIN2OCT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `HEX2BIN`, `HEX2DEC`, `HEX2OCT`, `OCT2BIN`, `OCT2DEC`, and `OCT2HEX`, with signed two's-complement interpretation and Excel's fixed-width numeric limits.
- `FLOOR` now returns `#NUM!` when a nonzero number and significance have different signs, consistent with Excel's compatibility function behavior.
- `FLOOR` and `CEILING` stabilize quotients that are within floating-point rounding tolerance of an integer. This avoids an incorrect one-step adjustment for decimal multiples such as 1.1 by 0.1.
- `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, and paired statistical/regression inputs now validate rows and columns, so equally sized but differently oriented ranges no longer pass as compatible.

## Chart trendlines

- Polynomial trendlines now use centered and scaled X values with QR factorization rather than solving raw-power normal equations. This reduces numerical instability for large X coordinates and higher polynomial orders.
- Polynomial equation labels explicitly define their normalized variable (`z`) so the displayed equation corresponds to the fitted model.
- Moving-average trendlines preserve each observed X coordinate for irregularly spaced scatter data. Forward/backward extensions continue from the average spacing between valid observations.
- Chart series now support Standard Error bars, calculated from each series' sample spread and sample size, exposed in the series editor and saved with native DrawingML error-bar metadata.

## Compatibility boundary

These are targeted improvements, not full Excel parity. The trendline editor still supports a bounded subset of Excel trendline styling and options. Formula locale rules, precision and behavior across the full function set still require broader comparison.

## Verification scope

- TypeScript type-check and production build completed.
- Windows installer and source/repository archives were compiled with the Windows test phases skipped.
- No automated tests, Windows runtime tests, installer launch, live UI checks, or Excel GUI checks were run.
