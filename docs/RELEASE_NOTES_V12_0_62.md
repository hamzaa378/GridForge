# GridForge v12.0.62

- Validate the optional `FORECAST.ETS` family parameters in their Excel positions: seasonality, data completion, and aggregation.
- `FORECAST.ETS.CONFINT` now adjusts its linear-model confidence radius for the selected confidence level, forecast horizon, and extrapolation leverage.
- Dynamic range dependencies used by `INDIRECT`, `OFFSET`, and Pivot lookups are indexed by worksheet column, reducing broad invalidation scans in workbooks with many such formulas.
- The forecasting family remains an offline linear approximation. Parameter validation does not implement Excel's missing-point completion, aggregation, automatic seasonality detection, or AAA smoothing model.
- Static TypeScript check and production build passed. Windows installer and portable packages are built with `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks are run.
