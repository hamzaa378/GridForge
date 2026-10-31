# GridForge v12.0.61

- The `FORECAST.ETS` family now uses Excel argument positions, validates matching values and timeline lengths, duplicate and regular timeline steps, forecast target bounds, seasonality, and statistics selectors.
- `FORECAST.ETS.CONFINT` now scales the interval by the requested confidence level instead of always using a fixed 95% multiplier.
- Forecast predictions remain a bounded linear-trend approximation; smoothing coefficients and Excel's full ETS algorithm are not implemented. `FORECAST.ETS.SEASONALITY` uses a local autocorrelation estimate rather than Excel's detector.
- Static TypeScript check and production build passed. Windows installer and portable packages compiled with `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks were run.
