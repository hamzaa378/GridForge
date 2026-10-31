# GridForge v12.0.64

- Seasonal `FORECAST.ETS` now fits an additive triple-exponential smoothing model with bounded grid search over alpha, beta, and gamma. Automatic seasonality uses the local autocorrelation estimate; short, nonseasonal, and oversized series use the linear fallback.
- `FORECAST.ETS.STAT` returns fitted smoothing coefficients when the seasonal model is active, and calculates error metrics from that model's in-sample residuals.
- `FORECAST.ETS.CONFINT` uses the seasonal model residual scale when available and retains a confidence/horizon approximation. Excel's exact AAA optimizer, duplicate aggregation, and full confidence calculation remain unsupported.
- Blank values in the forecast series now preserve their timeline positions. `data_completion=0` fills missing values with zero; the default `data_completion=1` interpolates between observations and carries the nearest value across an endpoint gap. More than 30% missing values are rejected. Duplicate-timeline aggregation remains unsupported.
- The parameter search is bounded to at most 4,000 observations to limit synchronous recalculation work.
- Static TypeScript check and production build passed. No tests, app launch, live UI, Windows runtime, or Excel GUI checks are run.
