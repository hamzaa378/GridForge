# GridForge v12.0.65

- `FORECAST.ETS` preserves blank observations at their timeline positions and applies `data_completion=0` as zero-fill or the default `data_completion=1` as bounded linear interpolation, including nearest-value endpoint completion.
- Duplicate timeline values are grouped using aggregation codes 0–6: average, sum, count, counta, minimum, maximum, and median. Numeric/text-aware COUNTA is available for duplicate groups.
- Forecast completion rejects series with more than 30% missing observations. Excel's AAA parameter optimizer and exact Excel confidence intervals remain unsupported.
- The seasonal additive triple-exponential model and offline linear fallback remain bounded as described in v12.0.64.
- Static TypeScript check and production build are used; no tests, app launch, live UI, Windows runtime, or Excel GUI checks are run.
