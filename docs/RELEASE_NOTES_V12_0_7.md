# GridForge 12.0.7

- Horizontal bar charts now draw error bars horizontally along the value axis, with direction, end caps, and supported statistical calculations.
- XLSX export writes horizontal (`x`) error bars for horizontal bar chart types; vertical chart types retain vertical (`y`) error bars.
- `npm run check`, production build, installer compilation, and archive integrity checks passed. Tests, application launch, live UI, Windows runtime, and Excel GUI checks were not run.
- Scatter chart X-error bars and custom per-point error ranges remain unsupported.
