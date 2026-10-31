# GridForge 12.0.8

- Scatter charts now support horizontal X error bars and vertical Y error bars, with direction, end caps, and the existing fixed, percentage, standard deviation, and standard error models.
- Scatter chart preview axes include the configured error ranges when automatic bounds are used, keeping the bars inside the plot.
- XLSX chart import/export preserves the error-bar axis direction (`x` or `y`) in DrawingML.
- `npm run check`, production build, installer compilation, and archive integrity checks passed. Tests, application launch, live UI, Windows runtime, and Excel GUI checks were not run.
- A series can currently store one error-bar direction at a time; custom per-point ranges and simultaneous X and Y error bars on the same series remain unsupported.
