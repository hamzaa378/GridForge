# GridForge 12.0.9

- Scatter chart data labels now honor the configured center, inside-end, outside-end, and best-fit placement instead of always appearing at the same offset. Best-fit labels shift inward near the plot boundaries.
- Carries forward v12.0.8 scatter X/Y error bars, automatic axis bounds, and XLSX error-direction round-trip.
- `npm run check`, production build, installer compilation, and archive integrity checks passed. Tests, application launch, live UI, Windows runtime, and Excel GUI checks were not run.
- A series can store one error-bar direction at a time; custom per-point ranges and simultaneous X and Y error bars on the same series remain unsupported.
