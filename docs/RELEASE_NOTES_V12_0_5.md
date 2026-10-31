# GridForge 12.0.5

- Chart error bars now offer both-sided, plus-only, and minus-only directions, with optional end caps.
- Charts can show blank source cells as gaps, zeroes, or connected line points; the setting is available in the editor and supported DrawingML XLSX round-trips.
- The selected error-bar direction and cap setting render in the chart preview and round-trip through supported DrawingML XLSX charts.
- `npm run check` and production packaging passed. Tests, application launch, live UI, Windows runtime, and Excel GUI checks were not run.
- Existing chart limitations remain, including custom error-bar values per data point and advanced chart formatting, blank-cell behavior on specialized ChartEx charts.
