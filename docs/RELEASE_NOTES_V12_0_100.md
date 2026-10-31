# GridForge v12.0.100

## Chart data-table accessibility

- Chart data-table SVG now exposes table, row, column-header, row-header, and data-cell roles for assistive technology.
- Cell labels retain their complete text even when the compact visual label is shortened to fit the chart.
- Visual row/category limits are unchanged, and accessibility role support can vary between screen readers and SVG engines.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
