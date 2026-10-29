# GridForge 11.14.0 Release Notes

## Funnel and treemap charts

- Added Funnel charts that show successive stage values as centered trapezoids. Stage order follows the source range; missing values are skipped and negative values show an explanatory chart message.
- Added Treemap charts. Select one or more hierarchy-label columns followed by a numeric value column. Duplicate full paths are aggregated and cells are laid out by relative value.
- Both chart types are available in the chart gallery, editor, and chart-design ribbon.

## Descriptive Statistics report

- Added an Excel ToolPak-style report with mean, standard error, median, mode, sample standard deviation and variance, kurtosis, skewness, range, minimum, maximum, sum and count.
- Text, blank cells and formula errors are excluded. Tied modes resolve to the smallest repeated value; no mode and insufficient sample metrics return `#N/A`.
- The report writes to a chosen destination outside the source range and confirms before overwriting existing cells.

## Compatibility and verification

- Specialized chart previews are still GridForge SVG charts. XLSX output uses a clustered-column fallback for histogram, waterfall, Pareto, box-and-whisker, Funnel and Treemap.
- `npm run check` and the full `npm test` suite passed: 972 automated code tests.
- Live UI, Windows desktop, installer and Excel GUI tests were not run, as requested.
