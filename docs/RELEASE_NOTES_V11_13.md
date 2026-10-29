# GridForge 11.13.0 Release Notes

## Pareto charts

- Added a Pareto chart that sums values for each category, ranks contributions from largest to smallest and draws a cumulative percentage line.
- The chart supports non-negative data and folds categories beyond the top 14 into an Other bar.

## Histogram bins

- Choose automatic bin sizing or set a whole-number bin count from 1 to 50 in the chart editor.

## Box-and-whisker charts

- Added distribution charts with quartiles, median, Tukey 1.5-IQR whiskers and outlier markers.
- Quartiles use linear interpolation and handle short and constant samples without special workbook metadata.

## Goal Seek

- Added a target-value search for one numeric input and one formula result on the active worksheet.
- The secant solver evaluates one cloned workbook document and reuses its dependency graph, has bounded iteration and value limits, and never changes source data during preview.
- Applying a preview is disabled if the workbook changed after calculation; the user must calculate a fresh preview.

## Compatibility and verification

- Pareto and box-and-whisker render in GridForge's SVG chart engine and are exposed through the chart gallery, editor and chart-design ribbon.
- XLSX export currently converts histogram, waterfall, Pareto and box-and-whisker chart types into clustered-column charts. The custom chart type is not preserved as an equivalent editable Excel chart.
- `npm run check`, `npm run build`, and 965 automated code tests passed.
- Live UI, Windows desktop, installer and Excel GUI tests were not run, per user instruction.
