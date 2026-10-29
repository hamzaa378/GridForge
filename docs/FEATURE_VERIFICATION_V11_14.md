# GridForge v11.14 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| Funnel chart | Sequential stage trapezoids scaled to the maximum; handles missing values and reports negative or empty input clearly | SVG source tests |
| Treemap chart | Hierarchy paths from one or more leading label columns; final column is the non-negative weight; duplicate paths aggregate; weighted rectangles use a squarified layout | Engine tests for aggregation, rendering and invalid negatives |
| Descriptive Statistics | Writes a 13-metric report plus headers to a separate 2-column range; ignores non-numeric values; rejects overlap and out-of-bounds reports | Source tests for metrics, source preservation and validation |
| XLSX chart compatibility | Funnel and treemap export through the existing clustered-column fallback | OOXML package tests; no Excel GUI test |
| Full verification | TypeScript check, frontend build, and all 972 automated source tests passed | Code-only; no live UI or Windows checks |

## Limits

- Funnel uses the first numeric series and preserves input stage order; it does not calculate conversion rates automatically.
- Treemap groups exact joined hierarchy labels, uses the last column as weights, and does not implement Excel's full group/level formatting or interactive drilldown.
- Descriptive Statistics produces static numeric values, not formula-linked metrics. It uses sample variance/deviation and Excel-style adjusted skewness/kurtosis formulas; unavailable sample metrics are represented by `#N/A`.
- Native XLSX chart export still falls back to clustered columns for histogram, waterfall, Pareto, box-and-whisker, funnel and treemap charts.
- Pivot Summary remains ordinary cells, and Power Query, Data Model, Power Pivot, DAX, OLAP, Solver and full PivotTable behavior remain absent.
- No live UI, Windows desktop, installer or Excel GUI tests were run.
