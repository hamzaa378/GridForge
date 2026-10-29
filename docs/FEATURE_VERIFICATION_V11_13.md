# GridForge v11.13 Feature Matrix

| Area | Behavior | Verification |
|---|---|---|
| Pareto chart | Orders categories by summed contribution, charts cumulative percentage, combines long tails as Other, rejects negative values | Source test and SVG string checks |
| Box-and-whisker chart | Uses interpolated quartiles, 1.5-IQR whiskers and outlier markers | Source test and SVG string checks |
| Goal Seek | Secant-based target search, bounded iterations and candidate values, single-clone workbook evaluation, stale-preview protection before applying | Source test; source inspection |
| Existing chart types | Histogram supports automatic bins or a user-selected count from 1 to 50; waterfall and prior standard chart types remain available | TypeScript check and build |
| XLSX export | New specialized chart definitions fall back to a clustered column chart in exported XLSX | Source implementation; no Excel GUI test |
| Validation | TypeScript check/build and all 965 automated code tests passed | Automated source checks only |

## Limits

- Goal Seek changes one numeric cell on the active worksheet. It is not Solver, supports no constraints or multiple changing cells, and can fail on discontinuous or difficult formulas.
- Pareto uses summed numeric values per category, supports non-negative inputs, and folds categories after the top 14 into Other.
- Box-and-whisker charting accepts numeric samples from the selected series and uses linear interpolation for quartiles plus Tukey 1.5-IQR outlier classification.
- Specialized histogram, waterfall, Pareto and box-and-whisker charts are not serialized as equivalent native DrawingML chart types; XLSX export uses a clustered-column fallback.
- Pivot summaries remain ordinary cells rather than native Excel PivotTables. Slicers, timelines, Power Query, Data Model, Power Pivot, DAX and OLAP remain unsupported.
- The 506 registered local function names do not establish exact Excel behavior. Ten service-dependent names remain unavailable offline.
- No live UI, Windows desktop, installer or Microsoft Excel interoperability tests were run, per user instruction.
