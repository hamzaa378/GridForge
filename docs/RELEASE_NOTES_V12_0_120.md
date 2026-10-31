# GridForge v12.0.120 — Larger sparse aggregation ranges

This release extends selected worksheet-local aggregation formulas to ranges larger than 20,000 cells, up to the configured worksheet dimensions. It covers `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS`, and `MINIFS` when their range arguments are local references.

The dependency graph stores these references as compact rectangular ranges and uses its column index to find affected formulas after edits. The evaluator's sparse paths visit populated cells and account for blank cells without allocating a dense array for the entire range. General formula ranges retain the 20,000-cell cap. `COUNTIF`, `SUMIF`, and `AVERAGEIF` retain that cap where array criteria or range alignment can require dense behavior.

This improves performance and capacity for sparse sheets. It does not establish full Excel parity: workbook dimensions, formula-specific semantics, and Excel's recalculation behavior still differ. Verification for this release is limited to static TypeScript checking and packaging; no tests, desktop launch, or live UI checks are performed.
