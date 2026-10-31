# GridForge v12.0.123 — Sparse order statistics

`MEDIAN`, `LARGE`, and `SMALL` now process large local reference ranges using the populated-row index instead of materializing every blank cell in a dense matrix. The formulas ignore nonnumeric values from referenced cells, include dynamic-array spill values, and propagate errors from referenced formulas.

`LARGE` and `SMALL` retain the existing rank argument handling. Mixed literal/array and cross-sheet range inputs continue to use the standard evaluator and its general range cap.

Verification is limited to the TypeScript check and release build/package. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.
