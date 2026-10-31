# GridForge v12.0.124 — Sparse aggregate family

`PRODUCT`, `STDEV`, and `VAR` now use the per-column populated-row index for large local reference ranges. They collect numeric values from occupied cells and spilled arrays, preserve the existing function implementations for aggregation and rank/error behavior, and propagate errors from referenced formulas.

Large ranges are enabled only when every function argument is a local cell or range reference. Mixed literals, arrays, and cross-sheet references retain the standard evaluator and the general range cap.

Verification is limited to the TypeScript check and release build/package. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.
