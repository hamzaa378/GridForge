# GridForge v12.0.125 — Excel variance aliases

The existing `STDEV.S`, `VAR.S`, `STDEV.P`, and `VAR.P` formulas now use the per-column populated-row index for large local reference ranges instead of materializing dense range matrices. They preserve the engine's sample/population calculation and argument behavior.

The modern dotted names were already registered in the formula engine; this release adds their sparse large-range evaluation. Formula import/export continues to preserve the entered function names.

Verification is limited to the TypeScript check, production build, and release packaging. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.
