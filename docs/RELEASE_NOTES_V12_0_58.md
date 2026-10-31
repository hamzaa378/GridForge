# GridForge v12.0.58

- XLSX legacy and dynamic array-formula anchors no longer recalculate as isolated formulas against their cached result cells. GridForge preserves the imported array outputs as cell values and warns that the formula itself cannot be recalculated.
- Static TypeScript check and production build passed. Windows installer and portable packages compiled with `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks were run.
