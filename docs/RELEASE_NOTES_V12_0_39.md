# GridForge 12.0.39

Formula-based PIVOTBY row and column sort orders now accept one-dimensional row or column vectors of signed field positions, enabling independent ascending/descending priorities across field levels. Vector sort orders are limited to available field positions; the existing scalar form continues to support sorting by a value measure.

Verification was limited to static TypeScript checking and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
