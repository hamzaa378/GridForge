# GridForge v12.0.66

- Draw and Page Layout now use a compact grid-based application layout while active. The toolbar occupies a fixed 96-pixel row, and the worksheet receives the remaining vertical space.
- This replaces the previous flex-only sizing lock that did not reliably contain the oversized ribbon seen in the supplied screenshot.
- v12.0.65 forecast completion and duplicate-timestamp aggregation improvements are included.
- Static TypeScript check and production/installer compilation are used; no tests, app launch, live UI, Windows runtime, or Excel GUI checks are run.
