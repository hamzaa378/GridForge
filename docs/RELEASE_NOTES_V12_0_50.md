# GridForge v12.0.50

- Updated general numeric rendering and text export to use 15 significant digits, avoiding the previous fixed ten-decimal rounding that could erase very small values or change larger results.
- Carries forward the corrected 506 local / 10 connected function reference and the bounded dynamic-array spill limit from v12.0.49.
- Static TypeScript check, production build, and installer compilation only; no automated tests or live UI checks were run.
