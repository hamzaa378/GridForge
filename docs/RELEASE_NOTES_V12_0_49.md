# GridForge v12.0.49

- Corrected the generated formula reference to reflect 506 locally evaluated and 10 connected function names.
- Updated function descriptions for supported scoped LAMBDA evaluation and bounded dynamic-array spilling.
- Added a workbook-level 20,000-cell limit before a computed array is placed into the worksheet grid, returning a clear `#NUM!` diagnostic for larger results.
- Verification for this package is limited to static TypeScript checking, production build, and installer compilation. No automated tests or live UI checks were run.
