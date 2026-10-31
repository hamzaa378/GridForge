# GridForge v12.0.104

## Financial solver compatibility

- `IRR` starts from the supplied guess (10% when omitted), uses a safeguarded domain check, and returns `#NUM!` after 20 failed iterations instead of searching another root. This preserves the starting guess's role when cash flows have multiple possible rates.
- `RATE` now validates payment timing as 0 or 1 and follows a 20-iteration, guess-driven convergence limit.
- `XIRR` uses the supplied guess (10% by default) and stops after 100 iterations; a separate bisection fallback no longer silently changes the selected root.
- The convergence tolerances follow Microsoft's documented 0.00001% for IRR/RATE and 0.000001% for XIRR. This update has not been directly compared to Excel.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
