# GridForge 506-function local audit — current source refresh

**506/506 offline function names invoked through the current GridForge Workbook** (the registry also has 10 service-dependent names excluded from this offline set).

- **500** returned a non-error value.
- **6** returned formula errors: `AVERAGEA` (`#VALUE!` for the deliberately invalid text argument), `NA` (`#N/A` by design), and four odd-coupon financial functions (`#NUM!` for the current invalid fixture arguments).
- **0** uncaught engine exceptions.
- **979/979** Node regression tests passed on the current source refresh.

## Function scopes

- scalar: 445
- limited-semantics: 29
- array-index: 27
- volatile: 5

## Interpretation

This is a local GridForge-side verification, not proof of comprehensive Excel behavior. The saved Microsoft Excel 16.0 (build 20430.0) run has the same fixture hash and was compared against freshly regenerated current-source outputs: 450 of 456 comparable one-case probes match (98.68%); 2 are differences and 4 are already classified as known semantic differences. Excel rejected or lacked 36 tested names, 6 Excel COM cases failed during formula assignment, 3 service-dependent cases were skipped, and 5 volatile results were excluded from equality. See `docs/PARITY_506_COMPARISON.md` for each outcome. One formula invocation per name does not cover edge cases, locales, error coercion, date systems, precision or full array-spill behavior.

See `docs/PARITY_506_WINDOWS_GUIDE.md` for the exact reproduction steps.
