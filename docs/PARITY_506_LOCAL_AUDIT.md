# GridForge 506-function local audit — current source refresh

**506/506 frozen Excel parity probes** invoked through the current GridForge Workbook. The current registry has 506 local/workbook-context names plus 10 connected names; this historical set includes the now-connected `CUBESETCOUNT` and predates workbook-context `GETPIVOTDATA`.

- **484** returned a non-error value.
- **22** returned formula errors, including expected `NA` and connected `CUBESETCOUNT` errors, intentionally invalid financial inputs, and unsupported or restricted lambda-based array cases.
- **0** uncaught engine exceptions.
- **1,080/1,080** Node regression tests passed in the latest full code-only suite.

## Function scopes

- scalar: 445
- limited-semantics: 29
- array-index: 27
- volatile: 5

## Interpretation

This is a local GridForge-side verification, not proof of comprehensive Excel behavior. The saved Microsoft Excel 16.0 (build 20430.0) run has the same frozen fixture hash and was compared against freshly recalculated current-source outputs: 456 of 466 comparable one-case probes match (97.85%); no unexplained differences remain and 10 are classified semantic differences. The comparator evaluates complex results by numeric components with explicit floating-point tolerances. Excel rejected or lacked 26 tested names, 6 Excel COM cases failed during formula assignment, 3 service-dependent cases were skipped, and 5 volatile results were excluded from equality. See `docs/PARITY_506_COMPARISON.md` for each outcome. One formula invocation per frozen parity case does not cover edge cases, locales, error coercion, date systems, precision or full array-spill behavior.

See `docs/PARITY_506_WINDOWS_GUIDE.md` for the exact reproduction steps.
