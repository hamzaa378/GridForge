# GridForge v12.0.105

## NPV reference handling

- `NPV` discounts only numeric cash-flow entries. Text, logical values, empty cells, and errors in references are ignored, and ignored entries do not consume a discount period.
- Out-of-bounds references and missing worksheet references still produce `#REF!`.
- The main-thread evaluator now handles NPV references directly so a referenced formula error can be skipped before normal range evaluation would propagate it.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
