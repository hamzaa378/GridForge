# GridForge 12.0.25

The formula engine now evaluates LAMBDA expressions in isolated local scopes and supports direct calls such as `=LAMBDA(value,value*2)(A1)`. `MAP`, `BYROW`, `BYCOL`, `MAKEARRAY`, `REDUCE`, and `SCAN` accept LAMBDA callbacks, and `ISOMITTED` distinguishes omitted LAMBDA parameters from supplied blank cells. Empty function arguments are now parsed as blanks so formulas can express Excel's omitted optional arguments.

Safety bounds: LAMBDA nesting is limited to 64 scopes, callback-generated arrays to 20,000 elements, and each LAMBDA to 253 parameters. Workbook-defined named LAMBDAs, recursive named functions, and Excel's complete higher-order function set are not included.

Verification: TypeScript static check and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.