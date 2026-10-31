# GridForge 12.0.19

Date validation entry, validation checks, Pivot date serial handling, and selected financial/date formulas now share one Excel 1900 date-system converter. ISO dates before March 1900 no longer receive the wrong serial offset. Serial 60 maps to Excel's 1900-02-29; the validation editor uses ISO text input so it can preserve that value. Saved bounds are restored after switching field modes, invalid dates are rejected with a clear format hint, and imported 1904-system date-validation bounds are converted to GridForge's 1900 system.

The 1904 workbook date system, locale-specific date entry, and broader Excel date-function behavior remain bounded as described in the compatibility audit.

Verification: TypeScript static check, Python syntax compilation, production build, and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.
