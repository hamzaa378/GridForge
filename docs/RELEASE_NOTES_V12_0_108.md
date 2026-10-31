# GridForge v12.0.108

## Date and time criteria

- Conditional criteria comparisons now parse supported locale date, time, and combined date-time text when the compared worksheet value is a numeric Excel serial, including localized digits and supported month names.
- This extends date/time criteria to `COUNTIF`, `SUMIF`, `AVERAGEIF`, and the corresponding `*IFS` functions through the shared criterion matcher.
- Ordinary text criteria and wildcard matching retain their text comparison path.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
