# GridForge v12.0.96

## `VALUE` currency text

- `VALUE` recognizes one Unicode currency symbol immediately before or after a locale-formatted number, including the documented `VALUE("$1,000")` form.
- Numeric parsing still uses the current locale's decimal/group marks and localized digits.
- `NUMBERVALUE` remains unchanged. Accounting parentheses, currency codes, and every regional symbol placement are outside this implementation's current scope.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
