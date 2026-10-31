# GridForge v12.0.95

## Formula compatibility

- `VALUE` now converts supported locale-formatted numeric, date, and time text to numeric values, including date-time strings that combine an accepted date and time.
- `TIMEVALUE` and the time portion of `VALUE` use the same parser, including locale digit normalization and 24-hour or AM/PM times.
- Invalid and empty text still return `#VALUE!`; this update does not claim full Excel parity for every locale, calendar, or ambiguous date string.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
