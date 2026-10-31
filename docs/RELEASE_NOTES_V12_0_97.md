# GridForge v12.0.97

## Time parsing

- `TIMEVALUE` now accepts a supported date followed by a time and returns the fractional-day time portion, matching Excel's documented behavior of ignoring valid date information.
- Date and time parsing share the same locale digit normalization as the related `VALUE` path.
- Invalid date prefixes remain errors; full Excel calendar and locale grammar is not claimed.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
