# GridForge v12.0.101

## Locale-aware time text

- Shared time parsing now recognizes the current locale's time separator and localized day-period labels, alongside colon and English AM/PM input.
- The derived locale rules are cached and reused by `VALUE`, `TIMEVALUE`, `HOUR`, `MINUTE`, and `SECOND`.
- Combined date-time strings use the same accepted clock grammar. Calendar parsing and locale-specific time syntax still have bounded coverage.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
