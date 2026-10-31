# GridForge v12.0.98

## `TIME` field handling

- `TIME` rejects negative hour, minute, or second values and rejects any field greater than 32,767.
- Accepted values above 23 hours or 59 minutes/seconds continue to roll over into the next day/hour/minute.
- Decimal inputs are truncated to whole fields before conversion.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
