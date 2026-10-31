# GridForge v12.0.99

## Time-component functions

- `HOUR`, `MINUTE`, and `SECOND` now accept supported time text as well as numeric serials.
- They share locale-digit normalization and time validation with `TIMEVALUE`.
- The support is bounded to time-only text for the component functions; broader Excel text coercion remains incomplete.
- Currency parsing rejects a currency symbol with no numeric text instead of interpreting an empty remainder as zero.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, Windows runtime checks, app launch, live UI, or Excel GUI checks were run.
