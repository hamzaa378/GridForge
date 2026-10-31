# GridForge v12.0.118

## Range-aware formula tracing

- Precedent tracing expands compact local range dependencies into highlighted worksheet cells, up to 50,000 cells.
- The dependency inspector continues to show each range as a compact reference and provides a bounded sample of its cell values.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
