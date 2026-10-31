# GridForge v12.0.79

## Chart image export

- Worksheet chart cards now offer PNG export alongside SVG.
- PNGs are rasterized at 2x chart dimensions with a white background for broad document and presentation compatibility.
- Export errors are reported in the app instead of silently failing.

## Verification

- TypeScript static check passed.
- Installer and archives are built with `-SkipTests`; no app launch, live UI, Windows runtime, or Excel GUI tests were run.
