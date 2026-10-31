# GridForge 12.0.3

- XLSX export now writes the configured data-bar fill mode (`solid` or `gradient`) and optional negative-value fill color. These settings were already available in GridForge's conditional-formatting editor and importer but were omitted on export.
- Charts now support a left-side legend in the editor, on-sheet renderer, and supported XLSX import/export paths.
- Carries forward v12.0.2 ChartEx data-label position and separator round-tripping, plus v12.0.1 ribbon sizing, formula correction, and import responsiveness improvements.
- `npm run check` and Windows packaging passed with test phases skipped. No automated tests, Windows app launch, live UI, or Excel GUI checks were run.
- Scope: this preserves the supported data-bar properties; Excel's additional axis-position, border, direction, and advanced threshold settings are not represented in GridForge yet.
