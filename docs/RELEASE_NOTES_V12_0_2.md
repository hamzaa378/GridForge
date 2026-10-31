# GridForge 12.0.2

- ChartEx charts now preserve the configured data-label position and separator during XLSX export and import, alongside the existing value/category/series visibility flags.
- Supported ChartEx positions are center, inside end, outside end, and best fit. Percentage visibility and other positions still require standard chart output or a future model extension.
- Carries forward the v12.0.1 fixed ribbon sizing, `TOCOL` / `TOROW` error-ignore correction, and yielding XLSX/CSV/query import improvements described in [v12.0.1 release notes](RELEASE_NOTES_V12_0_1.md).
- `npm run check` and the Windows package build passed with test phases skipped. No live UI, Windows runtime, or Excel GUI checks were run.
