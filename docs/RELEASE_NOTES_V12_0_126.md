# GridForge v12.0.126 — Safer ChartEx imports

ChartEx import now checks that a specialized chart contains one supported series and that its category and value formulas resolve to matching, connected local ranges. Unsupported layout names and source layouts that GridForge cannot safely represent are reported in the workbook import warnings and skipped, rather than being reduced to a chart with altered data.

This is an import fidelity safeguard. It does not add multi-series ChartEx editing or full Excel chart styling.

Verification is limited to the TypeScript check, production build, and release packaging. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.
