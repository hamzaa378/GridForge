# GridForge v12.0.127 — Compact Draw and Page Layout ribbons

The Draw and Page Layout toolbars now enforce a fixed 96 px height directly as app children. This prevents workspace and appearance styles from stretching the command area over the worksheet, addressing the oversized blank toolbar region visible in the supplied screenshot.

Verification is limited to the TypeScript check, production build, and release packaging. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.
