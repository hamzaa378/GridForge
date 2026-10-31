# GridForge 12.0.37

Formula-based PIVOTBY now accepts up to eight fields on each row/column axis, alongside its existing multiple measure columns. It builds tuple-based groups, supports sort selection across the configured field levels or measures, emits hierarchy header rows, and lists each row field in its own output column. Grand-total controls remain available; parent-level subtotals and full Excel `relative_to` hierarchy behavior are still unsupported.

Verification was limited to static TypeScript checking and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
