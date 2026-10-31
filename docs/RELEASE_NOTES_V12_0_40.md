# GridForge 12.0.40

Routine workbook autosaves now write snapshots to IndexedDB rather than synchronously JSON-stringifying and writing the entire workbook to local storage. Existing local-storage workbooks are restored and migrated automatically. If IndexedDB is unavailable or a write fails, GridForge falls back to local storage. Building the workbook snapshot from the in-memory model still runs on the app thread; this change removes the synchronous serialization and storage write in the normal path.

Verification was limited to static TypeScript checking and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
