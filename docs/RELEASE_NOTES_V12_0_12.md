# GridForge 12.0.12

CSV and JSON exports now run their heavy work in a module worker, alongside XLSX export introduced in 12.0.11. CSV formulas are calculated from a restored workbook snapshot in the worker and serialized as display values; JSON formatting also happens there. The save controls show a preparation status while work is in progress.

The workbook snapshot is still created on the interface thread, as is the final desktop file-dialog handoff. Very large workbooks can therefore still pause briefly during snapshot creation or native download encoding. This update reduces export work on the interface thread; it does not guarantee that all operations remain responsive at every workbook size.

Verification for this source release: TypeScript code check, Python syntax compilation, and Windows packaging with `-SkipTests`. No application launch, Windows UI test, or Excel GUI interoperability test was run.
