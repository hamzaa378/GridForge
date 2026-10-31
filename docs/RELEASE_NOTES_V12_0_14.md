# GridForge 12.0.14

CSV import now transfers file bytes to a module worker. The worker parses quoted CSV fields, constructs the worksheet, applies edits in batches, and calculates imported formulas away from the interface thread. It reports progress and returns the completed workbook snapshot to the app. The import keeps the existing 1,000-row by 104-column worksheet cap and now clearly limits source files to 64 MB, 5,000,000 fields, and 1,000,000 nonempty cells.

The app still reads the selected file into memory before transferring it, then restores the worker's finished workbook snapshot on the interface thread. These steps can still briefly use substantial memory for large inputs.

Verification: static TypeScript check, Python syntax compilation, production build, and Windows packaging with `-SkipTests`. No application launch, live UI test, or Excel GUI interoperability check was run.
