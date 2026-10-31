# GridForge 12.0.13

CSV export now checks the selected sheet's full used rectangle before expanding it. It stops with an actionable message when the area exceeds 2,000,000 cells or serialized text exceeds the size guard. JSON and CSV exports also enforce the text-size guard. This protects against a single distant cell causing an accidental enormous CSV allocation.

These safeguards complement the background export workers introduced in v12.0.11 and v12.0.12. Workbook snapshot creation and the native save dialog handoff still use the interface thread; this release does not eliminate every possible pause on very large workbooks.

Verification: static TypeScript check, Python syntax compilation, production build and Windows packaging with `-SkipTests`. No application launch, live UI test, or Excel GUI interoperability check was run.
