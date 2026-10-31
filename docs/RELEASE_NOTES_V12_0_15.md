# GridForge 12.0.15

GridForge JSON imports now transfer file bytes to a module worker, where JSON parsing and document restoration take place. The worker reports the current phase and returns the restored snapshot for the application to install. Inputs are limited to 64 MB. This complements the CSV import worker in v12.0.14.

The desktop app still reads the chosen file into memory before transfer and restores the returned snapshot on the interface thread. Large documents can therefore still cause a short pause during those boundary steps.

Verification: static TypeScript check, Python syntax compilation, production build, and Windows packaging with `-SkipTests`. No application launch, live UI test, or Excel GUI interoperability check was run.
