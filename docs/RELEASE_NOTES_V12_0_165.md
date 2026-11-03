# GridForge v12.0.165 — ODS formula fallback boundaries

ODS formula export now checks the function tree before writing OpenFormula. Formulas containing the ten registered service-dependent functions or Excel-only dynamic-array/connected functions outside the OpenFormula subset are exported as their cached values. GridForge retains the original formula text in its namespaced cell attribute, so it restores when the ODS is reopened in GridForge. Nested unsupported calls cause the full formula to use the safe cached-value fallback.

The OpenFormula language allows vendor extensions, but those extensions do not guarantee another spreadsheet application can evaluate a vendor function. This change avoids writing known Excel-only function names as if they were portable OpenFormula. Supported formulas continue to use native ODF syntax. See the [OpenFormula 1.4 specification](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part4-formula/OpenDocument-v1.4-os-part4-formula.html).

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
