# GridForge 12.0.23

XLSX export now writes `_xlfn.` markers for `PERCENTOF`, `ISOMITTED`, and `DETECTLANGUAGE`, matching Microsoft's current function names and compatibility markers. The importer removes these markers before GridForge evaluates supported formulas.

`PERCENTOF` and `ISOMITTED` are modern Excel functions; `DETECTLANGUAGE` is backed by Microsoft Translation Services and still uses GridForge's offline heuristic implementation locally. Its exported formula therefore retains the Excel function name, but GridForge does not claim equivalent language-detection quality or service behavior.

Verification: TypeScript static check and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.