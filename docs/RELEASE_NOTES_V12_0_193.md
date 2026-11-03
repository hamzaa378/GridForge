# GridForge v12.0.193 — ODS chart axis and formula import fidelity

ODS chart axis styles now remain attached when the only configured option is a number format, category reversal, label rotation, or logarithmic scale. Previously these settings could be written into ODF styles but omitted from the chart axis unless a bound or interval was also present.

ODS formula import now strips `COM.MICROSOFT.` only when it prefixes a function call. This avoids changing sheet or defined-name identifiers that happen to contain the same text.

These are interoperability improvements; chart styling remains bounded by GridForge's chart model and Calc application round-tripping was not performed.

Verification: TypeScript check and production build passed. Windows installer and portable packaging with `-SkipTests` completed successfully. No automated tests, live UI, or LibreOffice GUI checks were run.

