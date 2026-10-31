# GridForge v12.0.63

- `FILTERXML` now uses the browser's offline XML DOM and XPath evaluator for scalar queries and node selections. Node selections return a vertical array for the existing spill engine.
- XML input is limited to 1 MB, XPath expressions to 1,024 characters, and node results to 20,000 values. Malformed XML, invalid expressions, and unmatched node queries return `#VALUE!`.
- Non-browser hosts retain a bounded single-tag fallback. This implementation does not provide unrestricted XPath parity across all Excel runtime environments.
- Static TypeScript check and production build passed. Installer and portable release packages use `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks are run.
