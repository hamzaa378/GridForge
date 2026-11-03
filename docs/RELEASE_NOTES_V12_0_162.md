# GridForge v12.0.162 — ODS custom number-format maps

ODS two- and three-section custom number formats now write negative and zero section mappings using OpenDocument's `style:map` element with the `style:condition` and `style:apply-style-name` attributes. Import recognizes the same standard representation and reconstructs the supported section formats.

This fixes interoperability with Calc and other ODF readers: the prior writer placed these map elements and attributes in the number-format namespace, so readers could ignore the alternate sections. Formats with colors, explicit conditional sections, text sections, or unsupported placeholders remain outside the translated subset.

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
