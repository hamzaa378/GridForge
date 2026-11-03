# GridForge v12.0.167 — ODS native drawing import

ODS import now converts cell-anchored OpenDocument lines, rectangles, ellipses, and text boxes into editable GridForge worksheet drawings. It restores supported stroke/fill colors and widths, text, and position using the anchor cell plus row/column dimensions. Import is capped at 300 objects per sheet and coordinates are checked against GridForge's drawing limits.

Images, arbitrary SVG/vector paths, connectors, and detailed line/arrow styles remain unsupported as editable objects. GridForge-authored drawings continue to restore from their namespaced metadata; ODS-exported drawings also remain visible as embedded SVG previews in other ODF readers.

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
