# GridForge v12.0.151 — Half-point font sizing

- Cell font sizes now accept half-point increments from 6 to 72 points, in addition to whole-point sizes.
- The Home size picker includes half-point choices. ODS font-size properties and XLSX font records retain supported fractional point sizes.
- Grid rendering and AutoFit now convert point sizes to CSS pixels using the 96 dpi display scale, so an 11 pt font displays at about 14.67 px.
- v12.0.150 formula parsing also normalizes localized digits in all locales and handles decimal-comma values with valid three-digit grouping marks.
- Advanced typography remains limited, and the receiving system still needs the requested font installed. Static TypeScript build and Windows installer/source packaging are run without automated or live UI tests.
