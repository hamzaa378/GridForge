# GridForge v12.0.191 — ODS border width fidelity

ODS border import now converts widths expressed in points, pixels, centimeters, millimeters, and inches before mapping them to GridForge’s thin, medium, and thick border weights. It also recognizes ODF's named thin, medium, and thick widths. The same parser handles cell edges and diagonal borders.

GridForge stores three border weight levels, so ODS line widths are still approximated to those levels. The model also shares one style, weight, and color across both diagonal directions; ODS files with different diagonal-line settings cannot be represented exactly.

Verification: TypeScript check passed. Production build and Windows packaging have not yet been run for v12.0.191. No automated tests, live UI, or LibreOffice GUI round-trip were run.
