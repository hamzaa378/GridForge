# GridForge v12.0.143 — Independent cell border edges

- Each selected cell edge can now carry its own border line style, weight, and color. The Home ribbon applies edge changes while retaining the other sides.
- Mixed four-edge border styles and colors now round-trip through XLSX and ODS. Older single-edge and all-edge formatting remains supported.
- The border-weight feature from v12.0.141 remains available, including thin, medium, and thick lines.

Diagonal borders, diagonal border weights, and advanced table-border behavior remain unsupported. XLSX simplifies thick dashed and dotted combinations to its nearest supported line preset.

Build: `npm run build`. No automated tests, app launch, live UI, or Office/LibreOffice checks were run.
