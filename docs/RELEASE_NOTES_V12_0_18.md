# GridForge 12.0.18

Explicit `@` implicit intersection now follows Excel's documented rule for computed arrays and returns their top-left item. Direct cell and range references retain GridForge's existing position-aware behavior: a row or column range intersects the formula cell, and a rectangle returns its overlapping cell.

This narrows one dynamic-array compatibility gap. Full array coercion, reference-returning functions, and broader Excel differential coverage remain incomplete.

Verification: `npm run check` and the production Windows build with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.
