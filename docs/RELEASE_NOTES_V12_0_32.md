# GridForge 12.0.32

The Name Box now jumps to a cell, selects a rectangular range such as `A1:C20`, and resolves a supported defined name on the active worksheet. Invalid, out-of-bounds, and multi-colon input leaves the current selection unchanged and displays guidance.

Verification was limited to static TypeScript checking and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
