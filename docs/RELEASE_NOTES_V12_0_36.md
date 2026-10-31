# GridForge 12.0.36

The Name Box now accepts cross-sheet cell and range addresses such as `Sheet2!A1:C20` and `'Quarter 1'!$B$2:$D$18`. It switches to a visible target sheet only after validating the address, and looks up supported defined names on that worksheet.

Verification was limited to static TypeScript checking and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
