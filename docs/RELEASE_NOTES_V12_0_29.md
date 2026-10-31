# GridForge 12.0.29

Added Excel-style Ctrl+PageUp/Ctrl+PageDown worksheet navigation, Shift+F11 worksheet insertion, and Ctrl+Shift+L filter toggling from the worksheet grid. Each worksheet now remembers its active cell, selected range, and scroll position when switching sheets, so returning to a sheet restores the prior working location.

Verification was limited to the TypeScript check and production build. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
