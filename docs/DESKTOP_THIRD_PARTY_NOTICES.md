# Desktop redistribution and third-party software

GridForge's own source code is provided under `LICENSE` (MIT). The desktop build adds:

- **pywebview 6.2.1** — Windows WebView2 embedding, BSD-style terms. See https://github.com/r0x0r/pywebview
- **PyInstaller 6.22.3** — executable freezer, GPLv2 with special exceptions. See https://pyinstaller.org/en/stable/license.html
- **pywin32** — optional Excel VBA automation bridge. See https://github.com/mhammond/pywin32
- **Inno Setup 6** — Windows installer compiler, separate publisher terms. See https://jrsoftware.org/isinfo.php
- **Microsoft WebView2 Runtime** — an end-user prerequisite managed and licensed by Microsoft, not bundled with GridForge. See Microsoft's WebView2 distribution documentation.

A public distributor should check each dependency's currently applicable license, include required copyright/license notices, and ensure the chosen installer compiler license allows their use case. **A code signing certificate is not included.**

The Excel desktop/VBA bridge requires the user to have a properly installed and licensed Microsoft Excel copy. GridForge does not redistribute Microsoft Excel or its VBA runtime.
