# GridForge desktop source release — actual verification

**Test environment:** Linux x86-64 with Node.js 22, TypeScript 5.8.3, Python, and headless Chromium. This is the v10.4 core spreadsheet release plus a desktop shell and native Save As bridge; the Windows PyInstaller build is not executable or verifiable in this Linux environment.

| Check | Result | Scope |
| --- | --- | --- |
| TypeScript compilation | PASS | Existing engine + desktop-aware Save As call |
| GridForge formula and workbook unit tests | **834/834 PASS** | Original v10.4 engine, Excel-style UI model, XLSX/XLSM components |
| Desktop Python unit tests | **5/5 PASS** | File allowlist, atomic Save As, loopback HTTP, asset checks, workflow metadata |
| Native Save As browser integration | PASS | XLSX export bridged to stubbed pywebview JS API; valid ZIP payload preserved |
| Existing Excel-style ribbon Chromium smoke | **19/19 PASS** | All tabs, themes, shortcuts, connection-warning and JS error checks |
| Python-only precompiled site smoke | PASS | Local HTTP app and JavaScript MIME types verified |
| 10,000-formula dependency benchmark | PASS | Source edit affected 26 formulas; isolated edit 1 |
| Python `--check` prebuilt assets | PASS | Desktop resources found |
| Actual `GridForge.exe` on Windows | **NOT YET VERIFIED** | Requires Windows GitHub Actions or Windows x64 build |
| Inno Setup installer build/install/uninstall | **NOT YET VERIFIED** | Workflow includes silent automated installer smoke; must run on Windows |
| WebView2 native GUI, file dialogs, persistence | **NOT YET VERIFIED** | Manual Windows acceptance required |
| Desktop Excel VBA macro run | **NOT YET VERIFIED** | Requires desktop Microsoft Excel installed and trusted macro |

The release does not claim native Windows parity from Linux tests. Additional instructions: [`DESKTOP_BUILD_AND_RELEASE.md`](DESKTOP_BUILD_AND_RELEASE.md).
