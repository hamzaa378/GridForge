> **GridForge v10 note:** The current release has 506 registered spreadsheet functions, a persistent Light/Dark switch and Excel familiar/Guided workspace modes. Full v10 behavior and test evidence are in [the v10 release report](V10_RELEASE_EVIDENCE.md). When running with Python alone, use the included compiled `dist/`; Node.js is only needed to rebuild source. Some older sections below retain their historical version numbers.

# GridForge v8 — setup reminder

**Run with Python only:** extract the ZIP and double-click `START_HERE_WINDOWS.bat` (Windows) or run `python run_gridforge.py`. Open the printed local URL. The bundled `dist/` app includes the v8 Data Studio commands; no pip/npm modules are needed for ordinary use. Import `examples/v8_data_studio_demo.json` to try the new tools.

**Develop/test:** Node.js 20+, `npm install`, `npm run dev`. To run full regression, also install Python Playwright and Chromium then run `npm run regression`. **Trusted Excel VBA execution** is optional and needs Windows Excel desktop and `pywin32`, with explicit user approval. Details follow.

---

> **v7 update (2026-10-21):** GridForge can now save edited `.xlsm` files with supported cell changes and preserve the original VBA project. Real VBA execution is available only through an opt-in local Windows bridge using Microsoft Excel desktop and Excel Trust Center settings, NOT in the browser. The older notes below describe the v6 behavior. Read [V7_NATIVE_XLSM_AND_VBA.md](V7_NATIVE_XLSM_AND_VBA.md) for current instructions and verified limits. Live Excel COM execution has not been tested in the Linux authoring environment.

**New v7 beginner path:** Python-only users can use `START_HERE_WINDOWS.bat` for GridForge, XLSX and Save XLSM. To **execute VBA**, separately install licensed desktop Microsoft Excel and optional `pywin32` (`py -3 -m pip install pywin32`), then run `START_VBA_BRIDGE_WINDOWS.bat` and paste its one-time token into Automate. The bridge is never required to run the ordinary app.

# Set up GridForge from scratch

**GridForge v9.0.0 — 406-function GDG Frontend Task 2 release with Excel-style ribbon, onboarding, Data Studio and XLSM support.** Follow **Path A** if the machine already has Python 3 but **no npm, Node.js, pip packages, databases, or project modules installed**. You only need a normal modern browser (Edge, Chrome, or Firefox). All code is contained in the release ZIP. No internet is needed to run the extracted precompiled app.

## Path A — Run the spreadsheet using only Python 3 (recommended for reviewers)

### Windows 10/11

1. Download `GridForge_GDG_Frontend_Task2_FINAL_GitHub_v3_0_2.zip` and right-click **Extract All**. Do **not** run the code directly inside the ZIP.
2. Open the extracted project folder. Confirm you can see `START_HERE_WINDOWS.bat`, `run_gridforge.py`, `dist/`, `README.md`, and `src/`. If they are inside one more folder, open that folder.
3. **Double-click `START_HERE_WINDOWS.bat`.** It checks for the `py` Python launcher, then falls back to `python`.
4. A console prints `Open in your browser: http://127.0.0.1:4174` (or another free port). Your default browser normally opens automatically. Otherwise copy the address into Edge or Chrome.
5. Use the spreadsheet. Select **Import** and choose `examples/sample_sales.xlsx` to see a multi-sheet workbook.
6. Keep the console window open while using GridForge. Press **Ctrl+C** to stop. Your workbook autosaves in browser-local storage for that particular host and port. To preserve it between computers or ports, export a **JSON** backup.

If double-clicking closes unexpectedly, open **PowerShell** inside the extracted folder and run:

```powershell
py -3 run_gridforge.py
```

If `py` is not recognized, use:

```powershell
python run_gridforge.py
```

If 4174 is busy, the Python launcher automatically tries ports 4175–4184. To request a specific port:

```powershell
python run_gridforge.py --port 4190
```

To start without opening the browser automatically:

```powershell
python run_gridforge.py --no-browser
```

**Important:** Python here is a small **static web server**, not a replacement for TypeScript. The included `dist/` folder contains the already compiled JS/CSS/HTML that the browser runs. There is **no Python package installation** and no formula calculation on the Python server.

### macOS / Linux

Extract the ZIP, open a terminal at the project root and run:

```sh
python3 run_gridforge.py
# Or: sh START_HERE_MAC_LINUX.sh
```

Use the URL printed in the terminal.

## Path B — Verify the Python-only package without installing anything

Run from the extracted project root:

```powershell
python run_gridforge.py --check
python scripts/python_smoke.py
```

The smoke test verifies the bundled assets and starts a **temporary localhost server**, checking that HTML, JavaScript and CSS are actually served with expected MIME types. It also verifies a missing URL returns 404. It requires only Python's **standard library**. It does not test the custom formula engine (that requires the Node tests), the actual browser UI, or Excel interoperability.

## Path C — Developer setup from a new machine (install Node.js)

To **change TypeScript source, recompile, and execute the full challenge regression**, install Node.js **20 or later** from https://nodejs.org/. npm is supplied with Node.js. Close and reopen the terminal, then check:

```powershell
node --version
npm --version
python --version
```

In the extracted repository root:

```powershell
npm install
npm run dev
```

Open the printed URL. This rebuilds `dist/` using TypeScript. No backend, database, cloud service, `.env`, API key or account is needed. Only **one** npm development dependency is declared (`typescript`); no spreadsheet parser/evaluator library is installed.

Run **non-browser tests**:

```powershell
npm run regression:core
```

For **all tests including Chromium**, install optional testing dependencies:

```powershell
py -3 -m pip install "playwright>=1.56,<2"
py -3 -m playwright install chromium
npm run regression
```

If Python was installed without `pip`, run `py -3 -m ensurepip --upgrade` or use its installation options. On macOS/Linux replace `py -3` with `python3`. The full regression launches a temporary browser on its own random port; do not start the app in a separate window to test.

Successful output should contain **65 passing Node tests**, the **10,000-formula dependency benchmark**, **14 Chromium checks**, and `REGRESSION PASS: 4/4 stages`. Performance timings vary by computer; test counts can increase in future versions.

## Common issues

| Issue | Solution |
| --- | --- |
| `py` not found | Run `python run_gridforge.py` instead; verify `python --version` reports Python 3. |
| Nothing opens | Read the URL printed in PowerShell and paste it into your browser. |
| `PORT already in use` | Python launcher automatically tries another port; or pass `--port 4190`. |
| `dist/` files missing | Re-extract the complete ZIP, including nested folders. Do not download only `README.md`. |
| Website is blank | Open Developer Tools Console for errors; verify you used `http://127.0.0.1:<port>` rather than double-clicking `dist/index.html` as a file. |
| `npm` not recognized | Not needed for Path A; install Node.js only for Path C. |
| `tsc` or module missing | Run `npm install` in the project root with internet access for initial developer setup. |
| Browser tests lack Playwright | Run the `pip install playwright` and `playwright install chromium` commands under Path C. |
| Saved data differs across ports | Browser storage is **origin-specific**. Export JSON from the old port, then import on the new port. |
| Internet is unavailable | Path A remains fully usable. Path C installs packages and requires access to the npm/Python package registries unless dependencies are already cached. |

## Publishing to GitHub

If `git` is already installed, see [`GITHUB_RELEASE.md`](GITHUB_RELEASE.md) for exact commands. The repository root is the folder holding `README.md`, `src/` and `.github/`. **GitHub does not automatically extract a ZIP uploaded as a file.** Extract it and upload the individual project files to the repository, or use the `git` command-line instructions. The `.github/workflows/` folder must be included for automated checks and deployment.

**What works with Python only:** running the already-built application, manually testing formulas, XLSX imports, spreadsheet tools, exporting data, and the Python-only server smoke test. **What requires additional installations:** rebuilding TypeScript, 65 Node engine tests, the 10,000-formula Node benchmark, and 14 Playwright Chromium checks.
