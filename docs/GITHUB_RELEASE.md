> **GridForge v10 note:** The current release has 506 registered spreadsheet functions, a persistent Light/Dark switch and Excel familiar/Guided workspace modes. Full v10 behavior and test evidence are in [the v10 release report](V10_RELEASE_EVIDENCE.md). When running with Python alone, use the included compiled `dist/`; Node.js is only needed to rebuild source. Some older sections below retain their historical version numbers.

# GridForge v9 release instructions

Extract the `GridForge_v9_406_Functions_Excel_UI_GitHub_Ready.zip` archive and upload its **root contents** to a new public GitHub repository. Do not include cache directories or local credentials. The existing workflow under `.github/workflows/` runs Node tests, the Python-only launcher, Chromium regression, XLSM binary tests and local VBA protocol checks. To deploy choose GitHub Pages → GitHub Actions. The Python-only launcher is for local use; Pages serves `dist/`. Python-only use requires no pip packages. Source rebuild and tests require Node.js. See `V9_RELEASE_EVIDENCE.md` and README.

---

# Latest build: GridForge v8

Upload the **contents of the root folder** after extracting `GridForge_v8_100_Data_Studio_GitHub_Ready.zip` (not a directory containing another nested repository). Keep `dist/`, `.github/workflows/`, `src/`, `scripts/`, `examples/`, and `docs/`. Avoid `node_modules`, `__pycache__`, local secrets, browser profiles, `.git`, and temporary logs. Github Pages serves the static compiled `dist/` folder. CI runs the full v8 regression suite (447 unit tests, 100 actions, 26 browser checks and benchmarks).

---

# GitHub submission and deployment — GridForge v5.0.0

The final ZIP contains **the repository files themselves** at archive root; there is no extra top-level parent directory. GitHub will **not extract an uploaded ZIP automatically**. First extract it to a folder and upload/push its contents. Preserve hidden `.github/workflows/` files.

## Upload via Git (recommended)

Prerequisites: Git installed (`git --version`). You do **not** need Git for running the spreadsheet with Python; Git is only one way to upload to GitHub.

1. Create a **public, empty** repository on GitHub, without pre-creating a README or .gitignore.
2. In the extracted root (where `package.json` and `README.md` exist), run:

```powershell
git init
git add .
git commit -m "GridForge v5.0.0: final GDG submission"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/GridForge.git
git push -u origin main
```

Replace the URL with the actual new repository URL. If Git requests identity, configure `git config --global user.name "Your Name"` and `git config --global user.email "you@example.com"` with **your own** identity. GitHub authentication normally uses browser or a credential helper, not your password.

## Without Git installed

You can use GitHub's website (**Add file → Upload files**) or GitHub Desktop. Extract the archive first and upload the **project contents**, not the ZIP as a single file. Carefully preserve nested folders, especially `.github/workflows`, `src/engine`, `dist/engine`, `tests`, `docs`, and `examples`. Git/GitHub Desktop is recommended for a project with this many files.

## Confirm the repository is correct

The **root** must show `README.md`, `package.json`, `src`, `dist`, `docs`, `examples`, `run_gridforge.py`, `START_HERE_WINDOWS.bat`, `.github/workflows`, `.gitignore`, and `LICENSE`. There should be no `node_modules/`, private datasets, browser profiles, passwords or `.env` secrets. Check GitHub Actions after pushing.

## Continuous integration

- `.github/workflows/ci.yml`: installs Node, Python Playwright and Chromium; runs TypeScript check, 241 Node tests, 10,000-formula benchmark and 21 browser checks via `npm run regression`. Also runs Python-only static server smoke validation. Actual totals may increase if tests are added.
- `.github/workflows/pages.yml`: builds and publishes `dist/` as a static app after tests.

GitHub-hosted actions have to download their development/testing tools; **that does not mean end users need them to run the precompiled app**.

## Publish on GitHub Pages

1. Open the repository → **Settings → Pages**.
2. Under **Build and deployment**, choose **GitHub Actions** as Source.
3. Open **Actions → Deploy GridForge to GitHub Pages** and wait for the run to complete.
4. Use the published URL shown by GitHub. There is **no verified live URL** until deployment finishes.
5. On the live site, manually test input, formula evaluation, `examples/sample_sales.xlsx` import, and workbook download. The app uses relative asset paths and runs fully in the browser.

## Recommended README details in submission

State that the engine is custom TypeScript, has 216 functions and dependency-driven recalculation, documents the limited Excel interoperability, and explains the VBA-safe XLSM and JSON action-macro limitations, and includes both a Python-only launch for non-developers and a full Node/Playwright regression for developers. Link the acceptance and function reference docs. Do not promise full Excel compatibility.

## Release verification

See `docs/TECHNICAL_COMPLIANCE.md`, `docs/V5_RELEASE_EVIDENCE.md` and `docs/SETUP_FROM_ZERO.md`. The earlier Windows log recorded 65 unit tests and 14 browser checks for the v3 baseline. The v5 engine has new functions and safe macro behavior, with 241 unit tests and 21 browser checks passing in the packaging environment. Real `.xlsm` compatibility and the browser storage reload should be tested on Windows before final submission.
