# GitHub-ready submission: GridForge v3.0

## Checklist before pushing

1. Ensure the extracted ZIP has `package.json`, `src/`, `tests/`, `public/`, `dist/`, `docs/`, `examples/`, `.github/workflows/`, `LICENSE`, `.gitignore` at the top level.
2. Install Node.js 20+ and run `npm install && npm test && npm run benchmark` on your Windows machine. Expect 57 passing tests and 26 affected cells in the demonstration input-change experiment. Compiler and results may differ with hardware.
3. Run `npm run dev`, open the printed local URL. Test a real `.xlsx` through Import/Export and ensure the Import Report is visible.
4. Run the steps in `docs/V3_ACCEPTANCE_CHECKLIST.md`. Capture screenshots/GIF for reviewers if desired.
5. Ensure no credentials, private datasets, `node_modules/`, `.env` or temporary files were added. Node/browser uses no external spreadsheet engine; dependencies are dev-only TypeScript.

## Push to empty repository

```powershell
git init
git add .
git commit -m "GridForge v3.0: XLSX interoperability and Excel-style tools"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/GridForge.git
git push -u origin main
```

Replace the remote URL with your actual repository. Create it first, without auto-generated README. If Git warns `remote origin already exists`, inspect `git remote -v` and use `git remote set-url origin ...` only if needed. Do not push to an unrelated repository.

## Deployment

The source contains `.github/workflows/ci.yml` (typecheck, tests, benchmark and compiled artifact) and `.github/workflows/pages.yml` (tests and static-site publication). In GitHub **Settings → Pages**, set source to **GitHub Actions** and wait for the Actions `Deploy GridForge to GitHub Pages` run. The resulting Pages URL is returned by GitHub; no deployment URL is pre-filled because it cannot be known before you publish it.

## Reviewer demo (2–3 minutes)

Start with the bundled `examples/sample_sales.xlsx`, import into GridForge and close the Import Report. Show two tabs and a formula `=B2*C2`. Filter the Sales header table to `Mouse` and show all. Apply a live greater-than threshold, change a referenced value and watch the highlight. Freeze the top row and first column while scrolling. Export as `.xlsx`, import the output and inspect preserved formulas/styles. End with the Formula Inspector and dependency benchmark output. Explain that formulas and recalc are custom implemented and that Excel interoperability is a **supported subset**, not all Excel features.

## GitHub Actions caveats

GitHub's CI runner installs the declared development tool `typescript` with `npm install`; it has Internet access for that. The optional Playwright/Chromium smoke test is not part of CI unless you explicitly install Chromium in the workflow. The included compiled `dist/` lets reviewers run the app with `node scripts/serve.mjs` even without installing dependencies. If GitHub Pages URLs need a non-root base path, the app loads its `./main.js`, `./styles.css` and relative JS imports from the deployment directory.
