# GitHub publishing and deployment guide

## 1. Local preflight

From the folder containing `package.json`:

```powershell
node --version
npm install
npm run check
npm test
npm run benchmark
npm run build
npm start
```

Open `http://localhost:4173`. If another process is using port 4173, set `PORT` in your terminal before `npm start` (PowerShell: `$env:PORT=4180`). The included `dist/` is meant for immediate demos; always run `npm run build` after changing source.

## 2. Create your GitHub repository

Create an **empty** GitHub repository, e.g. `gridforge-spreadsheet`, with no automatically generated README, then from the extracted folder:

```powershell
git init
git add .
git commit -m "GridForge v2 submission"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git push -u origin main
```

The repository URL above is **only a placeholder**. Replace it with your actual repository's URL. If GitHub says the remote already exists, run `git remote -v` and update it with `git remote set-url origin <actual-url>`. Authentication is handled by your configured Git credential manager or GitHub SSH setup. Never commit a password, token, or `.env` file.

## 3. Check CI and deploy GitHub Pages

The included Actions workflow `.github/workflows/ci.yml` checks TypeScript, runs tests, records benchmark results, and uploads a compiled-site artifact. `.github/workflows/pages.yml` builds and deploys the static site.

In your repo: **Settings → Pages → Build and deployment → Source → GitHub Actions**. Then visit **Actions** and verify both workflows. Once Pages finishes, find the actual published URL under **Settings → Pages** or the deployment environment. The app uses `./styles.css` and `./main.js` relative paths, which work under repository subpaths; no API, server database, or environment keys are necessary.

If Actions fail because dependencies cannot be installed, verify npm registry accessibility in Actions and that `npm install` works. This source snapshot intentionally does not contain a generated package-lock file because npm registry access was unavailable in the packaging environment; the workflow uses `npm install`, not `npm ci`. For stricter dependency pinning after installing locally, commit the resulting `package-lock.json`, then change workflows to `npm ci`.

## 4. Demo to reviewers (2–4 minutes)

1. Open the demo sheet. Change B4 from 125 to 225. Show that D4 changes from 4000 to 7200 and the recalculation counter stays selective.
2. Select D4; show Inspector AST, dependencies, and Impact panel.
3. Add another sheet. Enter `42` in A1. Switch back and show the original demo is intact; duplicate and rename a sheet.
4. Select an example cell and apply Bold, Percent/Number/Currency, alignment or color; demonstrate undo. Pick a range and sort ascending/descending, then undo.
5. Open Find, search for `P-103`, jump to the matching cell. Show Replace All with formula protection.
6. Select A4:B11; open Analyze to show sums, blank/error counts and live chart.
7. Open What-if; preview a changed value and show the original value remains intact. Export JSON, then reimport it to show all tabs and formatting persist.
8. Show tests, the benchmark report, and the public repository's CI status.

## 5. Final checklist

- [ ] README loads and preview image displays in the GitHub repository.
- [ ] `npm run check`, `npm test`, and `npm run benchmark` pass on your machine.
- [ ] GitHub Actions CI is green and Pages is configured correctly.
- [ ] Live URL loads without a missing stylesheet or `.js` 404.
- [ ] Show at least one useful error diagnostic (`=SUM(A1, B2`) and cycle (`=C1` in C1).
- [ ] Validate copy/paste and drag-fill in a regular Chrome or Edge session.
- [ ] Confirm JSON export and reimport preserve multiple sheets and formatting.
- [ ] Add actual GitHub / live-demo URLs to your submission form; do not claim a deployment before checking it.
