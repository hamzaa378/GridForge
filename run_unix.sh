#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
if ! command -v node >/dev/null 2>&1; then
  echo 'Node.js 20+ is required. Install it from https://nodejs.org/' >&2
  exit 1
fi
if [ ! -f dist/index.html ]; then
  echo 'Run npm install and npm run build first.' >&2
  exit 1
fi
echo 'Starting GridForge; use the URL printed below (Ctrl+C to stop)'
node scripts/serve.mjs
