#!/usr/bin/env bash
# Zip each office-* skill folder into dist/.zip ready for upload to
# LibreChat (POST /api/skills/import accepts .zip with SKILL.md at root
# or one level deep).
set -euo pipefail

cd "$(dirname "$0")"
mkdir -p dist
rm -f dist/*.zip

for dir in office-docx office-xlsx office-pptx; do
  [ -d "$dir" ] || continue
  (cd "$dir" && zip -qr "../dist/${dir}.zip" .)
  echo "built dist/${dir}.zip"
done
