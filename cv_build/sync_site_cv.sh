#!/usr/bin/env bash
# Mirror ../SeunghanLee_CV.pdf to the user-site repo, which serves the CV
# that index.html links to: https://seunghan96.github.io/assets/pdf/SeunghanLee_CV.pdf
# Called by .git/hooks/pre-push; safe to run by hand. No-op if already in sync.
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
PDF="$HERE/../SeunghanLee_CV.pdf"
SITE_REPO="https://github.com/seunghan96/seunghan96.github.io.git"
TARGET="assets/pdf/SeunghanLee_CV.pdf"

tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

# Blobless sparse clone: fetches only the one PDF, not the whole site.
git clone -q --depth 1 --filter=blob:none --sparse "$SITE_REPO" "$tmp/site"
git -C "$tmp/site" sparse-checkout set --no-cone "/$TARGET"

if cmp -s "$PDF" "$tmp/site/$TARGET"; then
  echo "sync_site_cv: site CV already up to date"
  exit 0
fi

cp "$PDF" "$tmp/site/$TARGET"
git -C "$tmp/site" add "$TARGET"
git -C "$tmp/site" commit -q -m "CV: sync with homepage ($(git -C "$HERE/.." log -1 --format=%h))"
git -C "$tmp/site" push -q origin HEAD
echo "sync_site_cv: pushed updated CV to seunghan96.github.io"
