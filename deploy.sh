#!/usr/bin/env bash
# One-shot deploy for seunghanlee.github.io
# Prerequisites (human-only steps):
#   1. Create a GitHub account named exactly: seunghanlee
#   2. Create an empty repo named:           seunghanlee.github.io  (Public)
#   3. Have a Personal Access Token (PAT) ready — github.com > Settings >
#      Developer settings > Personal access tokens > "Generate new token"
#      (classic) with the "repo" scope. Use it as the password when git asks.
set -e
cd "$(dirname "$0")"

git add -A
git commit -m "Update homepage" || echo "(nothing new to commit)"
git branch -M main

# When prompted: Username = seunghanlee, Password = <your PAT>
git push -u origin main

echo
echo "Pushed. Now enable Pages: repo > Settings > Pages > Source: main / root."
echo "Site will be live at https://seunghanlee.github.io within a minute or two."
