#!/usr/bin/env bash
# One-shot deploy as a PROJECT page on the existing seunghan96 account.
# Resulting URL: https://seunghan96.github.io/seunghanlee
# (Your existing seunghan96.github.io stays untouched.)
#
# Prerequisites (human-only steps):
#   1. On the seunghan96 account, create an empty repo named: seunghanlee (Public)
#   2. Have a Personal Access Token (PAT) ready — github.com > Settings >
#      Developer settings > Personal access tokens > "Generate new token"
#      (classic) with the "repo" scope. Use it as the password when git asks.
set -e
cd "$(dirname "$0")"

git add -A
git commit -m "Update homepage" || echo "(nothing new to commit)"
git branch -M main

# When prompted: Username = seunghan96, Password = <your PAT>
git push -u origin main

echo
echo "Pushed. Now enable Pages: repo > Settings > Pages > Source: main / root."
echo "Site will be live at https://seunghan96.github.io/seunghanlee within a minute or two."
