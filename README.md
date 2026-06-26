# seunghanlee.github.io

Personal academic homepage for **Seunghan Lee** (이승한).
Single static page (`index.html`, no build step) styled after the Jon Barron template.

## Deploy to https://seunghanlee.github.io

GitHub user-site URLs are `https://<username>.github.io`, so the repo must live under a
GitHub account literally named **`seunghanlee`**, and be named **`seunghanlee.github.io`**.

```bash
# from this folder
git init
git add -A
git commit -m "Initial homepage"
git branch -M main
git remote add origin https://github.com/seunghanlee/seunghanlee.github.io.git
git push -u origin main
```

Then in the repo: **Settings → Pages → Source: `main` / root**. Live in a minute or two.

## Edit

- Profile photo: drop a square image at `assets/profile.jpg` (falls back to "SL" initials).
- Everything else (bio, news, publications, service) is plain HTML in `index.html`.
