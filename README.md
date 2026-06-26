# seunghanlee.github.io

Personal academic homepage for **Seunghan Lee** (이승한).
Single static page (`index.html`, no build step) styled after the Jon Barron template.

## Deploy as a project page on the existing seunghan96 account

Keeps `seunghan96.github.io` untouched and serves this site at a sub-path.
Just create one new **Public** repo named `seunghanlee` on the `seunghan96` account,
then run `./deploy.sh` (or the steps below).

```bash
# from this folder (remote already set to github.com/seunghan96/seunghanlee)
git add -A
git commit -m "Update homepage"
git push -u origin main   # Username: seunghan96   Password: <your PAT>
```

Then in the repo: **Settings → Pages → Source: `main` / root**.
Live at **https://seunghan96.github.io/seunghanlee** in a minute or two.

## Edit

- Profile photo: drop a square image at `assets/profile.jpg` (falls back to "SL" initials).
- Everything else (bio, news, publications, service) is plain HTML in `index.html`.
