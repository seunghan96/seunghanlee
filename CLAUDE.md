# Homepage + CV: always in sync (mandatory)

Any change to `index.html` (papers, news, interests, bio, ...) must be mirrored
in the CV in the same commit:

1. Edit `cv_build/build_cv.py` to match, rebuild per `cv_build/README.md`, and
   commit `index.html`, `cv_build/build_cv.py`, `SeunghanLee_CV.pdf` and
   `SeunghanLee_CV.docx` together.
2. `git push` to origin/main. The pre-push hook runs `cv_build/sync_site_cv.sh`,
   which copies the PDF to the **seunghan96.github.io** repo
   (`assets/pdf/SeunghanLee_CV.pdf`). That is the file the homepage's CV link opens.
3. Verify both: https://seunghan96.github.io/seunghanlee/ and
   https://seunghan96.github.io/assets/pdf/SeunghanLee_CV.pdf (check last-modified).

The local hooks in `.git/hooks/` enforce this. pre-commit blocks a commit that
stages `index.html` without the CV files; use `SKIP_CV_CHECK=1` only for
styling-only edits. Don't add a Co-Authored-By trailer to commits.
