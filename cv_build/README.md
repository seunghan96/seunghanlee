# CV build toolkit

Regenerates `SeunghanLee_CV.docx` / `SeunghanLee_CV.pdf` (the "banner" design).

```bash
cd cv_build
python3 build_cv.py                     # content -> ../SeunghanLee_CV.docx (Gulim source)
python3 render_theme.py banner ../cv_candidates/CV_banner.pdf
python3 restyle_banner.py               # -> ../cv_candidates/SeunghanLee_CV_banner.docx (Lato embedded)
cp ../cv_candidates/CV_banner.pdf ../SeunghanLee_CV.pdf
cp ../cv_candidates/SeunghanLee_CV_banner.docx ../SeunghanLee_CV.docx
```

- All CV content lives in `build_cv.py` — edit there, never the .docx/.pdf.
- `render_theme.py` themes: `classic` (Gulim, the Word original), `modern`, `serif`,
  `banner` (shipped), `banner_src`. Only `banner` is used.
- Fonts: `fonts_dl/` (Lato, PT Serif — open licences). The `classic` theme also needs
  `gulim_*.ttf` in this directory; those are proprietary subsets and are gitignored,
  so a fresh clone can build `banner` but not `classic`.
- Needs `python-docx`, `reportlab`, `fontTools` (and `pymupdf` to inspect the result).
