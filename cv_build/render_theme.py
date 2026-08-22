# -*- coding: utf-8 -*-
"""Render the CV .docx to PDF under a chosen visual theme.

Content comes entirely from the Word file; a theme only changes typography,
colour, heading treatment and page furniture.

usage: render_theme.py THEME out.pdf [in.docx]
"""
import io
import os
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Flowable, Frame, HRFlowable, Image, KeepTogether,
    PageBreak, PageTemplate, Paragraph as P, Spacer, Table as RLTable, TableStyle,
)

HERE = os.path.dirname(os.path.abspath(__file__))
DL = os.path.join(HERE, "fonts_dl")

GULIM = dict(r="gulim_regular.ttf", b="gulim_bold.ttf",
             i="gulim_italic.ttf", bi="gulim_bolditalic.ttf")
LATO = dict(r=DL + "/Lato-Regular.ttf", b=DL + "/Lato-Bold.ttf",
            i=DL + "/Lato-Italic.ttf", bi=DL + "/Lato-BoldItalic.ttf")
PTSERIF = dict(r=DL + "/PT_Serif-Web-Regular.ttf", b=DL + "/PT_Serif-Web-Bold.ttf",
               i=DL + "/PT_Serif-Web-Italic.ttf", bi=DL + "/PT_Serif-Web-Bold.ttf")

THEMES = {
    # A — faithful to the Word original
    "classic": dict(
        fonts=GULIM, heading_fonts=None, scale=1.0, leading=1.398 * 1.15,
        accent="#1F497D", accent2="#FF9900", muted="#212121", link="#0000FF",
        heading="rule", heading_size=17, entry_underline=True,
        bullets={1: ("•", "SYM", 10.0), 2: ("o", "CO", 10.0), 3: ("§", "WI", 10.0)},
        footer=False, band=False, margins=(1.9, 1.9, 2.2, 2.2), table_head="#D9E2F3",
    ),
    # B — modern sans, accent bars, airy
    "modern": dict(
        fonts=LATO, heading_fonts=None, scale=0.87, leading=1.42,
        accent="#1F5C8B", accent2="#C4761A", muted="#4A4A4A", link="#1F5C8B",
        heading="bar", heading_size=15.5, entry_underline=False,
        bullets={1: ("▪", None, 7.5), 2: ("–", None, 9.5), 3: ("·", None, 10.0)},
        footer=True, band=False, margins=(2.0, 1.8, 1.9, 1.9), table_head="#EAF0F6",
    ),
    # C — editorial serif
    "serif": dict(
        fonts=PTSERIF, heading_fonts=LATO, scale=0.86, leading=1.45,
        accent="#7A2E2E", accent2="#A9762B", muted="#4A4A4A", link="#7A2E2E",
        heading="caps", heading_size=13.5, entry_underline=False,
        bullets={1: ("•", None, 9.0), 2: ("–", None, 9.5), 3: ("·", None, 10.0)},
        footer=True, band=False, margins=(2.2, 2.0, 2.0, 2.0), table_head="#F1E9E4",
    ),
    # D — header band + sidebar rule, compact
    "banner": dict(
        fonts=LATO, heading_fonts=LATO, scale=0.85, leading=1.36,
        accent="#0F3D57", accent2="#C4761A", muted="#4A4A4A", link="#0F3D57",
        heading="band", heading_size=14.5, entry_underline=False,
        bullets={1: ("•", None, 9.0), 2: ("◦", None, 8.5), 3: ("–", None, 9.0)},
        footer=True, band=True, margins=(1.9, 1.7, 1.8, 1.8), table_head="#E6EDF2",
    ),
    # E — the banner Word file, whose sizes/colours are already baked in
    "banner_src": dict(
        fonts=LATO, heading_fonts=LATO, scale=1.0, leading=1.36,
        accent="#0F3D57", accent2="#C4761A", muted="#4A4A4A", link="#0F3D57",
        heading="band", heading_size=14.5, entry_underline=False,
        bullets={1: ("•", None, 8.5), 2: ("\u25E6", None, 8.0), 3: ("–", None, 8.5)},
        footer=True, band=True, margins=(1.9, 1.7, 1.8, 1.8), table_head="#E6EDF2",
    ),
}

theme_name = sys.argv[1]
OUT = sys.argv[2]
DOCX = sys.argv[3] if len(sys.argv) > 3 else "/nfsdata/home/seunghan.lee/web/SeunghanLee_CV.docx"
T = THEMES[theme_name]

pfx = theme_name[:2].upper()
FR, FB, FI, FBI = pfx + "R", pfx + "B", pfx + "I", pfx + "BI"
for nm, key in ((FR, "r"), (FB, "b"), (FI, "i"), (FBI, "bi")):
    path = T["fonts"][key]
    pdfmetrics.registerFont(TTFont(nm, path if os.path.isabs(path) else os.path.join(HERE, path)))
pdfmetrics.registerFontFamily(FR, normal=FR, bold=FB, italic=FI, boldItalic=FBI)

HR, HB = FR, FB
if T["heading_fonts"]:
    HR, HB = pfx + "HR", pfx + "HB"
    pdfmetrics.registerFont(TTFont(HR, T["heading_fonts"]["r"]))
    pdfmetrics.registerFont(TTFont(HB, T["heading_fonts"]["b"]))

SYMBOL_FONTS = {}
if theme_name == "classic":
    for nm, fn in (("SYM", "sym_bullet.ttf"), ("CO", "sym_o.ttf"), ("WI", "sym_wing.ttf")):
        pdfmetrics.registerFont(TTFont(nm, os.path.join(HERE, fn)))
        SYMBOL_FONTS[nm] = nm

ACCENT, ACCENT2 = HexColor(T["accent"]), HexColor(T["accent2"])
MUTED, LINK = HexColor(T["muted"]), HexColor(T["link"])
COLOR_MAP = {  # colours used in the docx -> theme colours
    "1F497D": T["accent"], "212121": T["muted"], "FF9900": T["accent2"],
    "333333": T["muted"], "0000FF": T["link"], "222222": T["muted"],
}
DOC_BULLETS = {"•": 1, "o": 2, "§": 3, "\u25E6": 2, "–": 3}

if theme_name != "classic":  # bullet glyphs must exist in the theme font
    from fontTools.ttLib import TTFont as _TT

    _cmap = set(_TT(T["fonts"]["r"]).getBestCmap())
    for _lvl, (_ch, _f, _sz) in list(T["bullets"].items()):
        if ord(_ch) not in _cmap:
            T["bullets"][_lvl] = ("\u2022", _f, _sz)

doc = Document(DOCX)
S = T["scale"]
try:
    MARK_PT = doc.styles["Normal"].font.size.pt
except Exception:
    MARK_PT = 12.0


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def rgb_of(run):
    try:
        c = run.font.color
        if c is not None and c.rgb is not None:
            return "%02X%02X%02X" % (c.rgb[0], c.rgb[1], c.rgb[2])
    except Exception:
        pass
    return None


def _mk_run(el, par):
    from docx.text.run import Run

    return Run(el, par)


def para_pieces(par):
    rels = par.part.rels
    for child in par._p.iterchildren():
        if child.tag == qn("w:r"):
            yield _mk_run(child, par), None
        elif child.tag == qn("w:hyperlink"):
            rid = child.get(qn("r:id"))
            url = rels[rid].target_ref if rid in rels else None
            for r in child.findall(qn("w:r")):
                yield _mk_run(r, par), url


def run_markup(run, base_size, url=None, force=None, heading=False):
    t = run.text
    if not t.strip() and not t:
        return ""
    size = (run.font.size.pt if run.font.size is not None else base_size) * S
    fr, fb, fi, fbi = (HR, HB, FI, FBI) if heading else (FR, FB, FI, FBI)
    face = (fbi if run.italic else fb) if run.bold else (fi if run.italic else fr)
    col = rgb_of(run)
    col = COLOR_MAP.get(col, "#" + col if col else None)
    if force:
        col = force
    out = '<font face="%s" size="%.2f"%s>%s</font>' % (
        face, size, ' color="%s"' % col if col else "", esc(t))
    if run.font.superscript:
        out = "<super>%s</super>" % out
    if run.underline and (T["entry_underline"] or url):
        out = "<u>%s</u>" % out
    if url:
        out = '<a href="%s">%s</a>' % (esc(url), out)
    return out


def para_image(par):
    blips = par._p.findall(".//" + qn("a:blip"))
    if not blips:
        return None
    part = par.part.rels[blips[0].get(qn("r:embed"))].target_part
    ext = par._p.findall(".//" + qn("wp:extent"))
    cx = int(ext[0].get("cx")) if ext else 3600000
    cy = int(ext[0].get("cy")) if ext else 3600000
    return Image(io.BytesIO(part.blob), width=cx / 914400.0 * 72, height=cy / 914400.0 * 72)


def has_page_break(par):
    return any(br.get(qn("w:type")) == "page" for br in par._p.findall(".//" + qn("w:br")))


def bottom_border(par):
    pPr = par._p.pPr
    if pPr is None:
        return None
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        return None
    bot = bdr.find(qn("w:bottom"))
    if bot is None or bot.get(qn("w:val")) in (None, "none", "nil"):
        return None
    return True


def pf_pt(v, d=0.0):
    return v.pt if v is not None else d


def line_spacing_of(par):
    ls = par.paragraph_format.line_spacing
    return ls if isinstance(ls, float) else 1.15


class AccentBar(Flowable):
    """Small colour block used as a heading marker."""

    def __init__(self, w, h, color, gap=0):
        Flowable.__init__(self)
        self.width, self.height, self.color, self.gap = w, h, color, gap

    def wrap(self, aw, ah):
        return (self.width, self.height + self.gap)

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, self.gap, self.width, self.height, stroke=0, fill=1)


def heading_flowables(par, text_flow):
    """Decorate a section heading according to the theme."""
    style = T["heading"]
    if style == "rule":
        return [text_flow, HRFlowable(width="100%", thickness=1.5, color=Color(0, 0, 0),
                                      spaceBefore=1, spaceAfter=4, lineCap="square")]
    if style == "bar":
        return [AccentBar(46, 2.6, ACCENT, gap=2), text_flow,
                Spacer(1, 1)]
    if style == "caps":
        return [text_flow, HRFlowable(width="100%", thickness=0.6, color=ACCENT,
                                      spaceBefore=2, spaceAfter=6)]
    if style == "band":
        return [text_flow, HRFlowable(width="26%", thickness=2.2, color=ACCENT2,
                                      spaceBefore=2, spaceAfter=6, hAlign="LEFT")]
    return [text_flow]


def style_for(par, size, leading, bullet=None, heading=False):
    pf = par.paragraph_format
    align = {None: 0, 0: 0, 1: 1, 2: 2, 3: 4}.get(
        pf.alignment.real if hasattr(pf.alignment, "real") else pf.alignment, 0)
    kw = dict(
        name="p", fontName=HR if heading else FR, fontSize=size, leading=leading,
        leftIndent=pf_pt(pf.left_indent) * (0.92 if S < 1 else 1),
        firstLineIndent=0 if bullet else pf_pt(pf.first_line_indent),
        spaceBefore=pf_pt(pf.space_before) * (1.15 if S < 1 else 1),
        spaceAfter=pf_pt(pf.space_after), alignment=align,
        allowWidows=1, allowOrphans=1, linkUnderline=1,
    )
    if bullet:
        kw.update(bulletIndent=(pf_pt(pf.left_indent) + pf_pt(pf.first_line_indent)) * (0.92 if S < 1 else 1),
                  bulletFontName=bullet[1], bulletFontSize=bullet[2])
    return ParagraphStyle(**kw)


def render_par(par, heading=False):
    pieces = list(para_pieces(par))
    bullet = None
    if pieces:
        first = pieces[0][0].text.strip()
        if first in DOC_BULLETS:
            lvl = DOC_BULLETS[first]
            ch, fnt, sz = T["bullets"][lvl]
            bullet = (ch, SYMBOL_FONTS.get(fnt, FR), sz * (1 if theme_name == "classic" else 1))
            pieces = pieces[1:]
    sizes = [r.font.size.pt for r, _ in pieces if r.font.size is not None and r.text.strip()]
    base = max(sizes) if sizes else 12.0
    if heading:
        base = T["heading_size"] / S
    leading = T["leading"] * max(base, MARK_PT) * S
    markup = "".join(run_markup(r, base, u, heading=heading) for r, u in pieces)
    if heading and T["heading"] == "caps":
        markup = markup.replace("</font>", "</font>")
    if not markup.strip():
        return Spacer(1, leading * 0.8)
    return P(markup, style_for(par, base * S, leading, bullet, heading),
             bulletText=bullet[0] if bullet else None)


def cell_flowables(cell):
    out = []
    for par in cell.paragraphs:
        img = para_image(par)
        out.append(img if img is not None else render_par(par))
    return out or [Spacer(1, 1)]


def shading_of(cell):
    tcPr = cell._tc.tcPr
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    fill = shd.get(qn("w:fill"))
    if not fill or fill == "auto":
        return None
    return HexColor(T["table_head"]) if fill.upper() == "D9E2F3" else HexColor("#" + fill)


def render_table(tbl):
    nrows, ncols = len(tbl.rows), len(tbl.columns)
    grid = [[tbl.cell(r, c) for c in range(ncols)] for r in range(nrows)]
    data, shades, extent = [], [], {}
    for r in range(nrows):
        row = []
        for c in range(ncols):
            cell = grid[r][c]
            key = id(cell._tc)
            if key in extent:
                e = extent[key]
                e[2], e[3] = max(e[2], r), max(e[3], c)
                row.append("")
            else:
                extent[key] = [r, c, r, c]
                row.append(cell_flowables(cell))
                sh = shading_of(cell)
                if sh is not None:
                    shades.append(("BACKGROUND", (c, r), (c, r), sh))
        data.append(row)
    spans = [(c0, r0, c1, r1) for r0, c0, r1, c1 in extent.values() if (r0, c0) != (r1, c1)]
    widths = []
    for c in range(ncols):
        w = grid[0][c].width or tbl.columns[c].width
        widths.append(w.pt * (0.99 if S < 1 else 1) if w is not None else None)
    has_grid = tbl.style is not None and "Grid" in (tbl.style.name or "")
    pad = 4 if has_grid else 0
    cmds = [("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), pad), ("RIGHTPADDING", (0, 0), (-1, -1), pad),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5 if has_grid else 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5 if has_grid else 0)]
    if has_grid:
        if theme_name == "classic":
            cmds.append(("GRID", (0, 0), (-1, -1), 0.6, Color(0.45, 0.45, 0.45)))
        else:  # horizontal rules only — lighter, more editorial
            cmds += [("LINEBELOW", (0, 0), (-1, -1), 0.4, HexColor("#C9CFD6")),
                     ("LINEABOVE", (0, 0), (-1, 0), 0.9, ACCENT),
                     ("LINEBELOW", (0, 0), (-1, 0), 0.9, ACCENT)]
    for c0, r0, c1, r1 in spans:
        cmds += [("SPAN", (c0, r0), (c1, r1)), ("VALIGN", (c0, r0), (c1, r1), "MIDDLE")]
    cmds += shades
    t = RLTable(data, colWidths=widths, repeatRows=1 if has_grid else 0, hAlign="LEFT")
    t.setStyle(TableStyle(cmds))
    return t


def classify(par):
    if bottom_border(par):
        return "heading"
    runs = [r for r, _ in para_pieces(par) if r.text.strip()]
    if not runs:
        return "blank"
    first = runs[0]
    if pf_pt(par.paragraph_format.left_indent) > 0 or par.text.lstrip().startswith("-"):
        return "cont"
    size = first.font.size.pt if first.font.size is not None else 12.0
    if first.bold and (size >= 12.5 or first.underline):
        return "entry"
    return "other"


items = []
for child in doc.element.body.iterchildren():
    if child.tag == qn("w:p"):
        par = Paragraph(child, doc)
        if has_page_break(par):
            items.append(("break", [PageBreak()]))
            if not par.text.strip():
                continue
        img = para_image(par)
        if img is not None:
            items.append(("other", [img]))
            continue
        kind = classify(par)
        if kind == "heading":
            items.append((kind, heading_flowables(par, render_par(par, heading=True))))
        else:
            items.append((kind, [render_par(par)]))
    elif child.tag == qn("w:tbl"):
        items.append(("table", [render_table(Table(child, doc))]))

MAX_GROUP = 11
story, i = [], 0
while i < len(items):
    kind, fl = items[i]
    if kind == "heading":
        group, j, taken = list(fl), i + 1, 0
        while j < len(items) and items[j][0] in ("entry", "cont", "other", "blank") and taken < 3:
            group += items[j][1]
            j += 1
            taken += 1
        story.append(KeepTogether(group))
        i = j
    elif kind == "entry":
        group, j, taken = list(fl), i + 1, 0
        while j < len(items) and items[j][0] == "cont" and taken < MAX_GROUP:
            group += items[j][1]
            j += 1
            taken += 1
        story.append(KeepTogether(group) if len(group) > 1 else group[0])
        i = j
    else:
        story.extend(fl)
        i += 1

from reportlab.lib.units import cm

ml, mr, mt, mb = T["margins"]


def furniture(canvas, docu):
    canvas.saveState()
    if T["band"]:
        canvas.setFillColor(ACCENT)
        canvas.rect(0, A4[1] - 0.42 * cm, A4[0], 0.42 * cm, stroke=0, fill=1)
    if T["footer"]:
        canvas.setFont(FR, 7.6)
        canvas.setFillColor(HexColor("#8A8F96"))
        canvas.drawString(ml * cm, mb * cm * 0.45, "Seunghan Lee · Curriculum Vitae")
        canvas.drawRightString(A4[0] - mr * cm, mb * cm * 0.45, "%d" % docu.page)
    canvas.restoreState()


pdf = BaseDocTemplate(OUT, pagesize=A4,
                      leftMargin=ml * cm, rightMargin=mr * cm,
                      topMargin=mt * cm, bottomMargin=mb * cm,
                      title="Seunghan Lee — CV", author="Seunghan Lee")
pdf.addPageTemplates([PageTemplate(id="all", onPage=furniture, frames=[Frame(
    pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height,
    leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="body")])])
pdf.build(story)
print("saved", OUT, "(theme: %s)" % theme_name)
