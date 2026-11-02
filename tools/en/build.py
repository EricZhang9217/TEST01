#!/usr/bin/env python3
"""Build the English IC deck from JSON slide specs onto the original WuXi master."""
import copy, json, math, os, re, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

E = 914400
SW, SH = 13.3333, 7.5

C_TITLE = RGBColor(0x00, 0x70, 0xC0)
C_DEEP  = RGBColor(0x1F, 0x38, 0x64)
C_HEAD  = RGBColor(0x30, 0x54, 0x96)
C_BODY  = RGBColor(0x26, 0x26, 0x26)
C_MUTE  = RGBColor(0x59, 0x59, 0x59)
C_GRAY  = RGBColor(0x80, 0x80, 0x80)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_TINT  = RGBColor(0xEA, 0xF1, 0xFA)
C_TINT2 = RGBColor(0xF5, 0xF8, 0xFC)
C_LINE  = RGBColor(0xBF, 0xD4, 0xE8)
C_ACC   = RGBColor(0xC5, 0x5A, 0x11)

L, R = 0.56, 0.56
CW = SW - L - R                 # 12.21
TOP_TITLE, H_TITLE = 0.58, 0.78
Y_TAG = 1.21
Y_BODY = 1.56
Y_BODY_END = 6.84
Y_SRC = 6.95
X_SRC, W_SRC = 2.10, 10.20
GAP = 0.22

WARN = []

# ---------- text helpers ----------
def _cpl(width_in, pt):
    """approx characters per line for Arial"""
    return max(8, int(width_in * 72 / (pt * 0.50)))

def est_lines(text, width_in, pt):
    if not text:
        return 1
    n = 0
    for para in str(text).split("\n"):
        n += max(1, math.ceil(len(para) / _cpl(width_in, pt)))
    return n

def est_h(text, width_in, pt, lead=1.22):
    return est_lines(text, width_in, pt) * pt * lead / 72.0

def set_run(r, text, pt, color, bold=False, italic=False):
    r.text = text
    f = r.font
    f.name = "Arial"; f.size = Pt(pt); f.bold = bold; f.italic = italic
    f.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {}); rPr.append(e)
        e.set("typeface", "Arial")

def tb(slide, x, y, w, h, wrap=True):
    s = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = s.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return s, tf

def para(tf, first, text, pt, color, bold=False, space_after=0, bullet=None,
         indent=0.0, align=None, italic=False, line=1.0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    set_run(p.add_run(), text, pt, color, bold, italic)
    p.space_after = Pt(space_after)
    if line != 1.0:
        p.line_spacing = line
    if align:
        p.alignment = align
    pPr = p._p.get_or_add_pPr()
    if indent:
        pPr.set("marL", str(int(indent * E)))
        pPr.set("indent", str(-int(0.16 * E)))
    else:
        pPr.set("marL", "0"); pPr.set("indent", "0")
    if bullet:
        bu = pPr.makeelement(qn("a:buChar"), {"char": bullet})
        clr = pPr.makeelement(qn("a:buClr"), {})
        sc = clr.makeelement(qn("a:srgbClr"), {"val": "0070C0"}); clr.append(sc)
        fnt = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
        for e in (clr, fnt, bu):
            pPr.append(e)
    else:
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))
    return p

def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(0.75)
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.10)
    tf.margin_top = tf.margin_bottom = Inches(0.06)
    return s, tf

# ---------- region renderers: each returns needed height given width ----------
def h_bullets(reg, w):
    h = 0.0
    if reg.get("heading"):
        h += 0.28
    for it in reg["items"]:
        if isinstance(it, str):
            it = {"text": it}
        h += est_h(it["text"], w - 0.22, 10.5) + 0.075
        for s in it.get("sub", []):
            h += est_h(s, w - 0.46, 9.5) + 0.04
    return h

def r_bullets(slide, reg, x, y, w, h):
    if reg.get("heading"):
        s, tf = tb(slide, x, y, w, 0.24)
        para(tf, True, reg["heading"], 11, C_HEAD, bold=True)
        y += 0.30; h -= 0.30
    pt, sub_pt = 10.5, 9.5
    while pt > 8.9:
        tot = 0.0
        for it in reg["items"]:
            if isinstance(it, str):
                it = {"text": it}
            tot += est_h(it["text"], w - 0.22, pt) + 0.075
            for s_ in it.get("sub", []):
                tot += est_h(s_, w - 0.46, pt - 1) + 0.04
        if tot <= h:
            break
        pt -= 0.5; sub_pt = pt - 1
    s, tf = tb(slide, x, y, w, h)
    first = True
    for it in reg["items"]:
        if isinstance(it, str):
            it = {"text": it}
        para(tf, first, it["text"], pt, C_BODY, bullet="▪", indent=0.19, space_after=5, line=1.04)
        first = False
        for s_ in it.get("sub", []):
            para(tf, False, s_, sub_pt, C_MUTE, bullet="–", indent=0.42, space_after=3, line=1.0)

def h_table(reg, w):
    cols = reg["columns"]; rows = reg["rows"]
    cwv = reg.get("col_widths") or [1.0 / len(cols)] * len(cols)
    h = 0.30 if reg.get("heading") else 0.0
    hh = max(est_h(c, w * cwv[i] - 0.16, 8.5) for i, c in enumerate(cols)) + 0.12
    h += max(0.26, hh)
    for row in rows:
        rh = max(est_h(str(c), w * cwv[i] - 0.16, 8.5) for i, c in enumerate(row)) + 0.13
        h += max(0.26, rh)
    return h

def r_table(slide, reg, x, y, w, h):
    if reg.get("heading"):
        s, tf = tb(slide, x, y, w, 0.24)
        para(tf, True, reg["heading"], 11, C_HEAD, bold=True)
        y += 0.30; h -= 0.30
    cols, rows = reg["columns"], reg["rows"]
    cwv = reg.get("col_widths") or [1.0 / len(cols)] * len(cols)
    tot = sum(cwv); cwv = [c / tot for c in cwv]
    pt = 8.5
    while pt > 7.4:
        hh = max(0.26, max(est_h(c, w * cwv[i] - 0.16, pt) for i, c in enumerate(cols)) + 0.12)
        rhs = [max(0.26, max(est_h(str(c), w * cwv[i] - 0.16, pt) for i, c in enumerate(r)) + 0.13) for r in rows]
        if hh + sum(rhs) <= h:
            break
        pt -= 0.5
    nat = hh + sum(rhs)
    if h > nat + 0.05:
        k = min(1.5, h / nat)
        hh *= k; rhs = [r * k for r in rhs]
    gt = slide.shapes.add_table(len(rows) + 1, len(cols), Inches(x), Inches(y),
                                Inches(w), Inches(hh + sum(rhs)))
    tbl = gt.table
    tblPr = tbl._tbl.find(qn("a:tblPr"))
    tblPr.set("firstRow", "1"); tblPr.set("bandRow", "0")
    for e in list(tblPr):
        if e.tag == qn("a:tableStyleId"):
            tblPr.remove(e)
    for i, cwi in enumerate(cwv):
        tbl.columns[i].width = Emu(int(w * cwi * E))
    tbl.rows[0].height = Emu(int(hh * E))
    for j, rh in enumerate(rhs):
        tbl.rows[j + 1].height = Emu(int(rh * E))
    for i, c in enumerate(cols):
        cell = tbl.cell(0, i)
        cell.fill.solid(); cell.fill.fore_color.rgb = C_HEAD
        cell.margin_left = cell.margin_right = Inches(0.07)
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(cell.text_frame, True, str(c), pt, C_WHITE, bold=True, line=1.0)
    fcb = reg.get("first_col_bold")
    for j, row in enumerate(rows):
        for i, c in enumerate(row):
            cell = tbl.cell(j + 1, i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_TINT2 if j % 2 == 0 else C_WHITE
            cell.margin_left = cell.margin_right = Inches(0.07)
            cell.margin_top = cell.margin_bottom = Inches(0.045)
            cell.vertical_anchor = MSO_ANCHOR.TOP
            bold = bool(fcb and i == 0)
            para(cell.text_frame, True, str(c), pt,
                 C_DEEP if bold else C_BODY, bold=bold, line=1.02)

def h_cards(reg, w):
    n = len(reg["items"])
    cw = (w - 0.18 * (n - 1)) / n
    h = 0.30 if reg.get("heading") else 0.0
    body = max(est_h(i["body"], cw - 0.26, 9.5) for i in reg["items"])
    hd = max(est_h(i["heading"], cw - 0.26, 11) for i in reg["items"])
    return h + hd + body + 0.34

def r_cards(slide, reg, x, y, w, h):
    if reg.get("heading"):
        s, tf = tb(slide, x, y, w, 0.24)
        para(tf, True, reg["heading"], 11, C_HEAD, bold=True)
        y += 0.30; h -= 0.30
    items = reg["items"]; n = len(items)
    cw = (w - 0.18 * (n - 1)) / n
    pt = 9.5
    while pt > 8.4:
        need = max(est_h(i["heading"], cw - 0.26, pt + 1.5) + est_h(i["body"], cw - 0.26, pt) for i in items) + 0.34
        if need <= h:
            break
        pt -= 0.5
    ch = min(h, (max(est_h(i["heading"], cw - 0.26, pt + 1.5) + est_h(i["body"], cw - 0.26, pt) for i in items) + 0.34) * 1.8)
    for k, it in enumerate(items):
        cx = x + k * (cw + 0.18)
        s, tf = rect(slide, cx, y, cw, ch, fill=C_TINT2, line=C_LINE,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.045)
        tf.margin_top = Inches(0.10)
        para(tf, True, it["heading"], pt + 1.5, C_DEEP, bold=True, space_after=3, line=1.0, align=PP_ALIGN.LEFT)
        para(tf, False, it["body"], pt, C_BODY, line=1.05, align=PP_ALIGN.LEFT)

def h_flow(reg, w):
    n = len(reg["steps"])
    cw = (w - 0.14 * (n - 1)) / n
    h = 0.30 if reg.get("heading") else 0.0
    return h + 0.30 + max(est_h(s["heading"], cw - 0.24, 9.5) for s in reg["steps"]) \
             + max(est_h(s.get("body", ""), cw - 0.24, 8.5) for s in reg["steps"]) + 0.30

def r_flow(slide, reg, x, y, w, h):
    if reg.get("heading"):
        s, tf = tb(slide, x, y, w, 0.24)
        para(tf, True, reg["heading"], 11, C_HEAD, bold=True)
        y += 0.30; h -= 0.30
    steps = reg["steps"]; n = len(steps)
    cw = (w - 0.14 * (n - 1)) / n
    pt = 9.0
    while pt > 7.6:
        need = 0.34 + max(est_h(s["heading"], cw - 0.24, pt + 0.5) for s in steps) \
                    + max(est_h(s.get("body", ""), cw - 0.24, pt - 0.5) for s in steps) + 0.26
        if need <= h:
            break
        pt -= 0.5
    ch = min(h, (0.34 + max(est_h(s["heading"], cw - 0.24, pt + 0.5) for s in steps)
                      + max(est_h(s.get("body", ""), cw - 0.24, pt - 0.5) for s in steps) + 0.26) * 1.5)
    for k, st in enumerate(steps):
        cx = x + k * (cw + 0.14)
        shp, tf = rect(slide, cx, y, cw, ch, fill=C_TINT, line=C_LINE,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        tf.margin_top = Inches(0.09)
        para(tf, True, str(k + 1), pt - 0.5, C_TITLE, bold=True, space_after=2, line=1.0, align=PP_ALIGN.LEFT)
        para(tf, False, st["heading"], pt + 0.5, C_DEEP, bold=True, space_after=2, line=1.0, align=PP_ALIGN.LEFT)
        if st.get("body"):
            para(tf, False, st["body"], pt - 0.5, C_MUTE, line=1.04, align=PP_ALIGN.LEFT)
        if k < n - 1:
            a = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                       Inches(cx + cw + 0.018), Inches(y + ch / 2 - 0.055),
                                       Inches(0.105), Inches(0.11))
            a.rotation = 90
            a.fill.solid(); a.fill.fore_color.rgb = C_LINE
            a.line.fill.background(); a.shadow.inherit = False

def h_kpis(reg, w):
    h = 0.30 if reg.get("heading") else 0.0
    n = len(reg["items"])
    cw = (w - 0.16 * (n - 1)) / n if n <= 2 else (w - 0.16) / 2
    lab = max(est_h(i["label"], cw - 0.26, 9) for i in reg["items"])
    rows = 1 if n <= 2 else math.ceil(n / 2)
    return h + rows * (0.46 + lab + 0.20) + (rows - 1) * 0.12

def r_kpis(slide, reg, x, y, w, h):
    if reg.get("heading"):
        s, tf = tb(slide, x, y, w, 0.24)
        para(tf, True, reg["heading"], 11, C_HEAD, bold=True)
        y += 0.30; h -= 0.30
    items = reg["items"]; n = len(items)
    per = n if n <= 2 else (4 if n == 4 and w > 8.0 else 2)
    rows = math.ceil(n / per)
    cw = (w - 0.16 * (per - 1)) / per
    lab = max(est_h(i["label"], cw - 0.26, 9) for i in items)
    ch = min((h - 0.12 * (rows - 1)) / rows, (0.46 + lab + 0.20) * 1.7)
    for k, it in enumerate(items):
        cx = x + (k % per) * (cw + 0.16)
        cy = y + (k // per) * (ch + 0.12)
        s, tf = rect(slide, cx, cy, cw, ch, fill=C_TINT, line=None,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        tf.margin_top = Inches(0.08)
        para(tf, True, it["value"], 21, C_TITLE, bold=True, space_after=1, line=0.94, align=PP_ALIGN.LEFT)
        para(tf, False, it["label"], 9, C_MUTE, line=1.03, align=PP_ALIGN.LEFT)

def h_callout(reg, w):
    return max(0.44, est_h(reg["text"], w - 0.60, 10.5) + 0.26)

def r_callout(slide, reg, x, y, w, h):
    pt = 10.5
    while pt > 8.9 and est_h(reg["text"], w - 0.60, pt) + 0.26 > h:
        pt -= 0.5
    s, tf = rect(slide, x, y, w, h, fill=C_TINT, line=None)
    tf.margin_left = Inches(0.30); tf.margin_right = Inches(0.26)
    tf.margin_top = Inches(0.11); tf.margin_bottom = Inches(0.09)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para(tf, True, reg["text"], pt, C_DEEP, bold=True, line=1.06, align=PP_ALIGN.LEFT)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.055), Inches(h))
    bar.fill.solid(); bar.fill.fore_color.rgb = C_TITLE
    bar.line.fill.background(); bar.shadow.inherit = False

def h_rows(reg, w):
    h = 0.30 if reg.get("heading") else 0.0
    for it in reg["items"]:
        h += max(0.52, est_h(it["body"], w - 1.35, 10) + 0.34) + 0.10
    return h

def r_rows(slide, reg, x, y, w, h):
    if reg.get("heading"):
        s, tf = tb(slide, x, y, w, 0.24)
        para(tf, True, reg["heading"], 11, C_HEAD, bold=True)
        y += 0.30; h -= 0.30
    items = reg["items"]
    pt = 10.0
    while pt > 8.4:
        need = sum(max(0.52, est_h(i["body"], w - 1.35, pt) + 0.34) + 0.10 for i in items)
        if need <= h:
            break
        pt -= 0.5
    cy = y
    for it in items:
        rh = max(0.52, est_h(it["body"], w - 1.35, pt) + 0.34)
        s, tf = rect(slide, x, cy, w, rh, fill=C_TINT2, line=None)
        num = rect(slide, x, cy, 0.52, rh, fill=C_TITLE, line=None)[0]
        ntf = num.text_frame
        ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(ntf, True, it.get("num", ""), 15, C_WHITE, bold=True, align=PP_ALIGN.CENTER, line=1.0)
        s2, tf2 = tb(slide, x + 0.70, cy + 0.11, w - 0.90, rh - 0.18)
        para(tf2, True, it["heading"], pt + 1.5, C_DEEP, bold=True, space_after=2, line=1.0)
        para(tf2, False, it["body"], pt, C_BODY, line=1.05)
        cy += rh + 0.10

HFN = {"bullets": h_bullets, "table": h_table, "cards": h_cards, "flow": h_flow,
       "kpis": h_kpis, "callout": h_callout, "rows": h_rows}
RFN = {"bullets": r_bullets, "table": r_table, "cards": r_cards, "flow": r_flow,
       "kpis": r_kpis, "callout": r_callout, "rows": r_rows}

# ---------- slide assembly ----------
def add_title(slide, text, section):
    s, tf = tb(slide, L, TOP_TITLE, CW, H_TITLE)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    pt = 18 if len(text) <= 78 else (16.5 if len(text) <= 104 else 15)
    para(tf, True, text, pt, C_TITLE, bold=True, line=1.02)
    if section:
        s2, tf2 = tb(slide, SW / 2, Y_TAG, CW / 2 - 0.0, 0.24)
        para(tf2, True, "/  " + section.upper(), 8.5, C_GRAY, bold=True, align=PP_ALIGN.RIGHT)

def add_source(slide, src):
    if not src:
        return
    s, tf = tb(slide, X_SRC, Y_SRC, W_SRC, 0.30)
    para(tf, True, src, 7.5, C_GRAY, line=1.0)

def place(slide, regions, label):
    avail = Y_BODY_END - Y_BODY
    lay = spec_layout
    if lay in ("split", "split_left_wide", "split_right_wide") and len(regions) >= 2:
        frac = {"split": 0.5, "split_left_wide": 0.60, "split_right_wide": 0.40}[lay]
        lw = CW * frac - 0.16
        rw = CW * (1 - frac) - 0.16
        bottom = regions[2] if len(regions) > 2 else None
        bh = HFN[bottom["type"]](bottom, CW) if bottom else 0.0
        if bottom:
            bh = min(bh, 1.5)
        colh = avail - (bh + GAP if bottom else 0)
        for reg, xx, ww in ((regions[0], L, lw), (regions[1], L + CW * frac + 0.16, rw)):
            need = HFN[reg["type"]](reg, ww)
            if need > colh + 0.02:
                WARN.append(f"{label}: {reg['type']} needs {need:.2f}in in {colh:.2f}in column")
            RFN[reg["type"]](slide, reg, xx, Y_BODY, ww, colh)
        if bottom:
            RFN[bottom["type"]](slide, bottom, L, Y_BODY + colh + GAP, CW, bh)
    else:
        needs = [HFN[r["type"]](r, CW) for r in regions]
        total = sum(needs) + GAP * (len(regions) - 1)
        if total > avail:
            scale = (avail - GAP * (len(regions) - 1)) / sum(needs)
            needs = [n * scale for n in needs]
            WARN.append(f"{label}: stack over by {total - avail:.2f}in, compressed")
        slack = avail - (sum(needs) + GAP * (len(regions) - 1))
        gap = GAP
        if slack > 0.2:
            if len(regions) > 1:
                add = min(0.42, slack / (len(regions) - 1))
                gap = GAP + add
                slack -= add * (len(regions) - 1)
            tot = sum(needs)
            needs = [n + slack * n / tot for n in needs]
        yy = Y_BODY
        for reg, nh in zip(regions, needs):
            RFN[reg["type"]](slide, reg, L, yy, CW, nh)
            yy += nh + gap

def make_divider(prs, lay, spec, num):
    slide = prs.slides.add_slide(lay)
    s, tf = tb(slide, L, 2.05, 2.0, 1.6)
    para(tf, True, str(num), 54, RGBColor(0xD6, 0xE4, 0xF2), bold=True, line=0.95)
    s, tf = tb(slide, 2.3, 2.46, 9.9, 0.95)
    para(tf, True, spec["title"], 30, C_TITLE, bold=True, line=1.02)
    if spec.get("subtitle"):
        s, tf = tb(slide, 2.3, 3.62, 9.4, 0.9)
        para(tf, True, spec["subtitle"], 13, C_MUTE, line=1.15)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(3.42), Inches(1.5), Inches(0.035))
    bar.fill.solid(); bar.fill.fore_color.rgb = C_TITLE
    bar.line.fill.background(); bar.shadow.inherit = False
    return slide

def set_notes(slide, text):
    if not text:
        return
    tf = slide.notes_slide.notes_text_frame
    tf.text = text

# ---------- main ----------
def main():
    specs = json.load(open("deck_spec.json"))
    prs = Presentation("deck_v2.pptx")
    content_lay = prs.slide_layouts[12]     # 7_Title/标题
    keep_orig = set(int(x) for x in sys.argv[1:]) if len(sys.argv) > 1 else set()

    cov = prs.slides[0]
    cov_txt = ["AOC NewCo  |  Investment Committee discussion materials",
               "Decision requested: a capped preparation period, released gate by gate on evidence",
               "2026-11-02      |      Ruochen Zhang, M.D.      |      Confidential - internal use only"]
    k = 0
    for shp in cov.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip() and k < 3:
            ptf = shp.text_frame
            keep = ptf.paragraphs[0].runs[0] if ptf.paragraphs[0].runs else None
            sz = keep.font.size if keep is not None and keep.font.size else Pt(14)
            bold = keep.font.bold if keep is not None else False
            try:
                col = keep.font.color.rgb
            except Exception:
                col = C_BODY
            for p in list(ptf.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            p0 = ptf.paragraphs[0]
            for r in list(p0.runs):
                r._r.getparent().remove(r._r)
            set_run(p0.add_run(), cov_txt[k], sz.pt, col, bold)
            k += 1

    global spec_layout
    n_div = 0
    new_slides = []
    for sp in specs:
        if sp.get("kind") == "divider":
            n_div += 1
            sl = make_divider(prs, content_lay, sp, n_div)
            new_slides.append(sl)
            continue
        sl = prs.slides.add_slide(content_lay)
        label = sp.get("title", "")[:48]
        add_title(sl, sp["title"], sp.get("section"))
        spec_layout = sp.get("layout", "full")
        regs = sp.get("regions", [])
        if regs:
            place(sl, regs, label)
        add_source(sl, sp.get("source"))
        set_notes(sl, sp.get("notes"))
        new_slides.append(sl)

    # reorder: cover (orig slide 1) + new slides, then kept originals (hidden)
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    n_orig = len(ids) - len(new_slides)
    orig_ids = ids[:n_orig]
    new_ids = ids[n_orig:]
    for e in ids:
        sldIdLst.remove(e)
    sldIdLst.append(orig_ids[0])            # cover stays first
    for e in new_ids:
        sldIdLst.append(e)
    for i, e in enumerate(orig_ids[1:], start=2):
        if i in keep_orig:
            sldIdLst.append(e)

    # hide the kept original pages
    for i, e in enumerate(list(sldIdLst)):
        if i == 0:
            continue
        if i > len(new_ids):
            pass
    for e in list(sldIdLst)[1 + len(new_ids):]:
        rid = e.get(qn("r:id"))
        part = prs.part.rels[rid].target_part
        part._element.set("show", "0")

    prs.save("deck_en.pptx")
    print("slides:", len(list(sldIdLst)), "| new:", len(new_ids), "| kept originals:", len(list(sldIdLst)) - 1 - len(new_ids))
    for w in WARN:
        print("WARN", w)

if __name__ == "__main__":
    main()
