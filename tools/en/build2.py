#!/usr/bin/env python3
"""Build the English deck using the ORIGINAL deck's page architecture.

Per-page frame (inches), copied from the Chinese original:
  title      0.56 0.60 12.09 0.70
  section    right-aligned, y 1.05
  lead strip 0.56 1.33 12.09 0.66
  figure     0.56  y=6.32-h   w x h   (h <= 4.24, w <= 8.60, native aspect)
  caption    0.56 6.34 figw 0.18
  points     x=0.56+figw+0.34, y 2.08, w, 2.70
  memo note  same x,            y 4.92, w, 1.44
  limits     0.56 6.44 12.09 0.46
  source     2.10 6.93 10.20 0.30
"""
import json, math, os, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

import build as B
from build import (tb, para, rect, est_h, set_run, C_TITLE, C_DEEP, C_HEAD, C_BODY,
                   C_MUTE, C_GRAY, C_WHITE, C_TINT, C_TINT2, C_LINE, C_ACC,
                   HFN, RFN, make_divider, set_notes, E, SW, SH)

L, CW = 0.56, 12.09
RIGHT = 12.65
Y_TITLE, H_TITLE = 0.60, 0.70
Y_TAG = 1.05
Y_LEAD, H_LEAD = 1.33, 0.66
Y_FIG_BOT = 6.32
Y_COL = 2.08
Y_NOTE = 4.92
Y_LIM, H_LIM = 6.44, 0.46
Y_SRC = 6.93

WARN = []


def add_frame(slide, sp):
    s, tf = tb(slide, L, Y_TITLE, CW, H_TITLE)
    t = sp["title"]
    pt = 18 if len(t) <= 78 else (16.5 if len(t) <= 104 else 15)
    para(tf, True, t, pt, C_TITLE, bold=True, line=1.02)
    if sp.get("section"):
        s, tf = tb(slide, SW / 2, Y_TAG, RIGHT - SW / 2, 0.24)
        para(tf, True, "/  " + sp["section"].upper(), 8.5, C_GRAY, bold=True, align=PP_ALIGN.RIGHT)
    if sp.get("lead"):
        lead = sp["lead"]
        lpt = 10.5
        while lpt > 8.5 and est_h(lead, CW - 0.60, lpt) > H_LEAD - 0.14:
            lpt -= 0.5
        s, tf = rect(slide, L, Y_LEAD, CW, H_LEAD, fill=C_TINT, line=None)
        tf.margin_left = Inches(0.28); tf.margin_right = Inches(0.22)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tf, True, lead, lpt, C_DEEP, bold=True, line=1.06, align=PP_ALIGN.LEFT)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(L), Inches(Y_LEAD),
                                     Inches(0.055), Inches(H_LEAD))
        bar.fill.solid(); bar.fill.fore_color.rgb = C_TITLE
        bar.line.fill.background(); bar.shadow.inherit = False
    if sp.get("limits"):
        lim = sp["limits"]
        ytop = max(Y_LIM, sp.get("_limtop", 0.0))
        hlim = max(0.22, Y_SRC - 0.03 - ytop)
        lpt = 7.5
        while lpt > 6.0 and sum(est_h(x, CW - 0.2, lpt) for x in lim) > hlim:
            lpt -= 0.25
        s, tf = tb(slide, L, ytop, CW, hlim)
        for i, x in enumerate(lim):
            para(tf, i == 0, x, lpt, C_MUTE, bullet="·", indent=0.12, space_after=1, line=1.05)
    if sp.get("source"):
        s, tf = tb(slide, 2.10, Y_SRC, 10.20, 0.30)
        para(tf, True, sp["source"], 7, C_GRAY, line=1.0)


def add_figure(slide, path, caption):
    with Image.open(path) as im:
        iw, ih = im.size
    asp = iw / ih
    h = min(4.24, 8.45 / asp)
    w = h * asp
    y = Y_COL + (4.24 - h) / 2.0
    slide.shapes.add_picture(path, Inches(L), Inches(y), Inches(w), Inches(h))
    cap_bot = 0.0
    if caption:
        cw = max(w, 6.6)
        cpt = 7.0
        while cpt > 5.8 and est_h(caption, cw, cpt) > 0.13:
            cpt -= 0.25
        cy = max(y + h + 0.03, 6.34)
        ch = est_h(caption, cw, cpt)
        s, tf = tb(slide, L, cy, cw, max(0.13, ch))
        para(tf, True, caption, cpt, C_GRAY, line=1.0)
        cap_bot = cy + ch + 0.03
    return w, cap_bot


def add_points(slide, x, w, points, note, start=1):
    if points and w > 5.9 and len(points) >= 3:
        half = (w - 0.3) / 2
        mid = (len(points) + 1) // 2
        add_points(slide, x, half, points[:mid], None, 1)
        add_points(slide, x + half + 0.3, half, points[mid:], None, mid + 1)
        if note:
            add_points(slide, x, w, None, note)
        return
    if points:
        h = (5.56 - Y_COL) if note else (Y_FIG_BOT - Y_COL)
        pt = 10
        while pt > 8.0 and sum(est_h(p, w - 0.34, pt) + 0.14 for p in points) > h:
            pt -= 0.5
        s, tf = tb(slide, x, Y_COL, w, h)
        for i, p in enumerate(points):
            pp = para(tf, i == 0, "", pt, C_BODY, space_after=9, line=1.1)
            set_run(pp.runs[0] if pp.runs else pp.add_run(), f"{i+start}   ", pt, C_TITLE, True)
            set_run(pp.add_run(), p, pt, C_BODY, False)
    if note:
        ytop = 5.66 if points else Y_COL
        h = Y_FIG_BOT - ytop
        npt = 8.5
        while npt > 7.0 and 0.22 + sum(est_h(n, w - 0.3, npt) + 0.07 for n in note) > h:
            npt -= 0.5
        s, tf = tb(slide, x, ytop, w, h)
        para(tf, True, "DETAIL IN THE MEMO", 7.5, C_GRAY, bold=True, space_after=4)
        for n in note:
            para(tf, False, n, npt, C_MUTE, bullet="·", indent=0.12, space_after=3, line=1.05)


def main():
    specs = json.load(open("deck_spec2.json"))
    prs = Presentation("deck_v2.pptx")
    lay = prs.slide_layouts[12]
    keep = set(int(x) for x in sys.argv[1:]) if len(sys.argv) > 1 else set()

    # English cover
    cov = prs.slides[0]
    cov_txt = ["AOC NewCo  |  Investment Committee discussion materials",
               "Decision requested: a capped preparation period, released gate by gate on evidence",
               "2026-11-02      |      Ruochen Zhang, M.D.      |      Confidential - internal use only"]
    k = 0
    for shp in cov.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip() and k < 3:
            ptf = shp.text_frame
            keep_r = ptf.paragraphs[0].runs[0] if ptf.paragraphs[0].runs else None
            sz = keep_r.font.size if keep_r is not None and keep_r.font.size else Pt(14)
            bold = keep_r.font.bold if keep_r is not None else False
            try:
                col = keep_r.font.color.rgb
            except Exception:
                col = C_BODY
            for p in list(ptf.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            p0 = ptf.paragraphs[0]
            for r in list(p0.runs):
                r._r.getparent().remove(r._r)
            set_run(p0.add_run(), cov_txt[k], sz.pt, col, bold)
            k += 1

    n_div = 0
    new = []
    for sp in specs:
        if sp.get("kind") == "divider":
            n_div += 1
            new.append(make_divider(prs, lay, sp, n_div))
            continue
        sl = prs.slides.add_slide(lay)
        new.append(sl)
        set_notes(sl, sp.get("notes"))
        label = sp["title"][:46]

        fig = sp.get("figure")
        if fig and os.path.exists(fig):
            figw, cap_bot = add_figure(sl, fig, sp.get("fig_caption"))
            sp["_limtop"] = cap_bot
            x = L + figw + 0.34
            w = RIGHT - x
            if w > 3.45:          # keep the column narrow; the figure keeps the rest
                x = RIGHT - 3.30
                w = 3.30
            if w < 1.9:
                WARN.append(f"{label}: right column only {w:.2f}in")
            else:
                add_points(sl, x, w, sp.get("points"), sp.get("table_note"))
        else:
            if fig:
                WARN.append(f"{label}: missing figure {fig}")
            regs = sp.get("regions") or []
            if not regs and sp.get("table"):
                regs = [dict(sp["table"], type="table")]
            if regs:
                B.spec_layout = sp.get("layout", "full")
                top = Y_COL if sp.get("lead") else Y_LEAD
                bot = Y_FIG_BOT + (0.38 if not sp.get("limits") else 0.0)
                B.Y_BODY, B.Y_BODY_END = top, bot
                B.place(sl, regs, label)
                WARN.extend(B.WARN[len(WARN):] if False else [])
            pts = sp.get("points")
            if pts and not regs:
                add_points(sl, L, CW, pts, sp.get("table_note"))

        add_frame(sl, sp)

    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    n_orig = len(ids) - len(new)
    orig_ids, new_ids = ids[:n_orig], ids[n_orig:]
    for e in ids:
        sldIdLst.remove(e)
    sldIdLst.append(orig_ids[0])
    for e in new_ids:
        sldIdLst.append(e)
    for i, e in enumerate(orig_ids[1:], start=2):
        if i in keep:
            sldIdLst.append(e)
    for e in list(sldIdLst)[1 + len(new_ids):]:
        prs.part.rels[e.get(qn("r:id"))].target_part._element.set("show", "0")

    prs.save("deck_en2.pptx")
    print("total", len(list(sldIdLst)), "| visible", 1 + len(new_ids),
          "| hidden", len(list(sldIdLst)) - 1 - len(new_ids))
    for w in set(WARN) | set(B.WARN):
        print("WARN", w)


if __name__ == "__main__":
    main()
