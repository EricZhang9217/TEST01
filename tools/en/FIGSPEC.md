# Rebuilding the deck's figures in English

We are producing an English version of a Chinese investment-committee deck. The earlier attempt replaced
the original diagrams with plain tables; the owner rejected that. **Keep each figure a figure** — same
visual structure, same reading order, same colour coding as the Chinese original — just in English.

Working dir: `/tmp/claude-0/-home-user-TEST01/d7e2d964-2f3d-5210-bb5f-d09701a545e0/scratchpad/work`

## Inputs for original slide NN
- `src/figNN_0.jpg` — the original figure, full resolution. **This is your source of truth for layout.**
- `src/slideNN.md` — the slide's text (title, lead statement, key points, summary table, limits, source).
- `src/r-NN.jpg` — a render of the whole slide, for context.
- `out_AC.json`, `out_B.json`, `out_D.json`, `out_E.json` — an earlier English rewrite of the same slides.
  Mine these for vocabulary and for the English phrasing of terms, but do NOT copy their structure.

## Part 1 — the figure

Write `fig/figNN.html` and render it:

```
cd fig && node render.js figNN.html
```

That writes `fig/figNN.png` at 1600 px wide, 2x device scale, and prints the pixel size.
**Then open the PNG and compare it against the original side by side.** Iterate until it reads as the
same diagram. Budget two or three iterations — first renders usually overflow or wrap badly.

Rules:
1. `<link rel="stylesheet" href="style.css">` — the shared house style. Read `style.css` first; it already
   provides `.ftitle .fsub .seclabel .fsrc .box .box.hl .box.dash .bh .bt .bs .num .row .cols .grid
   .arrow .chip .chip.ok .chip.bad .chip.na table.ftab .legend .sw`. Use these classes; add inline styles
   or a small `<style>` block only for geometry this set cannot express. Do not change `style.css`
   (other agents share it) and do not invent a new palette.
2. **Match the original's form**: a flow stays a flow, a matrix stays a matrix, a quadrant chart stays a
   quadrant chart, a card wall stays a card wall, a funnel stays a funnel. Keep the same number of boxes,
   the same grouping, the same arrows, and the same emphasis (the original highlights the critical item in
   brown — keep that).
3. **Target aspect ratio** is given in your task list as width:height. The rendered PNG must come out close
   to it (within about 10%), because the slide reserves a box of that shape. Control height with the amount
   of padding and font size, not by stretching. The printed size from `render.js` tells you what you got;
   if the ratio is off, adjust and re-render.
4. **Editorial rules** (the owner insisted): no internal codes anywhere in the figure — no M01–M36,
   no G1–G6, no E0–E4, no 图/表 numbers, no `[A] x.y` source indices, no F01–F26. Name the gates
   descriptively: rights & backbone-chemistry gate, receptor-entry & payload-fit gate, patient-cell biology
   gate, platform-transfer gate, tissue translation & safety gate, CMC & cost gate. The experiments are the
   freedom-to-operate & backbone study, the productive-delivery study (with patient-cell arm), the
   second-payload (DYSF) study, the in vivo target-cell PD & safety study, and the conjugate
   characterization study. Never write "this page shows…". Undecided amounts are "TBD".
5. Keep every number, gene name, variant, NCT identifier, company name and date exactly as in the original.
   If a label in the original is unreadable, leave it out rather than guess, and say so in your report.
6. Text in figures must be legible when the image is about 6 inches wide on a slide: nothing below 14 px
   at the 1600 px authoring width.

## Part 2 — the slide content

The English deck reuses the original deck's page architecture, so give me the same slots it has.
Append one object per slide to **your own** output file (named in your task list), as a JSON array:

```json
{
  "orig": 14,
  "title": "Conclusion headline, no trailing period, <= 120 chars",
  "lead": "One sentence: the claim this page makes. 180-300 chars. No label, no 'this page'.",
  "figure": "fig/fig14.png",
  "fig_caption": "One line describing what the figure shows and how to read it, <= 150 chars",
  "points": ["Short phrase", "Short phrase", "Short phrase"],
  "table_note": ["Short bullet summarising a table that stays in the memo", "..."],
  "limits": ["A caveat that qualifies the conclusion", "..."],
  "source": "Source: IC Memo v2 (§4.1, §4.1.2); data as of 2026-10-13",
  "notes": "2-5 sentences of talk track, plus anything cut for space."
}
```

Slot budgets — these are hard, the builder does not shrink below 9 pt:
- `points`: 3–4 items, each **≤ 46 characters**. They are labels, not sentences.
- `table_note`: 0–3 items, each ≤ 110 characters.
- `limits`: 1–3 items, each ≤ 190 characters.
- `lead` ≤ 300 characters; `title` ≤ 120; `fig_caption` ≤ 150.

Convert `[A] x.y` to `IC Memo v2 §x.y` in `source`. Keep NCT numbers and dates.

## Finishing
- Validate: `python3 -c "import json;json.load(open('<your file>'))"`.
- Confirm every `fig/figNN.png` exists and its aspect is close to target.
- Final reply: 4–8 lines. Which figures you built, any whose aspect you could not hit, anything in the
  original you could not read. Do not paste JSON, HTML or Chinese source text into your reply.

---## REVISION 2 (supersedes everything above where they conflict)

My earlier "14 px at a 1600 px canvas" floor was wrong and it is why the figures came out unreadable.
A 1600 px canvas shown in the slide's 8.45 in figure box renders 1 authored px as **0.38 pt**, so a 14 px
label lands at **5.3 pt** on the slide. Nobody can read that. The numbers below replace it.

### Canvas and type (hard)
- Override the shared width in each figure's own `<style>` block: **`body { width: 1150px; }`**.
  Do not use 1600 px and do not go wider - a wider canvas makes the type *smaller* on the slide, not larger.
- Height about **575 px**, i.e. **aspect ~2.0**; acceptable 1.85 to 2.60 (height 442-622 px).
- **Minimum font size 16 px**, which is 8.4 pt on the slide. Suggested scale:
  figure title 26 px, section label 19 px, box heading 18 px, body and cell text 16-17 px,
  legend 16 px. The `.fsrc` source line may stay at 14 px - it is the only exception.
- Nothing else below 16 px: no 15 px, no 13.5 px, no 12 px badge digits.

### What this costs, and what to do about it
The canvas drops ~28% while the floor rises, so a dense figure cannot carry the same wordcount.
**Roughly half the words have to come out of the dense ones.** Decide what goes by this order:

Keep, always: the structure itself - every box, cell, row, column, lane and arrow; the grouping; the
colour coding and the brown emphasis; every number, gene name, variant, date, company name and identifier.

Compress: full sentences become noun phrases. "The oligonucleotide blocks a pseudoexon that a deep-intronic
variant causes to be included" becomes "Blocks pseudoexon from deep-intronic variant". Cut restatements,
parentheticals and hedges inside cells.

Move out: explanatory sentences belong in the slide text, not in the figure. Put them in `points` - which
is now 3 to 5 bullets of 60-130 characters explaining what the figure shows and what follows from it - or
in `limits`. That is where the detail goes; it is not lost.

**Do not reduce the number of cells, boxes or rows.** The structure is the content. If a figure still will
not fit, tell me which rows you would drop and why, and leave it at its best fit rather than cutting.

### Check before you report
Render, then confirm: (1) the printed size divides out to 1.85-2.60; (2) `grep -oE '[0-9.]+px' figNN.html`
shows no font-size below 16 px; (3) open the PNG and read it at the size it will appear - if you would
squint at it on a projector, it is still too dense.

Report per figure: final px size, aspect, and the smallest font size used.
