# English IC deck — slide spec format and editorial rules

## Context
We are turning a Chinese investment-committee (IC) deck about setting up an AOC (antibody–oligonucleotide conjugate) NewCo into a clean **English** IC deck. The owner reviewed the Chinese deck and left comments; the rules below encode them. The output is JSON slide specs that a Python builder renders onto the existing slide master (16:9, 13.33" x 7.5", WuXi AppTec branded, Arial).

Source material per original slide N (N = page number in the annotated Chinese deck):
- `src/slideNN.md` — all text on the slide (titles, bullets, tables, source line). Some table cells were already truncated with "…" in the original; infer the meaning from the figure, or drop the clause rather than inventing facts.
- `src/figNN_0.jpg` — the slide's main figure at full resolution (Chinese). Read it: its content usually carries the real substance and must be rebuilt as native text/table/flow regions (the English deck will NOT reuse the Chinese images).
- `src/r-NN.jpg` — a render of the whole slide for overall context.

## Editorial rules (from the owner's comments — mandatory)
1. **No internal codes anywhere.** Remove unit numbers (M01, M02 …), gate codes (G1–G6), experiment codes (E0–E4), figure/table numbers (图 M03-1, 表 M11-1), and source-index codes like "[A] 4.1.7", "[A] 1 第 9 条", "F01". Replace a gate/experiment code with a short descriptive name, e.g. G1 → "Rights & backbone-chemistry gate", G2 → "Receptor-entry & payload-fit gate", G3 → "Patient-cell biology gate", G4 → "Platform-transfer gate", G5 → "Tissue translation & safety gate", G6 → "CMC & cost gate"; E0 → "freedom-to-operate & backbone study", E1 → "productive-delivery study (with patient-cell arm)", E2 → "second-payload (DYSF) study", E3 → "in vivo target-cell PD & safety study", E4 → "conjugate characterization study". Use these exact names consistently.
2. **No AI/meta narration.** Drop phrases such as 本页主张 / 本页要点 / 本页表（全表见备忘录）/ 本轮请求 / 本轮请批 / 本轮授权 / 本轮归属 / 关口挂接 / 拨付挂读出、不挂商务事件与日历 / 从左向右 / "本页待定 N 条" / "新绘" / "研究推断" labels, and anything that describes the document itself rather than the business. Do not write "this page shows…".
3. Avoid the jargon words 现状 / 链条 / 档位 / 边界 as labels (i.e., don't write "status quo", "chain", "tier/grade band", "boundary" as framing labels). Say plainly what is known, what must be proven, and what limits the conclusion.
4. "最可能先断的证据链" → "evidence still to be completed" (never "most likely to break").
5. Amounts/valuation/ownership that are undecided: write "TBD" (optionally "TBD – Finance"), once per table column if possible, not on every cell.
6. Tone: investment-committee English. Each title is a **conclusion headline** (one sentence, ≤ 120 characters, no trailing period). Concise, concrete, no hype. Keep numbers, trial names, NCT numbers, dates, and gene/target names exactly as in the source.
7. Keep the substance: every slide must preserve the original's key facts and decisions; compress wording, don't drop the argument. You may merge two original slides into one only if they are clearly redundant; say so in `notes`.
8. Sources: one short line, e.g. "Source: IC Memo v2 (§4.1, §7.6); HARBOR Phase 3 NCT06411288 topline (2026-09-08); team analysis". Convert "[A] x.y" to "IC Memo v2 §x.y".
9. Speaker notes (`notes`): 2–5 sentences of talk track, plus anything you had to cut from the slide for space.

## Space budget (the builder does NOT shrink text below 9 pt — respect these limits)
- Title ≤ 120 chars. Region heading ≤ 60 chars.
- `bullets`: ≤ 6 items, each ≤ 170 chars; optional `sub` ≤ 2 items each ≤ 120 chars.
- `table`: ≤ 6 columns, ≤ 7 body rows; each cell ≤ 90 chars (header cell ≤ 30 chars). Prefer 3–5 columns.
- `cards`: 2–4 cards; heading ≤ 40 chars, body ≤ 220 chars.
- `flow`: 3–6 steps; heading ≤ 30 chars, body ≤ 140 chars.
- `kpis`: 2–4 items; value ≤ 10 chars, label ≤ 60 chars.
- A slide has at most 3 regions in total.

## JSON format
Write a JSON array of slide objects to the output file you are given.

```json
{
  "orig": [13],                       // original slide number(s) this replaces ([] for new)
  "section": "How we evaluate",       // section tag shown top-right (use the one you are given)
  "title": "Conclusion headline",
  "layout": "full | split | split_left_wide | split_right_wide | stack",
  "regions": [ REGION, ... ],         // full: 1–2 stacked; split*: [left, right] (+ optional 3rd = bottom full-width strip); stack: top then bottom
  "source": "Source: …",
  "notes": "Talk track …"
}
```
REGION types:
```json
{"type":"bullets","heading":"optional","items":[{"text":"…","sub":["…"]}, "plain string also ok"]}
{"type":"table","heading":"optional","columns":["…"],"rows":[["…"]],"col_widths":[0.2,0.3,0.5], "first_col_bold":true}
{"type":"cards","heading":"optional","items":[{"heading":"…","body":"…"}]}
{"type":"flow","heading":"optional","steps":[{"heading":"…","body":"…"}]}
{"type":"kpis","heading":"optional","items":[{"value":"45","label":"…"}]}
{"type":"callout","text":"one key takeaway sentence ≤ 200 chars"}
```
Layout guidance: `split` = 50/50; `split_left_wide` = 60/40; `split_right_wide` = 40/60; `stack` = top region then bottom region full width. A `callout` works best as the last region (renders as a tinted strip). Vary layouts across slides — don't make every slide a table.

## Section-divider slides
`{"orig":[], "kind":"divider", "section":"…", "title":"Section title", "subtitle":"one line on what this section answers"}`
