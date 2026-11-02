#!/usr/bin/env python3
"""Merge figure-based specs, converted no-figure specs and the new pages into deck_spec2.json."""
import json, os, re, sys

W = os.path.dirname(os.path.abspath(__file__))
def load(n):
    p = os.path.join(W, n)
    return json.load(open(p)) if os.path.exists(p) else []

FIG = {}
for f in ("out2_A.json", "out2_B.json", "out2_C.json", "out2_D.json",
          "out2_E.json", "out2_F.json", "out2_MAIN.json"):
    for s in load(f):
        k = s.get("key") or s.get("orig")
        if k in (0, None):
            continue
        FIG[k] = s

V1 = {}
for f in ("out_AC.json", "out_B.json", "out_D.json", "out_E.json"):
    for s in load(f):
        if s.get("orig"):
            V1[s["orig"][0]] = s

def convert(o):
    """No-figure slide: reuse the v1 English content inside the original page frame."""
    s = V1.get(o)
    if not s:
        print(f"  !! no source for {o}")
        return None
    regs = [r for r in s.get("regions", [])]
    lead = ""
    for r in list(regs):
        if r["type"] == "callout":
            lead = r["text"]; regs.remove(r); break
    if not lead:
        for r in regs:
            if r["type"] == "bullets" and r.get("items"):
                it = r["items"][0]
                lead = it["text"] if isinstance(it, dict) else it
                r["items"] = r["items"][1:]
                if not r["items"]:
                    regs.remove(r)
                break
    return {"orig": o, "title": s["title"], "lead": lead, "layout": s.get("layout", "full"),
            "regions": regs, "source": s.get("source"), "notes": s.get("notes"), "limits": []}

def get(o):
    return FIG.get(o) or convert(o)

SEC = {"exec": "Executive summary", "why_now": "Why now", "eval": "How we evaluate",
       "why_us": "Why us", "build": "How we build", "rec": "Recommendation"}

deck, missing = [], []
def add(items, section):
    out = []
    for it in items:
        s = FIG.get(it) if isinstance(it, str) else get(it)
        if s is None:
            missing.append(it); continue
        s = dict(s); s["section"] = section
        out.append(s)
    return out

deck += add(["exec", "scorecard", "contents"], SEC["exec"])

sec = add(["aoc_macro", "aoc_comp", 17, 18, 20, 21], SEC["why_now"])
deck.append({"kind": "divider", "section": SEC["why_now"], "title": "Why now",
             "subtitle": "What conjugates made possible, who already occupies the field, "
                         "and which part of the problem is still open"})
deck += sec

sec = add([13, 14, 15, 23, 24, 26, 27, 28, 29, 30], SEC["eval"])
deck.append({"kind": "divider", "section": SEC["eval"], "title": "How we evaluate",
             "subtitle": "Taking a conjugate apart into binder, payload and chemistry, "
                         "and the evidence a disease candidate must produce"})
deck += sec

sec = add([32, 33, 35, 36, "wuxi_oligo", "wuxi_aoc"], SEC["why_us"])
deck.append({"kind": "divider", "section": SEC["why_us"], "title": "Why us",
             "subtitle": "What the group can already execute end to end, and what a new "
                         "company must own itself"})
deck += sec

sec = add([50, 38, 39, 40, 41, 42, 43, 44, 46, 47, 48, 51, 52, 53, 54, 55, 56], SEC["build"])
deck.append({"kind": "divider", "section": SEC["build"], "title": "How we build",
             "subtitle": "Entry route, the three configurations, first programmes, team, "
                         "budget and the governance we ask the IC to approve"})
deck += sec

sec = add(["recommendation"], SEC["rec"])
deck.append({"kind": "divider", "section": SEC["rec"], "title": "Recommendation",
             "subtitle": "What we ask the investment committee to approve today"})
deck += sec

BAD = re.compile(r"\b[MGE]\d{1,2}\b|\[A\]\s*\d|本页|本轮|[一-鿿]")
issues = []
nofig = []
for i, s in enumerate(deck, 2):
    if s.get("kind"):
        continue
    for m in set(BAD.findall(json.dumps(s, ensure_ascii=False))):
        issues.append(f"slide {i}: {m!r} in {s.get('title','')[:46]}")
    if not s.get("figure"):
        nofig.append(s.get("orig"))
    elif not os.path.exists(os.path.join(W, s["figure"])):
        issues.append(f"slide {i}: figure not rendered yet: {s['figure']}")

json.dump(deck, open(os.path.join(W, "deck_spec2.json"), "w"), ensure_ascii=False, indent=1)
print(f"deck_spec2.json: {len(deck)} slides (+cover = {len(deck)+1})")
print("missing specs:", missing)
print("pages without a figure:", nofig)
for x in issues[:30]:
    print("  CHECK", x)
