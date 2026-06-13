---
name: pptx-animation
description: Use when a PowerPoint .pptx deck needs click-by-click entrance animations, grouped column/card/flow reveals, semantic block animation, animation QA, or animation XML/API troubleshooting. Adds entrance animations to existing slide elements while preserving the static layout.
---

# PPTX Animation

Add click-by-click entrance (fade) animations to an existing `.pptx`, preserving the static slide. Run scripts from the skill root (`python3 scripts/…`).

## Two methods — pick by environment

- **Method A — pure-file injection (default, recommended).** Writes PowerPoint's canonical `<p:timing>` XML straight into the deck via zip/lxml. Never opens PowerPoint, so it can't crash/hang the app, runs headless, scales to large decks.
- **Method B — PowerPoint AppleScript API** (macOS, PowerPoint open). Only for small decks where you want PowerPoint to author the XML.

Driving the AppleScript API per-slide is unreliable on large decks (hangs, AppleEvent timeouts, can crash the user's PowerPoint). And naively hand-written timing XML is rejected by PowerPoint (a "repaired" dialog drops the animations). Method A clones the exact structure PowerPoint emits, so it is accepted. Prefer Method A unless you specifically need PowerPoint in the loop.

## Animate by SEMANTIC BLOCKS, not by element

One click reveals one whole visual unit — a complete card (box+title+body+image+chip+arrow together), a whole column, a row, a table, a panel-with-its-rows. Per-element reveals look broken. Order by reading logic: top lead/headline first → body blocks in natural flow (columns left→right, rows top→bottom, pipeline in process order, arrows with the block they point into) → result/banner punchline LAST. Cover, dividers, thank-you, and verbatim/embedded slides get NO animation.

## Method A workflow

```bash
# 1. auto-cluster slides into semantic blocks -> spec.json; skip non-animated slides by 1-based index
python3 scripts/cluster_blocks.py deck.pptx --out spec.json --skip 1,2,3,54 --map /tmp/bm
# 2. VERIFY the block-map (below), hand-edit spec.json groups if needed
# 3. inject canonical animation XML (no PowerPoint involved)
python3 scripts/inject_animations.py deck.pptx deck_animated.pptx spec.json --duration 0.45
# 4. render to PDF (soffice --headless --convert-to pdf; pdftoppm) to confirm file valid + static state intact
```

Spec: `slide` is the 1-based presentation index; each inner list is ONE click (members fade in together, first=`clickEffect`, rest=`withEffect`); members are shape ids for Method A.
```json
{"duration":0.45,"slides":[{"slide":8,"groups":[[7,9,12],[20,21]]}]}
```

**Verify without opening PowerPoint:** `cluster_blocks.py --map <dir>` writes `blockmap.json` (per-slide block boxes in reveal order). Render the deck to JPGs at 96 dpi and overlay each box with its reveal number — confirm grouping (whole cards/columns) and order (lead first, banner last, columns left→right) by eye. Keep tuning `--skip` / `spec.json` until correct.

## Method B workflow (AppleScript, small interactive decks)

```bash
python3 scripts/list_shapes.py deck.pptx --slides 5 6 --output shapes.csv   # prefer UNIQUE shape names
python3 scripts/add_click_animations.py deck.pptx out.pptx spec.json --per-slide
python3 scripts/count_effects.py out.pptx --slides 5 6                       # confirm PowerPoint recognizes effects
```
Group members are shape names (safer than indices — AppleScript shape order ≠ python-pptx order). `--per-slide` opens/saves/closes per slide to avoid stale timeline objects (slow, closes open decks). If it hangs/crashes across many slides, switch to Method A.

## Always
Work on a copy; keep the un-animated deck as the primary deliverable and ship the animated one alongside.

## Canonical XML, structure details, and gotchas
See `references/ooxml_timing.md` — the exact `<p:timing>` skeleton, the two nesting details that cause a "repaired" dialog if wrong, the `python-pptx` slide-renumbering trap, and the validity-vs-playback QA notes.
