# PPTX Animation Skill

An agent skill for adding click-by-click PowerPoint entrance animations to existing `.pptx` decks while preserving the static slide layout. Agent-agnostic — works with any skill-capable coding agent (Claude Code, Codex, etc.).

## What It Does

- Animates by **semantic blocks** — one click reveals a whole card/column/row/table, in reading order — not element-by-element.
- **Method A (recommended): pure-file OOXML injection.** Writes PowerPoint's canonical `<p:timing>` XML directly into the deck via zip/lxml. Never opens PowerPoint, so it can't crash or hang the app, runs headless, and scales to large decks.
- **Method B: PowerPoint AppleScript API** (macOS) for small interactive decks where you want PowerPoint to author the XML.
- Helper scripts to auto-cluster blocks, list shapes, and count recognized effects.

## Why Method A exists

Driving PowerPoint's AppleScript API per-slide is unreliable on large decks (repeated open/save/close hangs, AppleEvent timeouts, and crashes of the user's PowerPoint). And naively hand-written timing XML is silently rejected by PowerPoint (it shows a "repaired" dialog and drops the animations). Method A reproduces the *exact* structure PowerPoint itself emits, so it is accepted without a repair prompt — see `pptx-animation/SKILL.md` for the canonical structure and the two nesting details that matter.

## Requirements

- Python 3, with `python-pptx` and `lxml` (`pip install -r requirements.txt`).
- Method A: nothing else (pure file manipulation). Rendering for QA uses LibreOffice (`soffice`) + `pdftoppm` if available.
- Method B only: macOS with Microsoft PowerPoint installed.

## Install The Skill

Copy the `pptx-animation/` folder into your agent's skills directory, e.g.:

```bash
cp -R pptx-animation ~/.codex/skills/      # Codex
cp -R pptx-animation ~/.claude/skills/     # Claude Code
```

Restart the agent if it does not hot-reload skills.

## Quick Start (Method A)

```bash
# 1. auto-cluster slides into semantic blocks -> spec.json (skip non-animated slides by index)
python pptx-animation/scripts/cluster_blocks.py deck.pptx --out spec.json --skip 1,2,3,54 --map /tmp/bm

# 2. (verify the block-map overlay, then) inject canonical animation XML — no PowerPoint involved
python pptx-animation/scripts/inject_animations.py deck.pptx deck_animated.pptx spec.json --duration 0.45
```

## Quick Start (Method B, macOS + PowerPoint)

```bash
python pptx-animation/scripts/list_shapes.py deck.pptx --slides 5 6 --output shapes.csv
python pptx-animation/scripts/add_click_animations.py deck.pptx out.pptx spec.json --per-slide
python pptx-animation/scripts/count_effects.py out.pptx --slides 5 6
```

## Scripts

| Script | Method | Purpose |
|---|---|---|
| `cluster_blocks.py` | A | Auto-cluster each slide into semantic blocks; emit the animation spec + a reviewable block-map. |
| `inject_animations.py` | A | Inject PowerPoint-canonical `<p:timing>` directly into the deck (no PowerPoint app). |
| `add_click_animations.py` | B | Add fade effects via the PowerPoint AppleScript API. |
| `count_effects.py` | B | Count PowerPoint-recognized effects per slide (validation). |
| `list_shapes.py` | both | Inventory shapes (id, name, position, text) to build/inspect specs. |

See `pptx-animation/SKILL.md` for the full workflow, the canonical timing structure, the verify-without-PowerPoint loop, and gotchas.

## License

MIT
