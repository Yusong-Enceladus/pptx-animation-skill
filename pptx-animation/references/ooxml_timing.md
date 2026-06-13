# PowerPoint canonical `<p:timing>` structure & gotchas

Read this when hand-writing/injecting animation XML (Method A, `inject_animations.py`) or debugging a "repaired" dialog. `inject_animations.py` already reproduces this exactly; this file explains *why* it must.

## The two details a "repaired" prompt is caused by

PowerPoint silently rejects timing XML whose shape differs from what it emits — it shows a *"PowerPoint found a problem and repaired it"* dialog and drops the animations. The two non-obvious requirements (verified by diffing PowerPoint's own repaired output against a naive version):

1. The `clickEffect`/`withEffect` `<p:par>` is a **direct child** of the click-wrapper's `<p:childTnLst>`. Do **not** insert an extra `<p:par>…<p:cTn fill="hold"><p:cond delay="0"/>` wrapper between the click-wrapper and the effect par.
2. The visibility `<p:set>`'s `<p:cTn>` is **self-closed**: `<p:cTn id=".." dur="1" fill="hold"/>` — with **no** `<p:stCondLst>` child.

## Skeleton

`<p:timing>` goes immediately before `<p:extLst>` in `slideN.xml`. IDs are a single sequential counter across the whole tree.

```xml
<p:timing><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
    <p:seq concurrent="1" nextAc="seek">
      <p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>

        <!-- ONE click group: a wrapper that waits for a page click, then N effect pars -->
        <p:par><p:cTn id="3" fill="hold">
          <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
          <p:childTnLst>
            <!-- shape 1 of this click: clickEffect -->
            <p:par><p:cTn id="4" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
              <p:stCondLst><p:cond delay="0"/></p:stCondLst>
              <p:childTnLst>
                <p:set><p:cBhvr>
                  <p:cTn id="5" dur="1" fill="hold"/>
                  <p:tgtEl><p:spTgt spid="7"/></p:tgtEl>
                  <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                </p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
                <p:animEffect transition="in" filter="fade"><p:cBhvr>
                  <p:cTn id="6" dur="450"/>
                  <p:tgtEl><p:spTgt spid="7"/></p:tgtEl>
                </p:cBhvr></p:animEffect>
              </p:childTnLst>
            </p:cTn></p:par>
            <!-- shape 2+ in the SAME click: identical but nodeType="withEffect" -->
          </p:childTnLst>
        </p:cTn></p:par>
        <!-- more click groups… -->

      </p:childTnLst></p:cTn>
      <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
      <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
    </p:seq>
  </p:childTnLst></p:cTn>
</p:par></p:tnLst></p:timing>
```

- `spid` = the shape's `<p:cNvPr id="..">`, identical to `python-pptx` `shape.shape_id`.
- `filter="fade"`, `presetID="10"` = the basic fade entrance. To get the canonical XML for a *different* effect, let PowerPoint author one via Method B, unzip the deck, and copy that `<p:timing>` verbatim as a new template.

## Other gotchas

- **`python-pptx` renumbers slide parts on save.** After a `python-pptx` round-trip, `slideN.xml` is re-indexed to presentation order. Always key skip/target by **presentation-order index**, never file number. Both scripts here do.
- **AppleScript (Method B) can crash/hang the user's PowerPoint** on large batches — repeated open/save/close, `AppleEvent timed out (-1712)`, `-9074`. Wrapping each slide's script in `with timeout of 600 seconds` helps, but Method A (no app) is the robust answer for big decks.
- **Static fallback is automatic.** Entrance animations leave the final slide fully visible; a viewer whose app ignores animations still sees everything. LibreOffice `--convert-to pdf` renders that final state — good for validity/QA, but it does NOT play the animation, so it can't verify reveal *timing*.
- **Verify decomposition, not playback.** When you can't open PowerPoint, overlay `cluster_blocks.py --map`'s `blockmap.json` boxes (numbered in reveal order) on 96-dpi slide renders and check by eye: each click = one whole card/column/row, lead first, banner last, columns left→right.
