#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inject click-by-click FADE entrance animations into a .pptx by writing PowerPoint's
CANONICAL <p:timing> XML directly into each slide part. Pure file/zip manipulation — it never
opens PowerPoint, so it cannot crash or hang the app and works headless / on large decks.

Why this exists: driving PowerPoint's AppleScript API per-slide is unreliable on large decks
(repeated open/save/close hangs and crashes, AppleEvent timeouts). And NAIVELY hand-written
timing XML is silently rejected by PowerPoint (it shows a "repaired" dialog and drops the
animations). The structure below is byte-for-byte the structure PowerPoint itself emits, so it
is accepted without a repair prompt. The two details that MUST be exact (a wrong nesting here is
exactly what triggers the repair):
  * the clickEffect <p:par> is a DIRECT child of the click-wrapper's <p:childTnLst>
    (do NOT add an extra <p:par delay="0"> wrapper between them);
  * the <p:set>'s <p:cTn> is self-closed `<p:cTn id=.. dur="1" fill="hold"/>` with NO stCondLst.

Spec format (JSON):  targets are shape ids (the <p:cNvPr id="..">, == python-pptx shape.shape_id)
  {
    "duration": 0.45,                       # seconds, optional (default 0.45)
    "slides": [
      {"slide": 8, "groups": [[7, 9, 12], [20, 21]]},   # slide = PRESENTATION-ORDER index (1-based)
      ...
    ]
  }
Each inner list is ONE click: its shapes fade in together (first = clickEffect, rest = withEffect).

Usage:
  python inject_animations.py in.pptx out.pptx spec.json
"""
import sys, json, zipfile, shutil, argparse, re
from lxml import etree

P="http://schemas.openxmlformats.org/presentationml/2006/main"
A="http://schemas.openxmlformats.org/drawingml/2006/main"
R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
def q(t): return f"{{{P}}}{t}"

def build_timing(groups, dur_ms=450):
    """Return a <p:timing> element matching PowerPoint's own canonical fade-entrance structure."""
    nid=[0]
    def nx(): nid[0]+=1; return str(nid[0])
    timing=etree.Element(q('timing'), nsmap={'p':P,'a':A,'r':R})
    tn=etree.SubElement(timing,q('tnLst'))
    ctn1=etree.SubElement(etree.SubElement(tn,q('par')),q('cTn'),id=nx(),dur="indefinite",restart="never",nodeType="tmRoot")
    seq=etree.SubElement(etree.SubElement(ctn1,q('childTnLst')),q('seq'),concurrent="1",nextAc="seek")
    ctn2=etree.SubElement(seq,q('cTn'),id=nx(),dur="indefinite",nodeType="mainSeq")
    mseq=etree.SubElement(ctn2,q('childTnLst'))
    for group in groups:
        # click-wrapper: waits for a page click (delay="indefinite")
        gctn=etree.SubElement(etree.SubElement(mseq,q('par')),q('cTn'),id=nx(),fill="hold")
        etree.SubElement(etree.SubElement(gctn,q('stCondLst')),q('cond'),delay="indefinite")
        gch=etree.SubElement(gctn,q('childTnLst'))
        for j,spid in enumerate(group):
            # effect par is a DIRECT child of gch (no extra wrapper — that is the repair-trigger)
            ectn=etree.SubElement(etree.SubElement(gch,q('par')),q('cTn'),id=nx(),presetID="10",
                                  presetClass="entr",presetSubtype="0",fill="hold",
                                  nodeType=("clickEffect" if j==0 else "withEffect"))
            etree.SubElement(etree.SubElement(ectn,q('stCondLst')),q('cond'),delay="0")
            ech=etree.SubElement(ectn,q('childTnLst'))
            # set style.visibility -> visible (its cTn is self-closed, NO stCondLst)
            setel=etree.SubElement(ech,q('set')); cb=etree.SubElement(setel,q('cBhvr'))
            etree.SubElement(cb,q('cTn'),id=nx(),dur="1",fill="hold")
            etree.SubElement(etree.SubElement(cb,q('tgtEl')),q('spTgt'),spid=str(spid))
            etree.SubElement(etree.SubElement(cb,q('attrNameLst')),q('attrName')).text="style.visibility"
            etree.SubElement(etree.SubElement(setel,q('to')),q('strVal'),val="visible")
            # fade entrance
            ae=etree.SubElement(ech,q('animEffect'),transition="in",filter="fade")
            cb2=etree.SubElement(ae,q('cBhvr'))
            etree.SubElement(cb2,q('cTn'),id=nx(),dur=str(dur_ms))
            etree.SubElement(etree.SubElement(cb2,q('tgtEl')),q('spTgt'),spid=str(spid))
    # seq prev/next click conditions (standard)
    pc=etree.SubElement(etree.SubElement(seq,q('prevCondLst')),q('cond'),evt="onPrev",delay="0")
    etree.SubElement(etree.SubElement(pc,q('tgtEl')),q('sldTgt'))
    nc=etree.SubElement(etree.SubElement(seq,q('nextCondLst')),q('cond'),evt="onNext",delay="0")
    etree.SubElement(etree.SubElement(nc,q('tgtEl')),q('sldTgt'))
    return timing

def insert_timing(slide_bytes, timing):
    root=etree.fromstring(slide_bytes)
    old=root.find(q('timing'))
    if old is not None: root.remove(old)
    ext=root.find(q('extLst'))
    if ext is not None: ext.addprevious(timing)      # <p:timing> must precede <p:extLst>
    else: root.append(timing)
    return etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)

def presentation_order_to_partname(pptx):
    """Map 1-based presentation index -> ppt/slides/slideN.xml (sldIdLst order)."""
    z=zipfile.ZipFile(pptx)
    pres=etree.fromstring(z.read('ppt/presentation.xml'))
    rels=etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    r2t={r.get('Id'):r.get('Target') for r in rels}
    out={}
    for i,sld in enumerate(pres.find(q('sldIdLst')),start=1):
        tgt=r2t[sld.get(f'{{{R}}}id')]
        out[i]='ppt/'+tgt if not tgt.startswith('/') else tgt.lstrip('/')
    z.close(); return out

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('input'); ap.add_argument('output'); ap.add_argument('spec')
    ap.add_argument('--duration', type=float, default=None, help='fade seconds (overrides spec)')
    a=ap.parse_args()
    spec=json.load(open(a.spec))
    dur_ms=int((a.duration if a.duration is not None else spec.get('duration',0.45))*1000)
    idx2part=presentation_order_to_partname(a.input)
    file_specs={}
    for e in spec['slides']:
        groups=[[int(s) for s in g] for g in e['groups'] if g]
        if groups: file_specs[idx2part[int(e['slide'])]]=groups
    shutil.copyfile(a.input, a.output)
    tmp=a.output+'.tmp'
    with zipfile.ZipFile(a.output,'r') as zin, zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data=zin.read(it.filename)
            if it.filename in file_specs:
                data=insert_timing(data, build_timing(file_specs[it.filename], dur_ms))
            zout.writestr(it, data)
    import os; os.replace(tmp, a.output)
    clicks=sum(len(g) for g in file_specs.values()); fx=sum(len(x) for g in file_specs.values() for x in g)
    print(f"injected {len(file_specs)} slides | {clicks} clicks | {fx} shape-effects -> {a.output}")

if __name__=='__main__': main()
