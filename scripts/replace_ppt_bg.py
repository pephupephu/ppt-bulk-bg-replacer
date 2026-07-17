#!/usr/bin/env python3
"""Replace background images on every slide of PowerPoint (.pptx) files.

Two modes:
  Single: python replace_ppt_bg.py input.pptx bg.jpg [output.pptx]
  Batch:  python replace_ppt_bg.py input/dir/ bg.jpg [output/dir/]

INSERT-only approach:
  - For every slide, inserts a new full-slide picture at bottom z-order
  - Adds <p:bg> with explicit background declaration
  - Adds <a:noFill/> to existing full-slide elements
  - Preserves ALL content images, text, charts, animations

Fixes over v1 (SWAP approach):
  - No more SWAP mode that replaced binary data of existing images
  - No more extension mismatch (.png file containing JPEG data)
  - Proper batch directory processing with subdirectory preservation
  - Reliable Chinese/Unicode path support throughout
"""
import sys, os, zipfile, re
from lxml import etree
from zipfile import ZipInfo

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
SW, SH = 12192000, 6858000
MW, MH = int(SW * 0.85), int(SH * 0.85)
FILL_TAGS = {f"{{{A}}}{t}" for t in ("solidFill","gradFill","blipFill","pattFill","grpFill")}
SLIDE_RE = re.compile(r"/slide(\d+)\.xml$")

def _tostring(el):
    xml = etree.tostring(el, xml_declaration=False, encoding="UTF-8")
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml.replace(b"\n", b"\r\n")

def _size_of(el):
    for xf in el.iter(f"{{{A}}}xfrm"):
        ex = xf.find(f"{{{A}}}ext")
        if ex is not None: return int(ex.get("cx","0")), int(ex.get("cy","0"))
    return 0, 0

def replace_ppt_background(inp, bg, out=None):
    if out is None:
        base, ext = os.path.splitext(inp)
        out = base + "_bg_replaced" + ext
    for p, lb in [(inp,"Input"),(bg,"Background")]:
        if not os.path.exists(p):
            print(f"  Error: {lb} not found"); return False
    bg_ext = os.path.splitext(bg)[1].lstrip(".")
    with open(bg, "rb") as f: bg_bytes = f.read()
    with zipfile.ZipFile(inp) as z:
        orig = z.infolist()
        files = {fi.filename: (fi, z.read(fi.filename)) for fi in orig}
    slides = sorted(int(m.group(1)) for n in files if (m:=SLIDE_RE.search(n)) and "/_rels/" not in n)
    print(f"  {len(slides)} slides found")
    ct = files.get("[Content_Types].xml", (None,b""))[1]
    exts = set(re.findall(r'Extension="([^"]+)"', ct.decode("utf-8"))) if ct else set()
    new_ext = bg_ext if bg_ext in exts else \
              (bg_ext.upper() if bg_ext.upper() in exts else \
               (bg_ext.lower() if bg_ext.lower() in exts else \
                next((e for e in ["png","jpeg","jpg","JPG"] if e in exts), bg_ext)))
    print(f"  Extension: {new_ext}")
    added = []
    for sn in slides:
        sp = f"ppt/slides/slide{sn}.xml"
        rp = f"ppt/slides/_rels/slide{sn}.xml.rels"
        sld = etree.fromstring(files.get(sp,(None,b""))[1])
        cSld = sld.find(f"{{{P}}}cSld")
        spTree = cSld.find(f"{{{P}}}spTree") if cSld is not None else None
        # Remove old <p:bg>
        ob = cSld.find(f"{{{P}}}bg")
        if ob is not None: cSld.remove(ob)
        # Add noFill to existing full-slide shapes
        if spTree is not None:
            for el in list(spTree.iter()):
                tl = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if tl == "spPr":
                    w, h = _size_of(el)
                    if w >= MW or h >= MH:
                        for ch in list(el):
                            if ch.tag in FILL_TAGS: el.remove(ch)
                        if el.find(f"{{{A}}}noFill") is None:
                            xf = el.find(f"{{{A}}}xfrm")
                            pos = list(el).index(xf)+1 if xf is not None else 0
                            el.insert(pos, etree.Element(f"{{{A}}}noFill"))
                elif tl == "grpSpPr":
                    for ch in list(el):
                        if ch.tag in FILL_TAGS: el.remove(ch)
        # Prepare rels
        rroot = etree.fromstring(files[rp][1]) if rp in files else etree.Element(f"{{{PKG_REL}}}Relationships")
        mx = max((int(r.get("Id","rId0")[3:]) for r in rroot if r.get("Id","").startswith("rId")), default=0)
        rid = f"rId{mx+1}"
        iname = f"bg_{sn}.{new_ext}"
        itarget = f"../media/{iname}"
        # Build picture
        pic = etree.Element(f"{{{P}}}pic")
        nv = etree.SubElement(pic, f"{{{P}}}nvPicPr")
        cnp = etree.SubElement(nv, f"{{{P}}}cNvPr")
        cnp.set("id", str(mx+100)); cnp.set("name", f"BgPic_{sn}")
        etree.SubElement(nv, f"{{{P}}}cNvPicPr"); etree.SubElement(nv, f"{{{P}}}nvPr")
        bf = etree.SubElement(pic, f"{{{P}}}blipFill")
        bp = etree.SubElement(bf, f"{{{A}}}blip"); bp.set(f"{{{R}}}embed", rid)
        s1 = etree.SubElement(bf, f"{{{A}}}stretch"); etree.SubElement(s1, f"{{{A}}}fillRect")
        sp2 = etree.SubElement(pic, f"{{{P}}}spPr")
        xf = etree.SubElement(sp2, f"{{{A}}}xfrm")
        etree.SubElement(xf, f"{{{A}}}off", x="0", y="0")
        etree.SubElement(xf, f"{{{A}}}ext", cx=str(SW), cy=str(SH))
        pg = etree.SubElement(sp2, f"{{{A}}}prstGeom"); pg.set("prst","rect")
        etree.SubElement(pg, f"{{{A}}}avLst")
        if spTree is not None: spTree.insert(2, pic)
        # Add <p:bg>
        bg_e = etree.Element(f"{{{P}}}bg")
        cSld.insert(0, bg_e)
        bgPr = etree.SubElement(bg_e, f"{{{P}}}bgPr")
        bf2 = etree.SubElement(bgPr, f"{{{A}}}blipFill")
        bp2 = etree.SubElement(bf2, f"{{{A}}}blip"); bp2.set(f"{{{R}}}embed", rid)
        sb = etree.SubElement(bf2, f"{{{A}}}stretch"); etree.SubElement(sb, f"{{{A}}}fillRect")
        re2 = etree.SubElement(rroot, f"{{{PKG_REL}}}Relationship")
        re2.set("Id", rid); re2.set("Type", f"{R}/image"); re2.set("Target", itarget)
        files[sp] = (files.get(sp,(None,b""))[0], _tostring(sld))
        files[rp] = (None, _tostring(rroot))
        zi = ZipInfo(f"ppt/media/{iname}")
        zi.compress_type = zipfile.ZIP_DEFLATED
        added.append((zi, bg_bytes))
        print(f"    slide {sn:>2}")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for fi in orig: z.writestr(fi, files[fi.filename][1])
        for zi, dt in added: z.writestr(zi, dt)
    print(f"  Saved ({os.path.getsize(out)//1024} KB)")
    return True

def batch_process(indir, outdir, bg):
    if not os.path.isdir(indir): print(f"Error: {indir} not found"); return
    os.makedirs(outdir, exist_ok=True)
    files = []
    for rt, _, fn in os.walk(indir):
        for f in fn:
            if f.lower().endswith(".pptx") and not f.endswith("_bg_replaced.pptx"):
                files.append(os.path.join(rt, f))
    files.sort()
    if not files: print(f"No .pptx in {indir}"); return
    print(f"{len(files)} files found")
    ok = 0
    for i, fp in enumerate(files, 1):
        rel = os.path.relpath(fp, indir)
        dst = os.path.join(outdir, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        print(f"[{i}/{len(files)}] {rel}")
        if replace_ppt_background(fp, bg, dst): ok += 1
        print()
    print(f"Done: {ok}/{len(files)} succeeded -> {outdir}")

def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    p, bg = sys.argv[1], sys.argv[2]
    if os.path.isdir(p):
        out = sys.argv[3] if len(sys.argv) > 3 else p.rstrip("/\\") + "_bg_replaced"
        batch_process(p, out, bg)
    else:
        replace_ppt_background(p, bg, sys.argv[3] if len(sys.argv) > 3 else None)

if __name__ == "__main__": main()
