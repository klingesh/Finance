# -*- coding: utf-8 -*-
"""End-to-end package verification for a generated .pptx.

Checks:
  1  zip integrity and XML well-formedness of every part
  2  a content-type declaration exists for every part, and no override is orphaned
  3  every relationship target resolves to a real part
  4  every r:id / r:embed used in a part is declared in that part's .rels
  5  presentation.xml sldIdLst resolves to actual slide parts, ids in legal range
  6  each slide reaches a layout, and that layout reaches the master
  7  donor-supplied boilerplate is byte-identical to the donor
  8  attribute value ranges PowerPoint enforces (sz, spc, alpha, srgbClr)
  9  shapes stay on canvas, and text fits its box / table rows do not grow past the footer
"""
import sys, zipfile, posixpath, re
import xml.etree.ElementTree as ET

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EMU = 914400.0
SW, SH = 12192000, 6858000
FOOTER_Y = 6.74
ADV = {("Calibri", False): 0.465, ("Calibri", True): 0.487,
       ("Georgia", False): 0.505, ("Georgia", True): 0.535}
LINE = 1.21

PATH = sys.argv[1]
DONOR = sys.argv[2] if len(sys.argv) > 2 else None
errors, warns = [], []

z = zipfile.ZipFile(PATH)
if z.testzip(): errors.append("corrupt zip entry")
names = set(z.namelist())

trees = {}
for n in sorted(names):
    if n.endswith(".xml") or n.endswith(".rels"):
        try:
            trees[n] = ET.fromstring(z.read(n))
        except ET.ParseError as e:
            errors.append("%s: not well formed: %s" % (n, e))

# 2 content types
ct = trees.get("[Content_Types].xml")
defaults, overrides = set(), set()
for c in (ct if ct is not None else []):
    tag = c.tag.split("}")[-1]
    if tag == "Default": defaults.add(c.get("Extension").lower())
    elif tag == "Override": overrides.add(c.get("PartName"))
for n in names:
    if n == "[Content_Types].xml": continue
    if "/" + n not in overrides and n.rsplit(".", 1)[-1].lower() not in defaults:
        errors.append("no content type declared for %s" % n)
for o in overrides:
    if o.lstrip("/") not in names:
        errors.append("content-type override points at missing part %s" % o)

# 3 + 4 relationships
relmap = {}
for n, t in trees.items():
    if not n.endswith(".rels"): continue
    base = posixpath.dirname(posixpath.dirname(n))
    owner = posixpath.join(base, posixpath.basename(n)[:-5]) if base else posixpath.basename(n)[:-5]
    m = {}
    for rel in t:
        if rel.get("TargetMode") == "External": continue
        res = posixpath.normpath(posixpath.join(base, rel.get("Target"))).lstrip("/")
        m[rel.get("Id")] = res
        if res not in names:
            errors.append("%s: target %s does not exist" % (n, rel.get("Target")))
    relmap[owner] = m
for n in sorted(names):
    if not n.endswith(".xml"): continue
    relp = posixpath.join(posixpath.dirname(n), "_rels", posixpath.basename(n) + ".rels")
    have = set(relmap.get(n, {}))
    used = set(re.findall(r'r:(?:id|embed)="([^"]+)"', z.read(n).decode("utf-8", "replace")))
    for u in used - have:
        errors.append("%s uses %s but it is not declared in %s" % (n, u, relp))

# 5 slide list
nslides = len([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)])
pres = trees.get("ppt/presentation.xml")
prel = relmap.get("ppt/presentation.xml", {})
if pres is not None:
    lst = pres.find(P + "sldIdLst")
    ids = [s.get(R + "id") for s in (lst if lst is not None else [])]
    if len(ids) != nslides:
        errors.append("sldIdLst has %d entries but there are %d slide parts" % (len(ids), nslides))
    for rid in ids:
        tgt = prel.get(rid)
        if not tgt or not re.match(r"ppt/slides/slide\d+\.xml$", tgt or ""):
            errors.append("sldIdLst entry %s does not resolve to a slide" % rid)
    nums = [int(s.get("id")) for s in (lst if lst is not None else [])]
    if len(set(nums)) != len(nums): errors.append("duplicate p:sldId id values")
    if any(v < 256 or v > 2147483647 for v in nums): errors.append("p:sldId id out of range")

# 6 slide -> layout -> master
for i in range(1, nslides + 1):
    sp = "ppt/slides/slide%d.xml" % i
    lay = [t for t in relmap.get(sp, {}).values() if "slideLayouts/" in t]
    if not lay:
        errors.append("%s has no slideLayout relationship" % sp); continue
    mrel = relmap.get(lay[0], {})
    if not any("slideMasters/" in t for t in mrel.values()):
        errors.append("%s does not reference a slideMaster" % lay[0])

# 7 donor parts identical
if DONOR:
    dz = zipfile.ZipFile(DONOR)
    for p in ["ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml",
              "ppt/slideLayouts/slideLayout2.xml", "ppt/notesMasters/notesMaster1.xml",
              "ppt/presProps.xml", "ppt/viewProps.xml", "ppt/tableStyles.xml"]:
        if p not in names:
            errors.append("donor part %s missing from output" % p)
        elif z.read(p) != dz.read(p):
            errors.append("donor part %s was modified (should be byte-identical)" % p)

# 8 attribute ranges
for n, t in trees.items():
    for e in t.iter(A + "srgbClr"):
        if not re.fullmatch(r"[0-9A-Fa-f]{6}", e.get("val") or ""):
            errors.append("%s: invalid srgbClr %r" % (n, e.get("val")))
    for e in t.iter(A + "alpha"):
        v = int(e.get("val"))
        if not (0 <= v <= 100000): errors.append("%s: alpha out of range %d" % (n, v))
    for e in t.iter():
        tag = e.tag.split("}")[-1]
        if tag in ("rPr", "defRPr", "endParaRPr"):
            sz, spc = e.get("sz"), e.get("spc")
            if sz is not None:
                if not re.fullmatch(r"\d+", sz) or not (100 <= int(sz) <= 400000):
                    errors.append("%s: illegal sz=%r" % (n, sz))
            if spc is not None:
                if not re.fullmatch(r"-?\d+", spc) or not (-400000 <= int(spc) <= 400000):
                    errors.append("%s: illegal spc=%r" % (n, spc))

# 9 geometry + text fit
def est_para(p, width_in):
    runs = p.findall(A + "r")
    if not runs: return 0.0
    ppr = p.find(A + "pPr")
    marL, lnpct, bef = 0.0, 100.0, 0.0
    if ppr is not None:
        marL = float(ppr.get("marL") or 0) / EMU
        ls = ppr.find(A + "lnSpc/" + A + "spcPct")
        if ls is not None: lnpct = float(ls.get("val")) / 1000.0
        sb = ppr.find(A + "spcBef/" + A + "spcPts")
        if sb is not None: bef = float(sb.get("val")) / 100.0 / 72.0
    avail = max(0.35, width_in - marL)
    w, maxsz = 0.0, 0.0
    for r in runs:
        rpr = r.find(A + "rPr")
        sz = float(rpr.get("sz") or 1800) / 100.0
        bold = rpr.get("b") == "1"
        spc = float(rpr.get("spc") or 0) / 100.0
        lat = rpr.find(A + "latin")
        font = lat.get("typeface") if lat is not None else "Calibri"
        w += len(r.find(A + "t").text or "") * (ADV.get((font, bold), 0.5) * sz + spc) / 72.0
        maxsz = max(maxsz, sz)
    lines = max(1, int(w / avail) + (1 if w % avail else 0))
    return lines * maxsz * LINE * (lnpct / 100.0) / 72.0 + bef

def body_h(tb, width_in):
    if tb is None: return 0.0
    bpr = tb.find(A + "bodyPr")
    lI = float(bpr.get("lIns") or 91440) / EMU
    rI = float(bpr.get("rIns") or 91440) / EMU
    tI = float(bpr.get("tIns") or 45720) / EMU
    bI = float(bpr.get("bIns") or 45720) / EMU
    return tI + bI + sum(est_para(p, max(0.3, width_in - lI - rI)) for p in tb.findall(A + "p"))

for i in range(1, nslides + 1):
    t = trees["ppt/slides/slide%d.xml" % i]
    tree = t.find(P + "cSld/" + P + "spTree")
    sids = [e.get("id") for e in tree.iter(P + "cNvPr")]
    if len(set(sids)) != len(sids): errors.append("slide%d: duplicate shape ids" % i)
    for sp in tree.findall(P + "sp"):
        xf = sp.find(P + "spPr/" + A + "xfrm")
        if xf is None: continue
        o, e = xf.find(A + "off"), xf.find(A + "ext")
        x, y = int(o.get("x")), int(o.get("y"))
        cx, cy = int(e.get("cx")), int(e.get("cy"))
        nm = sp.find(".//" + P + "cNvPr").get("name") or ""
        if not nm.startswith("Arc"):
            if y + cy > SH + 6000 and cy < SH:
                warns.append("slide%d '%s' bottom %.2fin" % (i, nm, (y + cy) / EMU))
            if x + cx > SW + 6000 and cx < SW:
                warns.append("slide%d '%s' right %.2fin" % (i, nm, (x + cx) / EMU))
        if not sp.findall(P + "txBody/" + A + "p/" + A + "r") or nm.startswith("Footer"):
            continue
        h = body_h(sp.find(P + "txBody"), cx / EMU)
        if h > cy / EMU + 0.06:
            issues = "slide%d OVERFLOW '%s' needs %.2fin, box %.2fin" % (i, nm, h, cy / EMU)
            warns.append(issues)
        if y / EMU + h > FOOTER_Y + 0.04:
            warns.append("slide%d '%s' text reaches %.2fin (footer %.2f)" % (i, nm, y / EMU + h,
                                                                            FOOTER_Y))
    for gf in tree.findall(P + "graphicFrame"):
        y = int(gf.find(P + "xfrm/" + A + "off").get("y")) / EMU
        tbl = gf.find(".//" + A + "tbl")
        cols = [int(g.get("w")) / EMU for g in tbl.find(A + "tblGrid")]
        grown = 0.0
        for tr in tbl.findall(A + "tr"):
            need = int(tr.get("h")) / EMU
            for ci, tc in enumerate(tr.findall(A + "tc")):
                if tc.get("hMerge") == "1": continue
                span = int(tc.get("gridSpan") or 1)
                tp = tc.find(A + "tcPr")
                mL = float(tp.get("marL") or 91440) / EMU
                mR = float(tp.get("marR") or 91440) / EMU
                mT = float(tp.get("marT") or 45720) / EMU
                mB = float(tp.get("marB") or 45720) / EMU
                ch = mT + mB + sum(est_para(p, max(0.25, sum(cols[ci:ci + span]) - mL - mR))
                                   for p in tc.find(A + "txBody").findall(A + "p"))
                need = max(need, ch)
            grown += need
        if y + grown > FOOTER_Y + 0.04:
            warns.append("slide%d table bottom %.2fin" % (i, y + grown))
    if not z.read("ppt/notesSlides/notesSlide%d.xml" % i).decode("utf-8").count("<a:t>"):
        warns.append("slide%d has no speaker notes" % i)

print("parts: %d   slides: %d" % (len(names), nslides))
print("WARNINGS (%d)" % len(warns))
for w in warns: print("   -", w)
print("ERRORS (%d)" % len(errors))
for e in errors: print("   !", e)
sys.exit(1 if errors else 0)
