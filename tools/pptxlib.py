# -*- coding: utf-8 -*-
"""Dependency-free .pptx writer with a red/gold/white design kit.

Packaging note
--------------
The theme, slide master, slide layouts and notes master are NOT hand-written. They are
copied verbatim from a donor .pptx that PowerPoint itself produced (see DONOR below).
Hand-rolling those parts is what made earlier builds unreadable in PowerPoint: they are
the parts with the strictest content models, and a single ordering mistake in the theme
or master makes PowerPoint reject the entire file rather than just that slide.

Only the parts we genuinely need to author are generated: the slides, the notes slides,
presentation.xml, the relationship parts and the content types.
"""
import os, re, shutil, zipfile
from xml.sax.saxutils import escape

# Donor deck: a PowerPoint-authored file whose boilerplate parts we reuse.
DONOR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "Finance", "MaxFashion_AI_Consulting(Lingesh K).pptx")

# Parts copied byte-for-byte from the donor.
_DONOR_PARTS = [
    "ppt/theme/theme1.xml",
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml",
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    "ppt/slideLayouts/slideLayout2.xml",
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    "ppt/notesMasters/notesMaster1.xml",
    "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    "ppt/presProps.xml",
    "ppt/viewProps.xml",
    "ppt/tableStyles.xml",
]
LAYOUT_FOR_SLIDES = "../slideLayouts/slideLayout2.xml"      # the donor's BLANK layout

# ---------------------------------------------------------------- palette
RED_DEEP = "5E0B15"
RED      = "9E1B32"
RED_MID  = "B02338"
RED_SOFT = "F6E7EA"
GOLD     = "C9A227"
GOLD_LT  = "E3C567"
GOLD_PALE= "FAF1D8"
GOLD_DK  = "8A6E10"
WHITE    = "FFFFFF"
OFFWHITE = "FCFAF6"
INK      = "22222A"
INK_SOFT = "4A4A55"
GREY     = "7A7A85"
GREY_LT  = "E6E2DA"
GREEN    = "1E7A4D"
GREEN_LT = "2FA46B"

HEAD = "Georgia"
BODY = "Calibri"

SW, SH = 12192000, 6858000

def EI(v): return int(round(v * 914400.0))
def PT(v): return int(round(v * 12700.0))

_ids = [1]
def nid():
    _ids[0] += 1
    return _ids[0]
def reset_ids():
    _ids[0] = 1

# ---------------------------------------------------------------- text
def run(t, sz=13, b=False, i=False, color=INK, font=BODY, spc=0, caps=None, u=None):
    return dict(t=t, sz=sz, b=b, i=i, color=color, font=font, spc=spc, caps=caps, u=u)

def para(runs=None, align="l", bef=0, aft=0, line=None, marL=None, indent=None,
         bullet=None, buColor=None, buSz=None):
    return dict(runs=runs or [], align=align, bef=bef, aft=aft, line=line,
                marL=marL, indent=indent, bullet=bullet, buColor=buColor, buSz=buSz)

def _rpr(r):
    a = ['lang="en-IN" sz="%d"' % int(round(r["sz"] * 100))]
    if r["b"]: a.append('b="1"')
    if r["i"]: a.append('i="1"')
    if r["u"]: a.append('u="%s"' % r["u"])
    if r["spc"]: a.append('spc="%d"' % int(r["spc"]))
    if r["caps"]: a.append('cap="%s"' % r["caps"])
    a.append('dirty="0"')
    return ('<a:rPr %s><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:cs typeface="%s"/></a:rPr>'
            % (" ".join(a), r["color"], r["font"], r["font"]))

def _ppr(p):
    at = []
    if p["marL"] is not None: at.append('marL="%d"' % p["marL"])
    if p["indent"] is not None: at.append('indent="%d"' % p["indent"])
    at.append('algn="%s"' % {"l": "l", "c": "ctr", "r": "r", "j": "just"}[p["align"]])
    inner = ""
    if p["line"]:  inner += '<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>' % int(p["line"] * 1000)
    if p["bef"]:   inner += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % int(p["bef"] * 100)
    if p["aft"]:   inner += '<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % int(p["aft"] * 100)
    if p["bullet"]:
        if p["buColor"]: inner += '<a:buClr><a:srgbClr val="%s"/></a:buClr>' % p["buColor"]
        if p["buSz"]:    inner += '<a:buSzPct val="%d"/>' % int(p["buSz"] * 1000)
        inner += '<a:buFont typeface="Arial" pitchFamily="34" charset="0"/>'
        inner += '<a:buChar char="%s"/>' % escape(p["bullet"])
    else:
        inner += '<a:buNone/>'
    return '<a:pPr %s>%s</a:pPr>' % (" ".join(at), inner)

def _paras(paras):
    out = []
    for p in paras:
        s = "<a:p>" + _ppr(p)
        for r in p["runs"]:
            s += "<a:r>" + _rpr(r) + "<a:t>" + escape(r["t"]) + "</a:t></a:r>"
        out.append(s + "</a:p>")
    return "".join(out) or "<a:p><a:pPr/></a:p>"

def _txbody(paras, anchor="t", ins=(0.08, 0.04, 0.08, 0.04), wrap=True, autofit="none"):
    l, t, r, b = ins
    fit = {"none": "<a:noAutofit/>", "norm": "<a:normAutofit/>", "shape": "<a:spAutoFit/>"}[autofit]
    return ('<p:txBody><a:bodyPr wrap="%s" lIns="%d" tIns="%d" rIns="%d" bIns="%d" '
            'anchor="%s" anchorCtr="0">%s</a:bodyPr><a:lstStyle/>%s</p:txBody>'
            % ("square" if wrap else "none", EI(l), EI(t), EI(r), EI(b), anchor, fit, _paras(paras)))

# ---------------------------------------------------------------- fills / lines
def solid(c, alpha=None):
    if alpha is None:
        return '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % c
    return ('<a:solidFill><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:solidFill>'
            % (c, int(alpha * 1000)))

def nofill(): return "<a:noFill/>"

def grad(c1, c2, ang=5400000, c3=None):
    stops = '<a:gs pos="0"><a:srgbClr val="%s"/></a:gs>' % c1
    if c3: stops += '<a:gs pos="55000"><a:srgbClr val="%s"/></a:gs>' % c3
    stops += '<a:gs pos="100000"><a:srgbClr val="%s"/></a:gs>' % c2
    return ('<a:gradFill flip="none" rotWithShape="1"><a:gsLst>%s</a:gsLst>'
            '<a:lin ang="%d" scaled="0"/></a:gradFill>' % (stops, int(ang)))

def line(c=None, w=1.0, alpha=None, dash="solid"):
    if c is None:
        return "<a:ln><a:noFill/></a:ln>"
    return ('<a:ln w="%d" cap="flat" cmpd="sng" algn="ctr">%s<a:prstDash val="%s"/>'
            '<a:round/></a:ln>' % (PT(w), solid(c, alpha), dash))

def shadow(blur=18, dist=5, alpha=9.0, direction=5400000):
    return ('<a:effectLst><a:outerShdw blurRad="%d" dist="%d" dir="%d" rotWithShape="0">'
            '<a:srgbClr val="000000"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>'
            % (PT(blur), PT(dist), direction, int(alpha * 1000)))

def fillx(f):
    """Accept either a hex colour or an already-built fill element."""
    return f if f.startswith("<") else solid(f)

# ---------------------------------------------------------------- shapes
def shape(x, y, cx, cy, geom="rect", fill=None, ln=None, eff="", paras=None,
          anchor="t", ins=(0.08, 0.04, 0.08, 0.04), adj=None, rot=None, name=None,
          wrap=True, autofit="none", flipH=False, flipV=False, path=None):
    at = ""
    if rot: at += ' rot="%d"' % int(rot)
    if flipH: at += ' flipH="1"'
    if flipV: at += ' flipV="1"'
    xfrm = ('<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            % (at, x, y, cx, cy))
    if path is not None:
        g = ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
             '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst>%s</a:pathLst></a:custGeom>' % path)
    else:
        av = "<a:avLst/>"
        if adj is not None:
            av = "<a:avLst>%s</a:avLst>" % "".join(
                '<a:gd name="adj%s" fmla="val %d"/>' % ("" if len(adj) == 1 else str(k + 1), v)
                for k, v in enumerate(adj))
        g = '<a:prstGeom prst="%s">%s</a:prstGeom>' % (geom, av)
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s%s%s%s%s</p:spPr>%s</p:sp>'
            % (nid(), escape(name or "Shape"), xfrm, g,
               fillx(fill) if fill is not None else nofill(),
               ln if ln is not None else line(None), eff,
               _txbody(paras or [], anchor=anchor, ins=ins, wrap=wrap, autofit=autofit)))

def text(x, y, cx, cy, paras, anchor="t", ins=(0, 0, 0, 0), wrap=True, name="Text",
         autofit="none"):
    return shape(x, y, cx, cy, fill=nofill(), ln=line(None), paras=paras, anchor=anchor,
                 ins=ins, name=name, wrap=wrap, autofit=autofit)

# ---------------------------------------------------------------- table
def cell(paras, fill=None, borders=None, ins=(0.07, 0.035, 0.07, 0.035), anchor="ctr",
         span=None, hmerge=False):
    return dict(paras=paras, fill=fill, borders=borders or {}, ins=ins, anchor=anchor,
                span=span, hmerge=hmerge)

def _cell_xml(c):
    l, t, r, b = c["ins"]
    bd = ""
    for side, tag in (("l", "lnL"), ("r", "lnR"), ("t", "lnT"), ("b", "lnB")):
        spec = c["borders"].get(side)
        if spec is None:
            bd += '<a:%s><a:noFill/></a:%s>' % (tag, tag)
        else:
            col, wd = spec
            bd += ('<a:%s w="%d" cap="flat" cmpd="sng" algn="ctr">%s<a:prstDash val="solid"/>'
                   '<a:round/></a:%s>' % (tag, PT(wd), solid(col), tag))
    body = (_txbody(c["paras"], anchor=c["anchor"], ins=(0, 0, 0, 0))
            .replace("<p:txBody>", "<a:txBody>").replace("</p:txBody>", "</a:txBody>"))
    at = ""
    if c.get("span"): at += ' gridSpan="%d"' % c["span"]
    if c.get("hmerge"): at += ' hMerge="1"'
    return ('<a:tc%s>%s<a:tcPr marL="%d" marR="%d" marT="%d" marB="%d" anchor="%s">%s%s</a:tcPr>'
            '</a:tc>' % (at, body, EI(l), EI(r), EI(t), EI(b), c["anchor"], bd,
                         fillx(c["fill"]) if c["fill"] is not None else nofill()))

def table(x, y, colw, rows, rowh, name="Table"):
    grid = "".join('<a:gridCol w="%d"/>' % w for w in colw)
    trs = "".join('<a:tr h="%d">%s</a:tr>' % (rowh[i], "".join(_cell_xml(c) for c in rw))
                  for i, rw in enumerate(rows))
    return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="%s"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            '<p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            '<a:tbl><a:tblPr/><a:tblGrid>%s</a:tblGrid>%s</a:tbl></a:graphicData></a:graphic>'
            '</p:graphicFrame>'
            % (nid(), escape(name), x, y, sum(colw), sum(rowh), grid, trs))

# ---------------------------------------------------------------- connectors
def vline(cx, y, h, col=GREY_LT, w=0.016):
    return shape(cx - EI(w) // 2, y, EI(w), h, fill=solid(col), ln=line(None), name="VLine")

def hline(x, y, w, col=GREY_LT, h=0.016):
    return shape(x, y - EI(h) // 2, w, EI(h), fill=solid(col), ln=line(None), name="HLine")

def arrow_down(cx, y, h, col=GREY, w=0.016, head=0.11):
    stem = max(EI(0.02), h - EI(head))
    return [shape(cx - EI(w) // 2, y, EI(w), stem, fill=solid(col), ln=line(None), name="Stem"),
            shape(cx - EI(head * 0.62), y + stem, EI(head * 1.24), EI(head), geom="triangle",
                  rot=10800000, fill=solid(col), ln=line(None), name="Head")]

# ---------------------------------------------------------------- ISSM crest
_SHIELD = ('<a:path w="1000" h="1160">'
           '<a:moveTo><a:pt x="55" y="40"/></a:moveTo>'
           '<a:lnTo><a:pt x="945" y="40"/></a:lnTo>'
           '<a:lnTo><a:pt x="945" y="620"/></a:lnTo>'
           '<a:cubicBezTo><a:pt x="945" y="890"/><a:pt x="760" y="1040"/>'
           '<a:pt x="500" y="1120"/></a:cubicBezTo>'
           '<a:cubicBezTo><a:pt x="240" y="1040"/><a:pt x="55" y="890"/>'
           '<a:pt x="55" y="620"/></a:cubicBezTo>'
           '<a:close/></a:path>')
_FLAME = ('<a:path w="100" h="140">'
          '<a:moveTo><a:pt x="50" y="0"/></a:moveTo>'
          '<a:cubicBezTo><a:pt x="86" y="46"/><a:pt x="100" y="72"/><a:pt x="76" y="106"/>'
          '</a:cubicBezTo>'
          '<a:cubicBezTo><a:pt x="63" y="124"/><a:pt x="37" y="124"/><a:pt x="24" y="106"/>'
          '</a:cubicBezTo>'
          '<a:cubicBezTo><a:pt x="0" y="72"/><a:pt x="14" y="46"/><a:pt x="50" y="0"/>'
          '</a:cubicBezTo><a:close/></a:path>')
_PAGE_L = ('<a:path w="100" h="62">'
           '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
           '<a:lnTo><a:pt x="100" y="22"/></a:lnTo>'
           '<a:lnTo><a:pt x="100" y="54"/></a:lnTo>'
           '<a:lnTo><a:pt x="0" y="40"/></a:lnTo>'
           '<a:close/></a:path>')
_RIBBON = ('<a:path w="1000" h="100">'
           '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
           '<a:lnTo><a:pt x="1000" y="0"/></a:lnTo>'
           '<a:lnTo><a:pt x="925" y="50"/></a:lnTo>'
           '<a:lnTo><a:pt x="1000" y="100"/></a:lnTo>'
           '<a:lnTo><a:pt x="0" y="100"/></a:lnTo>'
           '<a:lnTo><a:pt x="75" y="50"/></a:lnTo>'
           '<a:close/></a:path>')

def issm_crest(x, y, h, ring=GOLD, body=RED_DEEP, on_dark=False):
    w = int(h * 0.862)
    out = [shape(x, y, w, h, fill=solid(body), ln=line(ring, 1.6 if h > EI(0.6) else 1.0),
                 path=_SHIELD, name="ISSM Crest")]
    pad = int(w * 0.075)
    out.append(shape(x + pad, y + int(h * 0.055), w - 2 * pad, int(h * 0.885),
                     fill=nofill(), ln=line(GOLD_LT, 0.75, alpha=70.0), path=_SHIELD,
                     name="Crest Inner"))
    fh = int(h * 0.160); fw = int(fh * 0.66)
    out.append(shape(x + (w - fw) // 2, y + int(h * 0.135), fw, fh, fill=solid(GOLD),
                     ln=line(None), path=_FLAME, name="Flame"))
    sw_ = max(EI(0.010), int(w * 0.042))
    out.append(shape(x + (w - sw_) // 2, y + int(h * 0.295), sw_, int(h * 0.145),
                     fill=solid(GOLD_LT), ln=line(None), name="Stem"))
    cw_ = int(w * 0.13)
    out.append(shape(x + (w - cw_) // 2, y + int(h * 0.425), cw_, max(EI(0.008), int(h * 0.022)),
                     fill=solid(GOLD), ln=line(None), name="Collar"))
    bw = int(w * 0.52); bh = int(bw * 0.42)
    bx = x + (w - bw) // 2; by = y + int(h * 0.475); half = int(bw * 0.475)
    out.append(shape(bx, by, half, bh, fill=solid(WHITE), ln=line(None), path=_PAGE_L,
                     name="Page L"))
    out.append(shape(bx + bw - half, by, half, bh, fill=solid(WHITE), ln=line(None),
                     path=_PAGE_L, flipH=True, name="Page R"))
    out.append(shape(bx + half, by + int(bh * 22 / 62.0), bw - 2 * half,
                     max(EI(0.008), int(bh * 32 / 62.0)), fill=solid(GOLD_LT), ln=line(None),
                     name="Spine"))
    rw = int(w * 0.56)
    out.append(shape(x + (w - rw) // 2, y + int(h * 0.695), rw, max(EI(0.020), int(h * 0.050)),
                     fill=solid(GOLD), ln=line(None), path=_RIBBON, name="Ribbon"))
    return out, w

def issm_lockup(x, y, h, on_dark=False, compact=False):
    shp, cw = issm_crest(x, y, h, on_dark=on_dark)
    gap = int(h * 0.16)
    tx = x + cw + gap
    name_col = WHITE if on_dark else RED
    sub_col = GOLD_LT if on_dark else GOLD
    tag_col = "D8CFC0" if on_dark else GREY
    inch = h / 914400.0
    if compact:
        tw = int(h * 3.05)
        shp.append(text(tx, y + int(h * 0.13), tw, int(h * 0.9), [
            para([run("ISSM", sz=inch * 22, b=True, color=name_col, font=HEAD, spc=40)]),
            para([run("BUSINESS SCHOOL", sz=inch * 10.2, b=True, color=sub_col, font=BODY,
                      spc=120)], bef=1)], wrap=False, name="ISSM Wordmark"))
        return shp, cw + gap + tw
    tw = int(h * 3.5)
    shp.append(text(tx, y + int(h * 0.055), tw, int(h * 1.0), [
        para([run("ISSM", sz=inch * 26, b=True, color=name_col, font=HEAD, spc=60)]),
        para([run("BUSINESS SCHOOL", sz=inch * 11.6, b=True, color=sub_col, font=BODY, spc=170)],
             bef=1),
        para([run("CHENNAI  |  HYDERABAD", sz=inch * 8.0, b=True, color=tag_col, font=BODY,
                  spc=130)], bef=3)], wrap=False, name="ISSM Wordmark"))
    return shp, cw + gap + tw

# ---------------------------------------------------------------- slide / notes XML
_TREE = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
         '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
         '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>%s</p:spTree>')

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')

def slide_xml(shapes, bg=WHITE):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld %s><p:cSld><p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>%s</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
            % (NS, fillx(bg), _TREE % "".join(shapes)))

def notes_slide(lines):
    """Mirrors the donor's notes-slide shape exactly: sldImg placeholder with an empty
    spPr and no txBody, then the body placeholder carrying the note text."""
    body = "".join(
        '<a:p><a:r><a:rPr lang="en-IN" dirty="0"/><a:t>%s</a:t></a:r></a:p>' % escape(l)
        for l in lines) or "<a:p><a:endParaRPr lang=\"en-IN\"/></a:p>"
    img = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>'
           '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>'
           '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>')
    ph = ('<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>'
          '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
          '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
          '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody></p:sp>' % body)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:notes %s><p:cSld>%s</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '</p:notes>' % (NS, _TREE % (img + ph)))

# ---------------------------------------------------------------- package
def write_pptx(path, slides, notes, title, author, donor=None):
    donor = donor or DONOR
    dz = zipfile.ZipFile(donor)
    missing = [p for p in _DONOR_PARTS if p not in dz.namelist()]
    if missing:
        raise RuntimeError("donor is missing required parts: %s" % missing)
    default_text_style = re.search(
        r"<p:defaultTextStyle>.*?</p:defaultTextStyle>",
        dz.read("ppt/presentation.xml").decode("utf-8"), re.S)
    dts = default_text_style.group(0) if default_text_style else ""

    n = len(slides)
    tmp = path + ".build"
    if os.path.isdir(tmp): shutil.rmtree(tmp)

    def W(rel, content):
        full = os.path.join(tmp, rel)
        os.makedirs(os.path.dirname(full), exist_ok=True)
        if isinstance(content, bytes): open(full, "wb").write(content)
        else: open(full, "w", encoding="utf-8").write(content)

    for p in _DONOR_PARTS:
        W(p, dz.read(p))

    ov = ['<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
          '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
          '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
          '<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
          '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>',
          '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
          '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
          '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>',
          '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>',
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
    for i in range(1, n + 1):
        ov.append('<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % i)
        ov.append('<Override PartName="/ppt/notesSlides/notesSlide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' % i)
    W("[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Default Extension="jpeg" ContentType="image/jpeg"/>'
      + "".join(ov) + "</Types>")

    W("_rels/.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
      '</Relationships>')

    W("docProps/core.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
      '<dc:title>%s</dc:title><dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>'
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-08-11T09:00:00Z</dcterms:created>'
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-08-11T09:00:00Z</dcterms:modified>'
      '</cp:coreProperties>' % (escape(title), escape(author), escape(author)))

    W("docProps/app.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
      '<Company>ISSM Business School</Company><Slides>%d</Slides>'
      '<Application>Microsoft Office PowerPoint</Application></Properties>' % n)

    # presentation.xml -- element order mirrors the donor, which PowerPoint accepts
    prels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    sldids = ""
    for i in range(1, n + 1):
        rid = "rId%d" % (i + 1)
        prels.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>' % (rid, i))
        sldids += '<p:sldId id="%d" r:id="%s"/>' % (255 + i, rid)
    nm_rid = "rId%d" % (n + 2)
    prels.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>' % nm_rid)
    prels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' % (n + 3))
    prels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>' % (n + 4))
    prels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' % (n + 5))
    prels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' % (n + 6))
    W("ppt/_rels/presentation.xml.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s'
      '</Relationships>' % "".join(prels))

    W("ppt/presentation.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<p:presentation %s saveSubsetFonts="1" autoCompressPictures="0">'
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
      '<p:sldIdLst>%s</p:sldIdLst>'
      '<p:notesMasterIdLst><p:notesMasterId r:id="%s"/></p:notesMasterIdLst>'
      '<p:sldSz cx="%d" cy="%d"/><p:notesSz cx="6858000" cy="12192000"/>%s'
      '</p:presentation>' % (NS, sldids, nm_rid, SW, SH, dts))

    for i, (sx, nl) in enumerate(zip(slides, notes), start=1):
        W("ppt/slides/slide%d.xml" % i, sx)
        W("ppt/slides/_rels/slide%d.xml.rels" % i,
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="%s"/>'
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide%d.xml"/>'
          '</Relationships>' % (LAYOUT_FOR_SLIDES, i))
        W("ppt/notesSlides/notesSlide%d.xml" % i, notes_slide(nl))
        W("ppt/notesSlides/_rels/notesSlide%d.xml.rels" % i,
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>'
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide%d.xml"/>'
          '</Relationships>' % i)

    if os.path.exists(path): os.remove(path)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        allf = []
        for root, _, files in os.walk(tmp):
            for f in files:
                allf.append(os.path.relpath(os.path.join(root, f), tmp).replace("\\", "/"))
        for first in ("[Content_Types].xml", "_rels/.rels"):
            if first in allf:
                z.write(os.path.join(tmp, first), first); allf.remove(first)
        for f in sorted(allf):
            z.write(os.path.join(tmp, f), f)
    shutil.rmtree(tmp)
    return path
