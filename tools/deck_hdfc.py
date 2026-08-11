# -*- coding: utf-8 -*-
"""HDFC Bank - AI-Driven Digital Banking Transformation (Case Study 2)."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptxlib import *

NAME  = "Prahadhesvaryaa K S"
REGNO = "OSI2509093"
TOTAL = 6

M     = EI(0.66)
CW    = SW - 2 * M
CT_   = EI(1.72)
SPINE = EI(0.13)

SLIDES, NOTES = [], []

def spine():
    return [shape(0, 0, SPINE, SH, fill=grad(RED_DEEP, RED, 5400000), ln=line(None), name="Spine"),
            shape(SPINE, 0, EI(0.022), SH, fill=solid(GOLD), ln=line(None), name="Spine Gold")]

def footer(pageno):
    out = [shape(M, EI(6.74), CW, EI(0.011), fill=solid(GREY_LT), ln=line(None),
                 name="Footer Rule")]
    out.append(text(M, EI(6.83), EI(7.4), EI(0.3), [para([
        run(NAME, sz=8.8, b=True, color=GREY, font=BODY, spc=60),
        run("   |   Register No: ", sz=8.8, color=GREY, font=BODY, spc=40),
        run(REGNO, sz=8.8, b=True, color=GOLD, font=BODY, spc=60)])], name="Footer L"))
    out.append(text(SW - M - EI(2.2), EI(6.83), EI(2.2), EI(0.3), [para([
        run("%02d" % pageno, sz=9.6, b=True, color=RED, font=HEAD),
        run("  /  %d" % TOTAL, sz=8.8, color=GREY, font=BODY)], align="r")], name="Footer R"))
    return out

def header(eyebrow, title, kicker=None, pageno=1):
    out = spine()
    lk, _w = issm_lockup(SW - M - EI(2.22), EI(0.40), EI(0.50), compact=True)
    out += lk
    out.append(text(M, EI(0.46), EI(7.6), EI(0.26), [para([
        run(eyebrow.upper(), sz=9.4, b=True, color=GOLD, font=BODY, spc=260)])], name="Eyebrow"))
    out.append(text(M, EI(0.74), EI(9.4), EI(0.62), [para([
        run(title, sz=29, b=True, color=RED_DEEP, font=HEAD, spc=-15)], line=98)], name="Title"))
    out.append(shape(M, EI(1.38), EI(1.05), EI(0.045), fill=solid(GOLD), ln=line(None),
                     name="Rule"))
    if kicker:
        out.append(text(M + EI(1.22), EI(1.29), CW - EI(1.30), EI(0.32), [para([
            run(kicker, sz=10.6, i=True, color=INK_SOFT, font=BODY)], line=112)], name="Kicker"))
    out += footer(pageno)
    return out

def card(x, y, cx, cy, fill=WHITE, border=GREY_LT, accent=None, accent_w=0.05, radius=None):
    out = []
    geom, adj = ("roundRect", (radius,)) if radius else ("rect", None)
    out.append(shape(x, y, cx, cy, geom=geom, adj=adj, fill=fill,
                     ln=line(border, 0.75) if border else line(None), eff=shadow(), name="Card"))
    if accent:
        out.append(shape(x, y, EI(accent_w), cy, fill=solid(accent), ln=line(None), name="Accent"))
    return out

def pill(x, y, w, h, label, fill=GOLD_PALE, border=GOLD, color=GOLD_DK, sz=8.0, spc=130):
    return shape(x, y, w, h, geom="roundRect", adj=(50000,), fill=fill,
                 ln=line(border, 0.7) if border else line(None), anchor="ctr", name="Pill",
                 paras=[para([run(label, sz=sz, b=True, color=color, font=BODY, spc=spc)],
                             align="c")])

def steps(x, y, w, items, col=GOLD, sz=9.6, gap=0.27):
    out = []
    for i, t in enumerate(items):
        yy = y + EI(gap) * i
        out.append(text(x, yy, EI(0.22), EI(0.26), [para([
            run("\u25b8", sz=10.5, b=True, color=col, font=BODY)])], name="Chev"))
        out.append(text(x + EI(0.22), yy, w - EI(0.22), EI(0.26), [para([
            run(t, sz=sz, color=INK_SOFT, font=BODY)], line=108)], name="Step"))
    return out

def add(shapes, notes, bg=OFFWHITE):
    SLIDES.append(slide_xml(shapes, bg=bg))
    NOTES.append(notes)

# ================================================================== 01 TITLE
def s01():
    reset_ids()
    s = [shape(0, 0, SW, SH, fill=solid(OFFWHITE), ln=line(None), name="BG")]
    s += spine()
    s.append(shape(EI(9.4), EI(-1.9), EI(6.4), EI(6.4), geom="ellipse", fill=nofill(),
                   ln=line(GOLD, 1.0, alpha=26.0), name="Arc1"))
    s.append(shape(EI(10.3), EI(-1.0), EI(6.4), EI(6.4), geom="ellipse", fill=nofill(),
                   ln=line(GOLD, 0.9, alpha=16.0), name="Arc2"))
    lk, _lw = issm_lockup(SW - M - EI(3.64), EI(0.54), EI(0.80))
    s += lk
    s.append(pill(M, EI(2.10), EI(2.42), EI(0.36), "CASE STUDY 2", fill=GOLD, border=None,
                  color=RED_DEEP, sz=10.2, spc=200))
    s.append(text(M + EI(2.60), EI(2.14), EI(6.0), EI(0.3), [para([
        run("AI IN BANKING  |  DIGITAL TRANSFORMATION", sz=9.8, b=True, color=GREY, font=BODY,
            spc=170)])], name="Badge Sub"))
    s.append(text(M, EI(2.66), EI(10.4), EI(0.9), [para([
        run("HDFC Bank", sz=52, b=True, color=RED_DEEP, font=HEAD, spc=-30)], line=94)],
        name="Main Title"))
    s.append(text(M, EI(3.50), EI(10.6), EI(0.7), [para([
        run("AI-Driven Digital Banking Transformation", sz=27, b=True, color=RED, font=HEAD,
            spc=-10)], line=100)], name="Subtitle"))
    s.append(shape(M, EI(4.24), EI(1.5), EI(0.05), fill=solid(GOLD), ln=line(None), name="Rule"))
    chips = ["RPA", "CHATBOTS", "PLATFORM AS A SERVICE", "ADAPTIVE FRAUD PREVENTION"]
    pw = (EI(10.0) - EI(0.20) * 3) // 4
    for k, c in enumerate(chips):
        s.append(pill(M + k * (pw + EI(0.20)), EI(4.50), pw, EI(0.42), c, fill=WHITE,
                      border=GOLD, color=RED, sz=9.0, spc=110))
    s += card(M, EI(5.26), EI(5.45), EI(1.10), fill=grad(RED_DEEP, RED, 2700000), border=None,
              radius=4000)
    s.append(shape(M, EI(5.26), EI(0.05), EI(1.10), fill=solid(GOLD), ln=line(None), name="Bar"))
    s.append(text(M + EI(0.34), EI(5.40), EI(5.0), EI(0.9), [
        para([run("PRESENTED BY", sz=8.4, b=True, color=GOLD_LT, font=BODY, spc=200)]),
        para([run(NAME, sz=19, b=True, color=WHITE, font=HEAD)], bef=3),
        para([run("MBA \u2014 Finance   |   Register No. ", sz=9.8, color="E4D2C6", font=BODY),
              run(REGNO, sz=9.8, b=True, color=GOLD_LT, font=BODY, spc=50)], bef=3)],
        name="Presenter"))
    s.append(text(SW - M - EI(4.0), EI(5.62), EI(4.0), EI(0.7), [
        para([run("SUBMITTED TO", sz=8.4, b=True, color=GOLD, font=BODY, spc=200)], align="r"),
        para([run("ISSM Business School", sz=13.4, b=True, color=RED_DEEP, font=HEAD)],
             align="r", bef=3),
        para([run("Chennai  |  Hyderabad", sz=9.2, color=GREY, font=BODY, spc=60)],
             align="r", bef=2)], name="Submitted"))
    add(s, ["Open by naming the four AI capabilities the case covers: RPA, chatbots, PaaS and "
            "adaptive fraud prevention.",
            "One-line framing: HDFC uses AI to make digital banking faster, safer and more "
            "convenient - the deck takes each of those in turn.",
            "Introduce yourself and your register number, then move on. Around 20 seconds."])

# ================================================================== 02 RPA + CHATBOTS
def s02():
    reset_ids()
    s = header("Pillars 1 and 2 \u2014 automation and service", "AI-Driven Banking Solutions",
               "Both target the same weakness: tasks that are high-volume, rule-based and "
               "time-critical, where waiting for a human adds cost but no judgement", pageno=2)
    cw = (CW - EI(0.36)) // 2
    blocks = [
        dict(x=M, tag="RPA", full="Robotic Process Automation", pilltxt="KYC VERIFICATION",
             stat="10", unit="seconds", statlab="ACCOUNT OPENED IN",
             body="Copies uploaded ID details, checks them against government databases, and "
                  "creates the account automatically \u2014 no branch visit, no manual data entry, "
                  "no overnight batch.",
             flow=["Customer uploads ID document",
                   "Bot reads and extracts the details",
                   "Details cross-checked against government databases",
                   "Account created and activated automatically"],
             col=RED, band=grad(RED_DEEP, RED, 0), sub=GOLD_LT),
        dict(x=M + cw + EI(0.36), tag="Chatbots", full="Conversational AI Support",
             pilltxt="24/7 CUSTOMER SUPPORT", stat="24/7", unit="",
             statlab="AVAILABLE WITHOUT A HUMAN AGENT",
             body="Instantly blocks a lost debit card and orders a replacement at 2 AM \u2014 at "
                  "the moment the customer notices, not when the call centre reopens.",
             flow=["Customer reports a lost card at 2 AM",
                   "Chatbot authenticates the customer",
                   "Card is blocked immediately",
                   "Replacement card ordered and confirmed"],
             col=GOLD_DK, band=grad(GOLD_DK, GOLD, 0), sub="FFF3D0"),
    ]
    for b in blocks:
        x = b["x"]
        s += card(x, CT_, cw, EI(4.72), radius=4000)
        s.append(shape(x, CT_, cw, EI(0.76), fill=b["band"], ln=line(None), name="Band"))
        s.append(text(x + EI(0.28), CT_ + EI(0.10), cw - EI(0.56), EI(0.6), [
            para([run(b["tag"].upper(), sz=17, b=True, color=WHITE, font=HEAD, spc=60)]),
            para([run(b["full"], sz=9.4, b=True, color=b["sub"], font=BODY, spc=110)], bef=2)],
            name="BandTxt"))
        s.append(pill(x + EI(0.28), CT_ + EI(0.94), EI(2.30), EI(0.32), b["pilltxt"]))
        s.append(text(x + EI(0.28), CT_ + EI(1.42), cw - EI(0.56), EI(0.78), [
            para([run(b["statlab"], sz=8.2, b=True, color=GREY, font=BODY, spc=150)]),
            para([run(b["stat"], sz=34, b=True, color=RED_DEEP, font=HEAD, spc=-20),
                  run(("  " + b["unit"]) if b["unit"] else "", sz=14, color=INK_SOFT, font=BODY)],
                 bef=2)], name="Stat"))
        s.append(text(x + EI(0.28), CT_ + EI(2.28), cw - EI(0.56), EI(0.86), [para([
            run(b["body"], sz=10.6, color=INK_SOFT, font=BODY)], line=126)], name="Body"))
        s.append(shape(x + EI(0.28), CT_ + EI(3.16), EI(0.70), EI(0.035), fill=solid(GOLD),
                       ln=line(None), name="Div"))
        s.append(text(x + EI(0.28), CT_ + EI(3.32), cw - EI(0.56), EI(0.26), [para([
            run("HOW IT WORKS", sz=8.4, b=True, color=RED, font=BODY, spc=170)])], name="HL"))
        s += steps(x + EI(0.28), CT_ + EI(3.62), cw - EI(0.56), b["flow"], col=b["col"])
    add(s, ["RPA first. The point is that KYC is a rules job, not a judgement job - read the ID, "
            "match it to a government database, open the account. Ten seconds instead of days.",
            "Chatbots second. The example to use is the 2 AM lost card: the loss does not wait "
            "for business hours, so neither should the block.",
            "Tie them together with the line in the subtitle: both remove human latency from "
            "high-volume, rule-based, time-critical work.",
            "If asked what is left for humans: exceptions, disputes and anything needing "
            "judgement or empathy."])

# ================================================================== 03 PAAS
def s03():
    reset_ids()
    s = header("Cloud foundation", "Platform as a Service (PaaS)",
               "The delivery model that lets HDFC build AI features without first building the "
               "infrastructure underneath them", pageno=3)
    lw = EI(5.35)
    s += card(M, CT_, lw, EI(2.32), radius=4000, accent=RED)
    s.append(text(M + EI(0.30), CT_ + EI(0.18), lw - EI(0.60), EI(2.05), [
        para([run("WHAT IT IS", sz=8.8, b=True, color=GOLD, font=BODY, spc=200)]),
        para([run("Platform as a Service (PaaS) is a cloud model where a vendor provides a "
                  "ready-made environment \u2014 hardware, operating systems, servers and "
                  "software tools \u2014 delivered over the internet.", sz=11.2, color=INK_SOFT,
                  font=BODY)], bef=8, line=126),
        para([run("Instead of spending time setting up physical servers, installing operating "
                  "systems or configuring databases, developers simply bring their code.",
                  sz=11.2, color=INK_SOFT, font=BODY)], bef=8, line=126)], name="Def"))
    s += card(M, CT_ + EI(2.48), lw, EI(2.40), radius=4000, accent=GOLD)
    s.append(text(M + EI(0.30), CT_ + EI(2.64), lw - EI(0.60), EI(0.3), [para([
        run("WHERE THE TIME GOES", sz=8.8, b=True, color=GOLD, font=BODY, spc=190)])], name="CH"))
    colw = (lw - EI(0.60) - EI(0.24)) // 2
    for k, (h, items, col) in enumerate([
        ("WITHOUT PaaS", ["Procure physical servers", "Install operating systems",
                          "Configure databases", "Then start building"], GREY),
        ("WITH PaaS", ["Start building"], RED)]):
        cx0 = M + EI(0.30) + k * (colw + EI(0.24))
        s.append(text(cx0, CT_ + EI(2.98), colw, EI(0.26), [para([
            run(h, sz=9.2, b=True, color=col, font=BODY, spc=110)])], name="CHd"))
        s.append(shape(cx0, CT_ + EI(3.22), colw, EI(0.028),
                       fill=solid(RED if k else GREY_LT), ln=line(None), name="CRule"))
        for i, it in enumerate(items):
            s.append(text(cx0, CT_ + EI(3.36) + EI(0.29) * i, colw, EI(0.28), [para([
                run(it, sz=9.8, b=bool(k), color=RED_DEEP if k else INK_SOFT, font=BODY)],
                line=108)], name="CIt"))
    s.append(text(M + EI(0.30) + colw + EI(0.24), CT_ + EI(3.72), colw, EI(0.9), [para([
        run("The first three steps disappear \u2014 the vendor has already done them.",
            sz=9.4, i=True, color=GREY, font=BODY)], line=116)], name="CNote"))

    rx = M + lw + EI(0.34)
    rw = CW - lw - EI(0.34)
    s.append(text(rx, CT_ - EI(0.04), rw, EI(0.28), [para([
        run("KEY CONCEPT", sz=9.2, b=True, color=RED, font=BODY, spc=180)])], name="KH"))
    s.append(text(rx, CT_ + EI(0.20), rw, EI(0.34), [para([
        run("A ready-made cloud environment", sz=15, b=True, color=RED_DEEP, font=HEAD)])],
        name="KT"))
    s += card(rx, CT_ + EI(0.66), rw, EI(2.34), fill=GOLD_PALE, border=GOLD, radius=4000)
    s.append(text(rx + EI(0.28), CT_ + EI(0.80), rw - EI(0.56), EI(0.28), [para([
        run("PROVIDED READY-MADE BY THE CLOUD VENDOR", sz=8.4, b=True, color=GOLD_DK, font=BODY,
            spc=150)])], name="SL"))
    for i, lay in enumerate(["Software tools", "Servers", "Operating system", "Hardware"]):
        s.append(shape(rx + EI(0.28), CT_ + EI(1.12) + EI(0.44) * i, rw - EI(0.56), EI(0.36),
                       geom="roundRect", adj=(12000,), fill=solid(WHITE), ln=line(GOLD, 0.8),
                       anchor="ctr", name="Layer",
                       paras=[para([run(lay, sz=10.8, b=True, color=RED_DEEP, font=BODY)],
                                   align="c")]))
    s += arrow_down(rx + rw // 2, CT_ + EI(3.12), EI(0.44), col=GOLD)
    s += card(rx, CT_ + EI(3.66), rw, EI(1.22), fill=grad(RED_DEEP, RED, 0), border=None,
              radius=4000)
    s.append(text(rx, CT_ + EI(3.98), rw, EI(0.6), [
        para([run("Developers bring their code", sz=16, b=True, color=WHITE, font=HEAD)],
             align="c"),
        para([run("and AI features go live", sz=9.6, color="E9D6C8", font=BODY)], align="c",
             bef=2)], name="Final"))
    add(s, ["Define PaaS in one sentence: the vendor hands over a ready-made environment, and you "
            "bring only your code.",
            "Use the stack on the right to make it concrete - hardware, operating system, servers "
            "and software tools all arrive pre-assembled.",
            "The 'where the time goes' comparison is the part to linger on: without PaaS the "
            "first three steps are pure setup that produces nothing a customer can see.",
            "If asked how PaaS differs from IaaS and SaaS: IaaS gives you raw machines, SaaS "
            "gives you finished software, PaaS sits in between - you write the application, the "
            "vendor runs everything below it."])

# ================================================================== 04 WHY PAAS
def s04():
    reset_ids()
    s = header("Why PaaS for HDFC Bank", "Three Benefits That Fit a Bank",
               "Speed, differentiation and data security \u2014 the three constraints a regulated "
               "bank has to satisfy at the same time", pageno=4)
    items = [
        ("01", "Fast and ready-to-use tools",
         "HDFC gets access to top-notch AI engines from providers such as Microsoft, Google or "
         "Amazon. The bank does not waste time building basic technology from zero.",
         "SOLVES:  time to market", RED),
        ("02", "Full control over custom features",
         "The bank can code and build its own custom loan algorithms and unique features \u2014 "
         "the things that set it apart from every other bank using the same cloud.",
         "SOLVES:  competitive differentiation", GOLD_DK),
        ("03", "Keeps private data safe",
         "PaaS lets the bank train its AI models inside its own private digital vault, so "
         "customer financial data never leaves the bank's control.",
         "SOLVES:  data security and compliance", RED),
    ]
    cwid = (CW - EI(0.34) * 2) // 3
    for k, (num, ttl, body, solves, col) in enumerate(items):
        x = M + k * (cwid + EI(0.34))
        s += card(x, CT_, cwid, EI(3.88), radius=4000)
        s.append(shape(x, CT_, cwid, EI(0.055), fill=solid(col), ln=line(None), name="Top"))
        s.append(shape(x + EI(0.28), CT_ + EI(0.26), EI(0.62), EI(0.62), geom="ellipse",
                       fill=solid(RED_DEEP), ln=line(GOLD, 1.1), anchor="ctr", name="Num",
                       paras=[para([run(num, sz=15, b=True, color=GOLD_LT, font=HEAD)],
                                   align="c")]))
        s.append(text(x + EI(0.28), CT_ + EI(1.02), cwid - EI(0.56), EI(0.72), [para([
            run(ttl, sz=14.6, b=True, color=RED_DEEP, font=HEAD)], line=102)], name="T"))
        s.append(shape(x + EI(0.28), CT_ + EI(1.82), EI(0.70), EI(0.035), fill=solid(GOLD),
                       ln=line(None), name="R"))
        s.append(text(x + EI(0.28), CT_ + EI(2.00), cwid - EI(0.56), EI(1.30), [para([
            run(body, sz=10.6, color=INK_SOFT, font=BODY)], line=126)], name="B"))
        s.append(pill(x + EI(0.24), CT_ + EI(3.16), cwid - EI(0.48), EI(0.46), solves,
                      fill=RED_SOFT, border=RED, color=RED, sz=9.0, spc=90))
    y = CT_ + EI(4.06)
    s += card(M, y, CW, EI(0.82), fill=grad(RED_DEEP, RED, 2700000), border=None, radius=4000)
    s.append(shape(M, y, EI(0.055), EI(0.82), fill=solid(GOLD), ln=line(None), name="B"))
    s.append(text(M + EI(0.34), y + EI(0.12), CW - EI(0.68), EI(0.66), [
        para([run("THE TRADE-OFF PAAS RESOLVES", sz=8.6, b=True, color=GOLD_LT, font=BODY,
                  spc=190)]),
        para([run("A bank normally has to choose between moving fast (buy something ready-made) "
                  "and staying in control (build it in-house). PaaS gives both \u2014 and does it "
                  "inside the bank's own security perimeter.", sz=10.6, color="EFDCDF",
                  font=BODY)], bef=4, line=118)], name="Strip"))
    add(s, ["Three benefits, one per card. Keep each to about fifteen seconds.",
            "Benefit 1 is speed - the AI engines already exist, so HDFC starts from a working "
            "platform rather than from nothing.",
            "Benefit 2 is the answer to an obvious objection: if every bank rents the same cloud, "
            "where is the advantage? In the custom layer - the bank's own loan algorithms.",
            "Benefit 3 is the one an examiner will press on, because banking is regulated. The "
            "model trains inside the bank's private environment, so customer data does not leave.",
            "The closing strip is the argument in one line: speed or control is normally a "
            "trade-off, and PaaS removes it."])

# ================================================================== 05 FRAUD FLOW
def s05():
    reset_ids()
    s = header("Pillar 3 \u2014 risk", "Adaptive Fraud Prevention",
               "Risk-based scoring instead of blanket rules \u2014 every payment is scored 0 to "
               "100 in real time and routed accordingly", pageno=5)
    lw = EI(3.86)
    s += card(M, CT_, lw, EI(2.10), radius=4000, accent=RED)
    s.append(text(M + EI(0.28), CT_ + EI(0.16), lw - EI(0.56), EI(1.85), [
        para([run("THE IDEA", sz=8.8, b=True, color=GOLD, font=BODY, spc=200)]),
        para([run("Instead of treating every transaction the same way, or randomly blocking "
                  "payments, the AI fraud engine assigns a ", sz=10.6, color=INK_SOFT, font=BODY),
              run("risk score from 0 to 100", sz=10.6, b=True, color=RED_DEEP, font=BODY),
              run(" to every payment in real time, based on how unusual it looks.", sz=10.6,
                  color=INK_SOFT, font=BODY)], bef=7, line=124)], name="Idea"))
    s += card(M, CT_ + EI(2.26), lw, EI(2.46), radius=4000, accent=GOLD)
    s.append(text(M + EI(0.28), CT_ + EI(2.42), lw - EI(0.56), EI(0.28), [para([
        run("WHY \u201CADAPTIVE\u201D MATTERS", sz=8.8, b=True, color=GOLD, font=BODY, spc=180)])],
        name="WH"))
    yy = CT_ + EI(2.76)
    for p in ["Genuine customers are not stopped for behaving normally.",
              "Extra friction is applied only where the risk justifies it.",
              "Truly suspicious payments are halted before money moves.",
              "The response is proportionate, so security does not cost convenience."]:
        s.append(shape(M + EI(0.30), yy + EI(0.04), EI(0.19), EI(0.19), geom="ellipse",
                       fill=solid(GOLD), ln=line(None), name="Dot"))
        s.append(text(M + EI(0.60), yy - EI(0.02), lw - EI(0.90), EI(0.5), [para([
            run(p, sz=10.2, color=INK_SOFT, font=BODY)], line=116)], name="P"))
        yy += EI(0.48)

    dx = M + lw + EI(0.34)
    dw = CW - lw - EI(0.34)
    dcx = dx + dw // 2

    def box(x, y, w, h, label, fill=WHITE, border=GREY, tcol=RED_DEEP, sz=10.6, bsz=0.9,
            sub=None, subcol=GREY):
        ps = [para([run(label, sz=sz, b=True, color=tcol, font=BODY)], align="c", line=104)]
        if sub:
            ps.append(para([run(sub, sz=8.8, b=True, color=subcol, font=BODY)], align="c", bef=2))
        return shape(x, y, w, h, geom="roundRect", adj=(6000,), fill=fill, ln=line(border, bsz),
                     eff=shadow(blur=10, dist=3, alpha=7.0), anchor="ctr", paras=ps, name="Box",
                     ins=(0.10, 0.05, 0.10, 0.05))

    b1w = EI(2.70)
    s.append(box(dcx - b1w // 2, CT_ + EI(0.06), b1w, EI(0.46), "Transaction Triggered",
                 fill=OFFWHITE, border=GREY_LT, tcol=INK, sz=11))
    s += arrow_down(dcx, CT_ + EI(0.58), EI(0.30), col=GREY)
    b2w = EI(3.40)
    s.append(box(dcx - b2w // 2, CT_ + EI(0.92), b2w, EI(0.52), "Real-Time AI Fraud Engine Scan",
                 fill=grad(RED_DEEP, RED, 0), border=RED_DEEP, tcol=WHITE, sz=11.6, bsz=1.0))
    s.append(vline(dcx, CT_ + EI(1.50), EI(0.22), col=GREY))
    bw_ = EI(2.28); gap_ = EI(0.34)
    total = bw_ * 3 + gap_ * 2
    x0 = dx + (dw - total) // 2
    centers = [x0 + bw_ // 2 + k * (bw_ + gap_) for k in range(3)]
    s.append(hline(centers[0], CT_ + EI(1.72), centers[2] - centers[0], col=GREY))
    tiers = [("LOW RISK", "Score 0 \u2013 30", GREEN, "Approve instant transaction"),
             ("MEDIUM RISK", "Score 31 \u2013 70", GOLD_DK,
              "Adaptive step-up\n(biometric / OTP)"),
             ("HIGH RISK", "Score 71 \u2013 100", RED, "Temporary block\n& manual review")]
    for k, (lab, score, col, outcome) in enumerate(tiers):
        cxk = centers[k]
        s += arrow_down(cxk, CT_ + EI(1.72), EI(0.34), col=col)
        s.append(box(cxk - bw_ // 2, CT_ + EI(2.10), bw_, EI(0.72), lab, fill=col, border=col,
                     tcol=WHITE, sz=11.4, bsz=1.0, sub=score, subcol=WHITE))
        s += arrow_down(cxk, CT_ + EI(2.88), EI(0.32), col=col)
        first, _, second = outcome.partition("\n")
        ps = [para([run(first, sz=10.6, b=True, color=col, font=BODY)], align="c", line=106)]
        if second:
            ps.append(para([run(second, sz=10.6, b=True, color=col, font=BODY)], align="c",
                           line=106))
        s.append(shape(cxk - bw_ // 2, CT_ + EI(3.24), bw_, EI(0.78), geom="roundRect",
                       adj=(6000,), fill=solid(WHITE), ln=line(col, 1.1),
                       eff=shadow(blur=10, dist=3, alpha=7.0), anchor="ctr", paras=ps,
                       name="Outcome", ins=(0.10, 0.05, 0.10, 0.05)))
    sy = CT_ + EI(4.22)
    s.append(text(x0, sy, total, EI(0.26), [para([
        run("RISK SCORE SCALE", sz=8.4, b=True, color=GREY, font=BODY, spc=160)])], name="SS"))
    sxp = x0
    for wgt, col in ((30, GREEN), (40, GOLD), (30, RED)):
        wseg = int(total * wgt / 100.0)
        s.append(shape(sxp, sy + EI(0.28), wseg, EI(0.20), fill=solid(col), ln=line(None),
                       name="Seg"))
        sxp += wseg
    for frac, lab in ((0.0, "0"), (0.30, "30"), (0.70, "70"), (1.0, "100")):
        s.append(text(x0 + int(total * frac) - EI(0.20), sy + EI(0.52), EI(0.40), EI(0.24),
                      [para([run(lab, sz=8.4, b=True, color=GREY, font=BODY)], align="c")],
                      name="Tick"))
    add(s, ["This is the diagram slide. Walk it top to bottom, then left to right.",
            "Top: a transaction is triggered and the AI engine scans it in real time. Nothing is "
            "pre-judged.",
            "The engine outputs a single number - a risk score from 0 to 100 - based on how "
            "unusual the payment looks for that customer.",
            "Then the three paths. Low risk 0 to 30 is approved instantly. Medium risk 31 to 70 "
            "gets a step-up check, biometric or OTP. High risk 71 to 100 is temporarily blocked "
            "and reviewed by a human.",
            "The word to emphasise is 'adaptive' - the friction matches the risk, so a normal "
            "customer never feels the security.",
            "Likely question: what happens at the boundaries? The thresholds are tunable, and "
            "banks tighten or loosen them as fraud patterns shift."])

# ================================================================== 06 CONCLUSION
def s06():
    reset_ids()
    s = header("In summary", "Conclusion",
               "HDFC Bank uses AI-driven solutions to make digital banking faster, safer and "
               "more convenient", pageno=6)
    items = [("RPA", "Automates KYC verification, opening an account in seconds.", RED),
             ("Chatbots", "Provides 24/7 customer support with no human agent needed.", GOLD_DK),
             ("PaaS", "Provides a ready-to-use environment for building and deploying AI.", RED),
             ("Adaptive fraud prevention",
              "Uses real-time risk scores to identify and respond to suspicious transactions.",
              GOLD_DK)]
    cwid = (CW - EI(0.30) * 3) // 4
    for k, (ttl, body, col) in enumerate(items):
        x = M + k * (cwid + EI(0.30))
        s += card(x, CT_, cwid, EI(2.06), radius=4000)
        s.append(shape(x, CT_, cwid, EI(0.05), fill=solid(col), ln=line(None), name="Top"))
        s.append(shape(x + EI(0.24), CT_ + EI(0.24), EI(0.42), EI(0.42), geom="ellipse",
                       fill=solid(GOLD_PALE), ln=line(GOLD, 0.8), anchor="ctr", name="Dot",
                       paras=[para([run("0%d" % (k + 1), sz=9.6, b=True, color=GOLD_DK,
                                        font=HEAD)], align="c")]))
        s.append(text(x + EI(0.24), CT_ + EI(0.78), cwid - EI(0.48), EI(0.52), [para([
            run(ttl, sz=13.2, b=True, color=RED_DEEP, font=HEAD)], line=102)], name="T"))
        s.append(text(x + EI(0.24), CT_ + EI(1.34), cwid - EI(0.48), EI(0.62), [para([
            run(body, sz=9.8, color=INK_SOFT, font=BODY)], line=120)], name="B"))
    pw = (CW - EI(0.30) * 2) // 3
    for k, (a, b_) in enumerate([("FASTER", "seconds, not days"),
                                 ("SAFER", "risk-scored in real time"),
                                 ("MORE CONVENIENT", "any hour, any channel")]):
        s.append(shape(M + k * (pw + EI(0.30)), CT_ + EI(2.26), pw, EI(0.62), geom="roundRect",
                       adj=(20000,), fill=solid(GOLD_PALE), ln=line(GOLD, 0.8), anchor="ctr",
                       name="Out", paras=[para([
            run(a, sz=11.6, b=True, color=RED_DEEP, font=HEAD, spc=80),
            run("   \u2014   " + b_, sz=9.8, color=GOLD_DK, font=BODY)], align="c")]))
    y = CT_ + EI(3.18)
    s += card(M, y, CW, EI(1.54), fill=grad(RED_DEEP, "3C060E", 2700000, c3=RED), border=None,
              radius=4000)
    s.append(shape(M, y, EI(0.06), EI(1.54), fill=solid(GOLD), ln=line(None), name="Bar"))
    s.append(text(M + EI(0.40), y + EI(0.34), EI(4.4), EI(0.9), [
        para([run("THANK YOU", sz=32, b=True, color=WHITE, font=HEAD, spc=60)]),
        para([run("Questions are welcome.", sz=10.4, color=GOLD_LT, font=BODY)], bef=4)],
        name="TY"))
    bwid = EI(2.16)
    for k, (lab, val) in enumerate([("PRESENTED BY", NAME), ("REGISTER NO", REGNO),
                                    ("PROGRAMME", "MBA \u2014 Finance")]):
        x = SW - M - EI(0.34) - (3 - k) * (bwid + EI(0.18)) + EI(0.18)
        s.append(shape(x, y + EI(0.42), bwid, EI(0.70), geom="roundRect", adj=(8000,),
                       fill=solid(WHITE, 10.0), ln=line(GOLD_LT, 0.8, alpha=45.0), name="Det",
                       paras=[
            para([run(lab, sz=7.6, b=True, color=GOLD_LT, font=BODY, spc=150)]),
            para([run(val, sz=11.2, b=True, color=WHITE, font=HEAD)], bef=3)],
                       ins=(0.18, 0.11, 0.12, 0.08)))
    add(s, ["Recap the four capabilities in one line each, then land the three outcomes: faster, "
            "safer, more convenient.",
            "The connecting idea, if you want a strong closing sentence: RPA and chatbots change "
            "the speed of banking, PaaS changes how quickly the bank can build, and adaptive "
            "fraud prevention changes how safely it can grow.",
            "Then stop and invite questions.",
            "Prepared answers: PaaS versus IaaS and SaaS; what happens to staff whose work is "
            "automated (they move to exceptions and advisory work); and how the risk thresholds "
            "are set (tuned from historical fraud data and adjusted as patterns change)."])

for fn in (s01, s02, s03, s04, s05, s06):
    fn()

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "Finance", "HDFC_AI_Digital_Banking_CaseStudy2(Prahadhesvaryaa K S).pptx")
write_pptx(OUT, SLIDES, NOTES,
           "Case Study 2 - HDFC Bank | AI-Driven Digital Banking Transformation", NAME)
print("Slides:", len(SLIDES))
print("Written:", OUT, os.path.getsize(OUT), "bytes")
