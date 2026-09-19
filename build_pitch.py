#!/usr/bin/env python3
"""Generate the ShakeStop gym pitch deck (16:9)."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- Brand palette (ShakeStop 2.0) ----
INK      = RGBColor(0x0B, 0x12, 0x0C)   # green-tinted near-black
INK2     = RGBColor(0x11, 0x1A, 0x12)
CARD     = RGBColor(0x18, 0x24, 0x18)
GREEN    = RGBColor(0x5C, 0xFF, 0x33)   # main / dominant
LIME     = RGBColor(0xAF, 0xFC, 0x41)   # 2nd main / complimentary
PURPLE   = RGBColor(0x97, 0x92, 0xE3)   # highlighter / contrast
TEAL     = GREEN                         # alias -> primary accent
TEAL_DK  = PURPLE                        # alias -> contrast fills
PINK     = PURPLE                        # alias
AMBER    = LIME                          # alias
WHITE    = RGBColor(0xEB, 0xFF, 0xE8)   # lightest neutral
MUTED    = RGBColor(0x9A, 0xB0, 0x9A)
CREAM    = RGBColor(0xDE, 0xFB, 0xD3)   # neutral

EMU_W, EMU_H = Inches(13.333), Inches(7.5)
FONT = "Segoe UI"

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


def slide(bg=INK):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, EMU_W, EMU_H)
    r.fill.solid(); r.fill.fore_color.rgb = bg; r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)
    return s


def box(s, x, y, w, h, fill=None, line=None, line_w=1.0, radius=False, shadow=False):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    return shp


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=6, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph is list of (txt,size,color,bold)."""
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(space_after); p.line_spacing = line_spacing
        for (t, sz, col, bold) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col; r.font.bold = bold
            r.font.name = FONT
    return tb


def logo(s, cx, cy, size):
    """Draw octagon badge + shaker cup + straw, centered at (cx,cy)."""
    half = size / 2
    oct_ = s.shapes.add_shape(MSO_SHAPE.OCTAGON, int(cx - half), int(cy - half), int(size), int(size))
    oct_.fill.solid(); oct_.fill.fore_color.rgb = TEAL; oct_.line.fill.background()
    oct_.shadow.inherit = False
    # cup body (rounded rect)
    bw, bh = size * 0.34, size * 0.42
    cup = box(s, int(cx - bw/2), int(cy - bh*0.30), int(bw), int(bh), fill=WHITE, radius=True)
    # lid
    lw, lh = size * 0.42, size * 0.11
    box(s, int(cx - lw/2), int(cy - bh*0.30 - lh*0.9), int(lw), int(lh), fill=WHITE, radius=True)
    # teal band on cup
    band = box(s, int(cx - bw/2), int(cy + bh*0.03), int(bw), int(bh*0.16), fill=GREEN)
    # straw
    straw = s.shapes.add_connector(2, int(cx + size*0.10), int(cy - bh*0.30 - lh*0.9),
                                   int(cx + size*0.26), int(cy - bh*0.72))
    straw.line.color.rgb = PINK; straw.line.width = Pt(max(2, size/EMU_per_pt/6))
    return oct_


# helper: emu per point for line width scaling
EMU_per_pt = 12700

def wordmark(s, x, y, size=40):
    tb = s.shapes.add_textbox(x, y, Inches(6), Inches(1))
    tf = tb.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = "Shake"; r1.font.size = Pt(size); r1.font.bold = True
    r1.font.color.rgb = WHITE; r1.font.name = FONT
    r2 = p.add_run(); r2.text = "Stop"; r2.font.size = Pt(size); r2.font.bold = True
    r2.font.color.rgb = TEAL; r2.font.name = FONT
    return tb


def chip(s, x, y, w, h, label, num_color=TEAL):
    box(s, x, y, w, h, fill=CARD, radius=True)
    return box


def kicker(s, x, y, txt, color=TEAL):
    text(s, x, y, Inches(8), Inches(0.4),
         [[(txt.upper(), 13, color, True)]])


def pagefoot(s, n):
    text(s, Inches(0.5), Inches(7.05), Inches(6), Inches(0.35),
         [[("ShakeStop  ·  Gym Partnership Proposal", 9, MUTED, False)]])
    text(s, Inches(11.6), Inches(7.05), Inches(1.2), Inches(0.35),
         [[(f"{n:02d}", 9, MUTED, True)]], align=PP_ALIGN.RIGHT)


# =========================================================
# SLIDE 1 — Title
# =========================================================
s = slide(INK)
box(s, 0, 0, EMU_W, Inches(0.18), fill=TEAL)
logo(s, Inches(3.6), Inches(2.7), Inches(1.7))
wordmark(s, Inches(4.7), Inches(2.35), size=54)
text(s, Inches(4.75), Inches(3.5), Inches(7), Inches(0.5),
     [[("FRESH PROTEIN · ANYTIME", 15, MUTED, True)]])
text(s, Inches(1.0), Inches(4.6), Inches(11.3), Inches(1.2),
     [[("Fresh protein shakes, right inside your gym.", 26, WHITE, True)]],
     align=PP_ALIGN.CENTER)
text(s, Inches(1.5), Inches(5.45), Inches(10.3), Inches(0.8),
     [[("A zero-effort, revenue-sharing partnership proposal", 16, MUTED, False)]],
     align=PP_ALIGN.CENTER)
box(s, Inches(4.5), Inches(6.25), Inches(4.3), Inches(0.6), fill=None, line=TEAL, line_w=1.25, radius=True)
text(s, Inches(4.5), Inches(6.32), Inches(4.3), Inches(0.5),
     [[("Prepared for  [ Gym Name ]", 13, TEAL, True)]], align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 2 — The Problem
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "The problem")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("After a workout, protein is a hassle", 30, WHITE, True)]])
text(s, Inches(0.6), Inches(1.75), Inches(12), Inches(0.6),
     [[("Members finish training wanting protein — but every option today has friction.", 15, MUTED, False)]])
prob = [
    ("🥤", "Make it themselves", "Carrying a shaker, powder and cleaning up is inconvenient."),
    ("📦", "Buy a packaged drink", "Expensive, full of preservatives, and not freshly made."),
    ("🚶", "Leave the gym", "They walk out to find a shake — losing time and spending elsewhere."),
]
cx = Inches(0.6); cw = Inches(3.95); gap = Inches(0.2); cy = Inches(2.7); ch = Inches(3.2)
for i, (emo, t, d) in enumerate(prob):
    x = Emu(int(cx) + i * (int(cw) + int(gap)))
    box(s, x, cy, cw, ch, fill=CARD, radius=True)
    text(s, x, Inches(3.05), cw, Inches(0.9), [[(emo, 40, WHITE, False)]], align=PP_ALIGN.CENTER)
    text(s, x, Inches(4.05), cw, Inches(0.6), [[(t, 18, TEAL, True)]], align=PP_ALIGN.CENTER)
    text(s, Emu(int(x)+Inches(0.25)), Inches(4.7), Emu(int(cw)-Inches(0.5)), Inches(1.2),
         [[(d, 13, MUTED, False)]], align=PP_ALIGN.CENTER)
text(s, Inches(0.6), Inches(6.2), Inches(12), Inches(0.6),
     [[("The result: an unmet member need — and revenue walking out the door.", 15, PINK, True)]])
pagefoot(s, 2)

# =========================================================
# SLIDE 3 — The Solution
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "The solution")
text(s, Inches(0.6), Inches(0.85), Inches(12.2), Inches(0.9),
     [[("A protein-shake vending machine — inside your gym", 28, WHITE, True)]])
text(s, Inches(0.6), Inches(1.95), Inches(7.4), Inches(3.6),
     [
        [("Members select a shake, pay digitally, and collect a ", 16, WHITE, False),
         ("freshly prepared", 16, TEAL, True),
         (" protein shake in minutes — without leaving the floor.", 16, WHITE, False)],
        [("", 8, MUTED, False)],
        [("• Freshly blended on demand — not a packaged drink", 15, MUTED, False)],
        [("• Multiple flavours + custom shakes, milk or water base", 15, MUTED, False)],
        [("• 100% digital payment (UPI / card / wallet)", 15, MUTED, False)],
        [("• Hygienic, sealed ingredient canisters", 15, MUTED, False)],
        [("• Compact footprint — fits in a corner", 15, MUTED, False)],
     ], line_spacing=1.15, space_after=8)
# machine mock on right
mx, my, mw, mh = Inches(8.5), Inches(1.9), Inches(3.6), Inches(4.9)
box(s, mx, my, mw, mh, fill=INK2, line=RGBColor(0x2B,0x34,0x44), line_w=1.25, radius=True)
box(s, Emu(int(mx)+Inches(0.3)), Inches(2.2), Emu(int(mw)-Inches(0.6)), Inches(0.55), fill=GREEN, radius=True)
text(s, Emu(int(mx)+Inches(0.3)), Inches(2.27), Emu(int(mw)-Inches(0.6)), Inches(0.45),
     [[("ShakeStop", 16, INK, True)]], align=PP_ALIGN.CENTER)
box(s, Emu(int(mx)+Inches(0.3)), Inches(2.95), Emu(int(mw)-Inches(0.6)), Inches(0.5), fill=RGBColor(0x22,0x14,0x40), radius=True)
text(s, Emu(int(mx)+Inches(0.3)), Inches(3.0), Emu(int(mw)-Inches(0.6)), Inches(0.4),
     [[("AD · Partner brands", 9, MUTED, True)]], align=PP_ALIGN.CENTER)
box(s, Emu(int(mx)+Inches(0.3)), Inches(3.6), Emu(int(mw)-Inches(0.6)), Inches(1.7), fill=RGBColor(0x0A,0x10,0x16), radius=True)
text(s, Emu(int(mx)+Inches(0.3)), Inches(3.7), Emu(int(mw)-Inches(0.6)), Inches(0.4),
     [[("TOUCH TO ORDER", 10, TEAL, True)]], align=PP_ALIGN.CENTER)
text(s, Emu(int(mx)+Inches(0.3)), Inches(4.05), Emu(int(mw)-Inches(0.6)), Inches(1.2),
     [[("🍫  🍓  🍦  ☕  ⚡", 22, WHITE, False)]], align=PP_ALIGN.CENTER)
box(s, Emu(int(mx)+Inches(0.3)), Inches(5.45), Emu(int(mw)-Inches(0.6)), Inches(1.05), fill=RGBColor(0x17,0x11,0x0A), line=AMBER, line_w=1.0, radius=True)
text(s, Emu(int(mx)+Inches(0.3)), Inches(5.85), Emu(int(mw)-Inches(0.6)), Inches(0.5),
     [[("🥤  Collect your shake", 13, AMBER, True)]], align=PP_ALIGN.CENTER)
pagefoot(s, 3)

# =========================================================
# SLIDE 4 — How it works
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "How it works")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("From tap to sip in under 2 minutes", 28, WHITE, True)]])
steps = [
    ("1", "Select", "Member picks a flavour or builds a custom shake on the touchscreen."),
    ("2", "Pay", "Instant digital payment — UPI, card or wallet. No cash, no queue."),
    ("3", "Blend", "The machine freshly blends the shake from sealed ingredients."),
    ("4", "Collect", "A fresh protein shake is dispensed in minutes. Done."),
]
cw = Inches(2.9); gap = Inches(0.25); x0 = Inches(0.6); cy = Inches(2.5); ch = Inches(3.0)
for i, (n, t, d) in enumerate(steps):
    x = Emu(int(x0) + i*(int(cw)+int(gap)))
    box(s, x, cy, cw, ch, fill=CARD, radius=True)
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(x)+Inches(0.3)), Inches(2.8), Inches(0.75), Inches(0.75))
    circ.fill.solid(); circ.fill.fore_color.rgb = TEAL; circ.line.fill.background(); circ.shadow.inherit=False
    tf = circ.text_frame; tf.word_wrap=False; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    rr=p.add_run(); rr.text=n; rr.font.size=Pt(24); rr.font.bold=True; rr.font.color.rgb=INK; rr.font.name=FONT
    text(s, Emu(int(x)+Inches(0.25)), Inches(3.75), Emu(int(cw)-Inches(0.5)), Inches(0.5),
         [[(t, 19, WHITE, True)]])
    text(s, Emu(int(x)+Inches(0.25)), Inches(4.25), Emu(int(cw)-Inches(0.5)), Inches(1.3),
         [[(d, 13, MUTED, False)]], line_spacing=1.15)
    if i < 3:
        text(s, Emu(int(x)+int(cw)-Inches(0.05)), Inches(3.7), Inches(0.4), Inches(0.6),
             [[("›", 30, TEAL, True)]], align=PP_ALIGN.CENTER)
pagefoot(s, 4)

# =========================================================
# SLIDE 5 — The Machine
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "The machine")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("Smart, hygienic, and built for gyms", 28, WHITE, True)]])
feats = [
    ("🖥️", "Two displays", "One for ordering, one for ads & promotions."),
    ("🧴", "8 sealed canisters", "3 protein flavours + milk, water, coffee, pre-workout, creatine."),
    ("🥛", "Fresh, custom shakes", "Choose base (milk/water), add boosts, or build your own."),
    ("💳", "100% digital payments", "UPI / card / wallet — fully cashless & tracked."),
    ("🧼", "Hygienic by design", "Sealed ingredients, auto-dosing, easy-clean dispensing bay."),
    ("📶", "Smart & connected", "Remote monitoring of stock, sales and health."),
]
cw = Inches(3.95); ch = Inches(1.55); gx = Inches(0.2); gy = Inches(0.25)
x0 = Inches(0.6); y0 = Inches(2.15)
for i, (emo, t, d) in enumerate(feats):
    r, c = divmod(i, 3)
    x = Emu(int(x0) + c*(int(cw)+int(gx)))
    y = Emu(int(y0) + r*(int(ch)+int(gy)))
    box(s, x, y, cw, ch, fill=CARD, radius=True)
    text(s, Emu(int(x)+Inches(0.25)), Emu(int(y)+Inches(0.2)), Inches(0.8), Inches(0.7),
         [[(emo, 26, WHITE, False)]])
    text(s, Emu(int(x)+Inches(1.05)), Emu(int(y)+Inches(0.18)), Emu(int(cw)-Inches(1.2)), Inches(0.45),
         [[(t, 15, TEAL, True)]])
    text(s, Emu(int(x)+Inches(1.05)), Emu(int(y)+Inches(0.62)), Emu(int(cw)-Inches(1.2)), Inches(0.85),
         [[(d, 11.5, MUTED, False)]], line_spacing=1.1)
text(s, Inches(0.6), Inches(5.75), Inches(12), Inches(0.5),
     [[("See the live interactive prototype:  ", 13, MUTED, False),
       ("arpitsing.github.io/shakestop", 13, TEAL, True)]])
pagefoot(s, 5)

# =========================================================
# SLIDE 6 — Zero effort for the gym
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "Effort split")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("Virtually zero work for your team", 28, WHITE, True)]])
# We handle
box(s, Inches(0.6), Inches(2.1), Inches(6.05), Inches(4.4), fill=CARD, radius=True)
text(s, Inches(0.9), Inches(2.35), Inches(5.5), Inches(0.5), [[("✅  We handle everything", 18, TEAL, True)]])
we = ["Machine supply & installation", "All ingredients (authentic whey)",
      "Refilling & stock management", "Daily cleaning & sanitisation",
      "Servicing, repairs & uptime", "Payments, pricing & support",
      "Menu, branding & promotions"]
text(s, Inches(0.95), Inches(3.0), Inches(5.4), Inches(3.4),
     [[("• "+w, 14, WHITE, False)] for w in we], line_spacing=1.2, space_after=7)
# Gym provides
box(s, Inches(6.85), Inches(2.1), Inches(5.85), Inches(4.4), fill=INK2, line=RGBColor(0x2B,0x34,0x44), line_w=1.0, radius=True)
text(s, Inches(7.15), Inches(2.35), Inches(5.3), Inches(0.5), [[("🤝  Gym simply provides", 18, PINK, True)]])
gym = ["A small floor space (approx. 1 m²)", "A standard power socket",
       "Wi-Fi (optional, for live monitoring)"]
text(s, Inches(7.2), Inches(3.0), Inches(5.2), Inches(2.0),
     [[("• "+g, 14, WHITE, False)] for g in gym], line_spacing=1.25, space_after=9)
text(s, Inches(7.2), Inches(5.05), Inches(5.2), Inches(1.3),
     [[("No café to set up. No staff to hire. No inventory to manage. "
        "No capital investment.", 14, MUTED, False)]], line_spacing=1.2)
pagefoot(s, 6)

# =========================================================
# SLIDE 7 — Revenue share model
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "Revenue model")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("You earn ", 28, WHITE, True), ("4% of every sale", 28, TEAL, True),
       (" — as pure passive income", 28, WHITE, True)]])
text(s, Inches(0.6), Inches(1.8), Inches(12), Inches(0.6),
     [[("No investment, no cost, no risk. We install and run it; you earn a share of gross sales, paid monthly.", 14, MUTED, False)]])
# Big 5% badge
oval = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.9), Inches(2.9), Inches(2.6), Inches(2.6))
oval.fill.solid(); oval.fill.fore_color.rgb = TEAL; oval.line.fill.background(); oval.shadow.inherit=False
tf=oval.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
rr=p.add_run(); rr.text="4%"; rr.font.size=Pt(60); rr.font.bold=True; rr.font.color.rgb=INK; rr.font.name=FONT
p2=tf.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
r2=p2.add_run(); r2.text="of gross sales to the gym"; r2.font.size=Pt(13); r2.font.bold=True; r2.font.color.rgb=INK; r2.font.name=FONT
# Illustrative table
tx, ty, tw = Inches(4.1), Inches(2.75), Inches(8.6)
text(s, tx, Inches(2.4), tw, Inches(0.4), [[("Illustrative monthly earnings (example only)", 12, MUTED, True)]])
rows = [
  ("Scenario", "Shakes / day", "Avg price", "Monthly sales", "Gym 4% / month"),
  ("Conservative", "30", "₹120", "₹1,08,000", "₹4,320"),
  ("Moderate", "50", "₹120", "₹1,80,000", "₹7,200"),
  ("Busy gym", "80", "₹120", "₹2,88,000", "₹11,520"),
]
rh = Inches(0.78); cw5 = [Inches(1.9), Inches(1.6), Inches(1.4), Inches(1.9), Inches(1.8)]
yy = ty
for ri, row in enumerate(rows):
    xx = tx
    header = ri == 0
    for ci, val in enumerate(row):
        fill = TEAL_DK if header else (INK2 if ri % 2 else CARD)
        box(s, xx, yy, cw5[ci], rh, fill=fill, line=INK, line_w=0.75)
        col = WHITE if header else (TEAL if ci == 4 and not header else WHITE)
        text(s, xx, Emu(int(yy)+Inches(0.16)), cw5[ci], Inches(0.5),
             [[(val, 12 if header else 12.5, col, header or ci==4 or ci==0)]],
             align=PP_ALIGN.CENTER)
        xx = Emu(int(xx)+int(cw5[ci]))
    yy = Emu(int(yy)+int(rh))
text(s, tx, Emu(int(yy)+Inches(0.05)), tw, Inches(0.5),
     [[("Figures are illustrative estimates; actuals depend on footfall & menu. Paid monthly with a transparent sales dashboard.", 10.5, MUTED, False)]])
pagefoot(s, 7)

# =========================================================
# SLIDE 8 — Why it's a win for the gym
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "Why partner with us")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("A win for the gym, and for members", 28, WHITE, True)]])
wins = [
    ("💰", "New passive income", "A fresh revenue stream with zero investment or effort."),
    ("⭐", "Better member experience", "A premium, modern amenity members love — boosts retention."),
    ("🏆", "Competitive edge", "Stand out from nearby gyms with an in-house shake bar."),
    ("🧾", "No overheads", "No staff, no inventory, no maintenance costs on you."),
    ("📊", "Full transparency", "Live dashboard of sales and your monthly share."),
    ("🔄", "Flexible & risk-free", "Simple agreement; we can remove the machine anytime."),
]
cw = Inches(3.95); ch = Inches(1.6); gx = Inches(0.2); gy = Inches(0.28)
x0 = Inches(0.6); y0 = Inches(2.15)
for i, (emo, t, d) in enumerate(wins):
    r, c = divmod(i, 3)
    x = Emu(int(x0)+c*(int(cw)+int(gx))); y = Emu(int(y0)+r*(int(ch)+int(gy)))
    box(s, x, y, cw, ch, fill=CARD, radius=True)
    text(s, Emu(int(x)+Inches(0.25)), Emu(int(y)+Inches(0.22)), Inches(0.8), Inches(0.7),
         [[(emo, 24, WHITE, False)]])
    text(s, Emu(int(x)+Inches(1.0)), Emu(int(y)+Inches(0.2)), Emu(int(cw)-Inches(1.2)), Inches(0.5),
         [[(t, 14.5, TEAL, True)]])
    text(s, Emu(int(x)+Inches(1.0)), Emu(int(y)+Inches(0.68)), Emu(int(cw)-Inches(1.2)), Inches(0.85),
         [[(d, 11.5, MUTED, False)]], line_spacing=1.1)
pagefoot(s, 8)

# =========================================================
# SLIDE 9 — Rollout / Next steps
# =========================================================
s = slide(INK)
kicker(s, Inches(0.6), Inches(0.5), "Next steps")
text(s, Inches(0.6), Inches(0.85), Inches(12), Inches(0.9),
     [[("Getting started is simple", 28, WHITE, True)]])
text(s, Inches(0.6), Inches(1.8), Inches(12), Inches(0.6),
     [[("A smooth, low-effort setup — we handle the heavy lifting.", 15, MUTED, False)]])
steps = [
    ("1", "Site visit", "We pick the best spot & confirm power/space."),
    ("2", "Install", "Machine delivered, installed & stocked — free."),
    ("3", "Go live", "Members start ordering; you start earning 4%."),
    ("4", "Review", "Gather feedback and enhance the menu & experience."),
]
cw = Inches(2.9); gap = Inches(0.25); x0 = Inches(0.6); cy = Inches(2.7); ch = Inches(2.6)
for i, (n, t, d) in enumerate(steps):
    x = Emu(int(x0)+i*(int(cw)+int(gap)))
    box(s, x, cy, cw, ch, fill=CARD, radius=True)
    box(s, x, cy, cw, Inches(0.12), fill=TEAL, radius=False)
    text(s, Emu(int(x)+Inches(0.25)), Emu(int(cy)+Inches(0.3)), Inches(1), Inches(0.6),
         [[("Step "+n, 12, TEAL, True)]])
    text(s, Emu(int(x)+Inches(0.25)), Emu(int(cy)+Inches(0.75)), Emu(int(cw)-Inches(0.5)), Inches(0.5),
         [[(t, 18, WHITE, True)]])
    text(s, Emu(int(x)+Inches(0.25)), Emu(int(cy)+Inches(1.3)), Emu(int(cw)-Inches(0.5)), Inches(1.1),
         [[(d, 13, MUTED, False)]], line_spacing=1.15)
pagefoot(s, 9)

# =========================================================
# SLIDE 10 — Contact / Thank you
# =========================================================
s = slide(INK)
box(s, 0, 0, EMU_W, Inches(0.18), fill=TEAL)
logo(s, Inches(6.66), Inches(2.3), Inches(1.7))
text(s, Inches(1), Inches(3.3), Inches(11.3), Inches(0.9),
     [[("Shake", 40, WHITE, True), ("Stop", 40, TEAL, True)]], align=PP_ALIGN.CENTER)
text(s, Inches(1), Inches(4.15), Inches(11.3), Inches(0.6),
     [[("Fresh protein for your members. Passive income for your gym.", 17, MUTED, False)]],
     align=PP_ALIGN.CENTER)
text(s, Inches(1), Inches(5.2), Inches(11.3), Inches(1.2),
     [
        [("Let's talk", 18, TEAL, True)],
        [("Riya Singla   ·   9779337982   ·   singlariya2406@gmail.com", 14, WHITE, False)],
        [("Live demo:  arpitsing.github.io/shakestop", 13, MUTED, False)],
     ], align=PP_ALIGN.CENTER, space_after=8)
pagefoot(s, 10)

out = "ShakeStop-Gym-Pitch.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
