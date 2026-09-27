#!/usr/bin/env python3
"""
TaskPilot — AI Hackathon Deck Generator
Built with python-pptx. Dark premium theme, cyan/violet accents.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.enum.text import MSO_AUTO_SIZE
import copy

# ----------------------------------------------------------------------------
# PALETTE
# ----------------------------------------------------------------------------
BG          = RGBColor(0x0A, 0x0D, 0x13)   # near-black charcoal
BG2         = RGBColor(0x0D, 0x11, 0x18)   # slightly lifted charcoal (for panels bg)
CARD        = RGBColor(0x12, 0x17, 0x20)   # glass panel fill
CARD_LIGHT  = RGBColor(0x16, 0x1D, 0x28)
BORDER      = RGBColor(0x24, 0x2E, 0x3B)   # thin technical border
BORDER_SOFT = RGBColor(0x1A, 0x22, 0x2C)
CYAN        = RGBColor(0x35, 0xD6, 0xF5)   # primary accent — electric cyan
CYAN_DIM    = RGBColor(0x1B, 0x6E, 0x84)
VIOLET      = RGBColor(0x9B, 0x6B, 0xF2)   # secondary accent — violet
WHITE       = RGBColor(0xF2, 0xF5, 0xF9)
MUTED       = RGBColor(0x8B, 0x96, 0xA6)   # muted grey text
MUTED_DIM   = RGBColor(0x5A, 0x64, 0x72)
GREEN       = RGBColor(0x4A, 0xDE, 0x80)
AMBER       = RGBColor(0xF5, 0xA6, 0x23)

FONT = "Arial"
FONT_HEAD = "Arial"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# LOW-LEVEL HELPERS
# ----------------------------------------------------------------------------

def new_slide():
    return prs.slides.add_slide(BLANK)


def set_alpha(shape, pct_opaque):
    """Set fill transparency. pct_opaque: 0-100 (100 = fully opaque)."""
    alpha_val = str(int(pct_opaque * 1000))
    sp = shape.fill._xPr.find(qn('a:solidFill'))
    if sp is None:
        return
    srgb = sp.find(qn('a:srgbClr'))
    if srgb is None:
        return
    for tag in ('a:alpha',):
        existing = srgb.find(qn(tag))
        if existing is not None:
            srgb.remove(existing)
    alpha = srgb.makeelement(qn('a:alpha'), {'val': alpha_val})
    srgb.append(alpha)


def set_line_alpha(shape, pct_opaque):
    alpha_val = str(int(pct_opaque * 1000))
    ln = shape.line._get_or_add_ln()
    sf = ln.find(qn('a:solidFill'))
    if sf is None:
        return
    srgb = sf.find(qn('a:srgbClr'))
    if srgb is None:
        return
    existing = srgb.find(qn('a:alpha'))
    if existing is not None:
        srgb.remove(existing)
    alpha = srgb.makeelement(qn('a:alpha'), {'val': alpha_val})
    srgb.append(alpha)


def no_shadow(shape):
    spPr = shape._element.spPr
    existing = spPr.find(qn('a:effectLst'))
    if existing is not None:
        spPr.remove(existing)
    effectLst = spPr.makeelement(qn('a:effectLst'), {})
    spPr.append(effectLst)


def soft_glow(shape, color, radius_pt=18, alpha_pct=45):
    """Adds a soft outer glow effect to a shape."""
    spPr = shape._element.spPr
    existing = spPr.find(qn('a:effectLst'))
    if existing is not None:
        spPr.remove(existing)
    effectLst = spPr.makeelement(qn('a:effectLst'), {})
    glow = effectLst.makeelement(qn('a:glow'), {'rad': str(Pt(radius_pt))})
    clr = glow.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (color[0], color[1], color[2])})
    alpha = clr.makeelement(qn('a:alpha'), {'val': str(alpha_pct * 1000)})
    clr.append(alpha)
    glow.append(clr)
    effectLst.append(glow)
    spPr.append(effectLst)


def bg_fill(slide, color=BG):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    rect.fill.solid()
    rect.fill.fore_color.rgb = color
    rect.line.fill.background()
    rect.shadow.inherit = False
    rect._element.spPr.append(rect._element.spPr.makeelement(qn('a:effectLst'), {}))
    # send to back
    sp = rect._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)
    return rect


def grid_lines(slide, color=BORDER_SOFT, n_v=6, n_h=4, alpha=100):
    """Subtle technical grid — thin faint lines across the slide."""
    for i in range(1, n_v):
        x = Emu(int(SLIDE_W * i / n_v))
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, 0, x, SLIDE_H)
        ln.line.color.rgb = color
        ln.line.width = Pt(0.4)
        set_line_alpha(ln, alpha)
    for j in range(1, n_h):
        y = Emu(int(SLIDE_H * j / n_h))
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, y, SLIDE_W, y)
        ln.line.color.rgb = color
        ln.line.width = Pt(0.4)
        set_line_alpha(ln, alpha)


def add_text(slide, x, y, w, h, text, size=14, color=WHITE, bold=False,
             italic=False, align=PP_ALIGN.LEFT, font=FONT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.0, letter_spacing=None, wrap=True, shrink=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if shrink:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.name = font
        r.font.color.rgb = color
    return tb


def add_multirun_text(slide, x, y, w, h, runs, size=14, align=PP_ALIGN.LEFT,
                       font=FONT, anchor=MSO_ANCHOR.TOP, line_spacing=1.0):
    """runs: list of (text, color, bold, size_override) tuples on ONE line."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    for (t, c, b, sz) in runs:
        r = p.add_run()
        r.text = t
        r.font.size = Pt(sz or size)
        r.font.bold = b
        r.font.name = font
        r.font.color.rgb = c
    return tb


def rounded_rect(slide, x, y, w, h, fill=CARD, line=BORDER, line_w=0.75,
                  radius=0.09, shadow=False, glow_color=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    if glow_color:
        soft_glow(shp, glow_color, radius_pt=14, alpha_pct=35)
    else:
        no_shadow(shp)
    return shp


def oval(slide, x, y, d, fill=CARD, line=None, line_w=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    no_shadow(shp)
    return shp


def icon_circle(slide, cx, cy, d, symbol, fill=CARD_LIGHT, border=CYAN,
                 symbol_color=CYAN, symbol_size=16, border_w=1.25):
    """cx, cy = center coordinates (Emu). Draws a ring circle with a glyph centered."""
    x = Emu(int(cx - d / 2))
    y = Emu(int(cy - d / 2))
    c = oval(slide, x, y, d, fill=fill, line=border, line_w=border_w)
    add_text(slide, x, y, d, d, symbol, size=symbol_size, color=symbol_color,
              bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    return c


def straight_arrow(slide, x1, y1, x2, y2, color=CYAN, width=1.5, dashed=False):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    line_el = ln.line._get_or_add_ln()
    tail = line_el.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
    line_el.append(tail)
    if dashed:
        dash = line_el.makeelement(qn('a:prstDash'), {'val': 'dash'})
        line_el.append(dash)
    return ln


def plain_line(slide, x1, y1, x2, y2, color=BORDER_SOFT, width=0.75, dashed=False):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    if dashed:
        line_el = ln.line._get_or_add_ln()
        dash = line_el.makeelement(qn('a:prstDash'), {'val': 'dash'})
        line_el.append(dash)
    return ln


def pill(slide, x, y, w, h, text, fill=None, line=CYAN, text_color=CYAN, size=11):
    shp = rounded_rect(slide, x, y, w, h, fill=fill, line=line, line_w=1.0, radius=0.5)
    add_text(slide, x, y, w, h, text, size=size, color=text_color, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    return shp


def kicker(slide, x, y, text, color=CYAN, size=12):
    """Small uppercase label above a title — dot + tracked caps text, no underline bar."""
    dot = oval(slide, x, Emu(int(y + Pt(size) * 6500)), Emu(90000), fill=color, line=None)
    add_text(slide, Emu(int(x + Emu(140000))), y, Inches(6), Inches(0.3), text.upper(),
              size=size, color=color, bold=True, align=PP_ALIGN.LEFT)


def footer(slide, page_no, total=10, label="TASKPILOT"):
    add_text(slide, Inches(0.55), Inches(7.14), Inches(3), Inches(0.3),
              label, size=9, color=MUTED_DIM, bold=True, letter_spacing=True)
    add_text(slide, Inches(12.0), Inches(7.14), Inches(0.8), Inches(0.3),
              f"{page_no:02d} / {total:02d}", size=9, color=MUTED_DIM,
              align=PP_ALIGN.RIGHT)


def section_number(slide, n_text):
    add_text(slide, Inches(10.1), Inches(0.55), Inches(2.68), Inches(0.35), n_text,
              size=11, color=MUTED_DIM, bold=True, align=PP_ALIGN.RIGHT, wrap=False)

# ============================================================================
# SLIDE 1 — HERO
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)

# ambient glow blobs (soft translucent circles) behind the hub
glow1 = oval(s, Inches(4.9), Inches(1.0), Inches(3.5), fill=CYAN, line=None)
set_alpha(glow1, 10)
glow2 = oval(s, Inches(6.5), Inches(2.8), Inches(3.0), fill=VIOLET, line=None)
set_alpha(glow2, 8)

# central AI node
hub_cx, hub_cy, hub_d = Inches(6.667), Inches(3.15), Inches(1.5)
hub = oval(s, Emu(int(hub_cx - hub_d/2)), Emu(int(hub_cy - hub_d/2)), hub_d,
           fill=RGBColor(0x10,0x16,0x1F), line=CYAN, line_w=1.75)
soft_glow(hub, CYAN, radius_pt=22, alpha_pct=55)
add_text(s, Emu(int(hub_cx - hub_d/2)), Emu(int(hub_cy - hub_d/2)), hub_d, hub_d,
          "AI", size=30, color=CYAN, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# three satellite nodes: Calendar, Gmail, Web
sat_d = Inches(0.85)
sats = [
    (Inches(3.35), Inches(1.75), "CAL", "Calendar"),
    (Inches(3.35), Inches(4.55), "WEB", "Web Search"),
    (Inches(9.95), Inches(3.15), "MAIL", "Gmail"),
]
for (sx, sy, glyph, label) in sats:
    straight_arrow(s, sx, sy, hub_cx, hub_cy, color=CYAN_DIM, width=1.25)
for (sx, sy, glyph, label) in sats:
    node = oval(s, Emu(int(sx - sat_d/2)), Emu(int(sy - sat_d/2)), sat_d,
                fill=RGBColor(0x0D,0x12,0x19), line=VIOLET, line_w=1.25)
    add_text(s, Emu(int(sx - sat_d/2)), Emu(int(sy - sat_d/2)), sat_d, sat_d,
              glyph, size=11, color=VIOLET, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    add_text(s, Emu(int(sx - Inches(0.9))), Emu(int(sy + sat_d/2 + Inches(0.06))), Inches(1.8), Inches(0.3),
              label, size=10, color=MUTED, align=PP_ALIGN.CENTER)

# Title block
add_text(s, Inches(0), Inches(5.15), Inches(13.333), Inches(1.0), "TASKPILOT",
          size=54, color=WHITE, bold=True, align=PP_ALIGN.CENTER, letter_spacing=True)
add_text(s, Inches(0), Inches(5.98), Inches(13.333), Inches(0.5), "From Intent to Action.",
          size=18, color=CYAN, italic=True, align=PP_ALIGN.CENTER)
add_text(s, Inches(0), Inches(6.5), Inches(13.333), Inches(0.35),
          "AUTONOMOUS AI AGENT FOR EVERYDAY DIGITAL TASKS",
          size=10.5, color=MUTED, bold=True, align=PP_ALIGN.CENTER, letter_spacing=True)
footer(s, 1)


# ============================================================================
# SLIDE 2 — THE PROBLEM
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "01 — PROBLEM")
kicker(s, Inches(0.6), Inches(0.5), "The Problem", color=CYAN)
add_text(s, Inches(0.55), Inches(0.78), Inches(11.5), Inches(0.9),
          "AI Can Answer.\nBut Can It Act?", size=34, color=WHITE, bold=True, line_spacing=1.02)

# comparison row: Traditional AI vs TaskPilot
col_y = Inches(2.35)
col_h = Inches(1.55)
col_w = Inches(5.6)
col1_x = Inches(0.6)
col2_x = Inches(7.1)

rounded_rect(s, col1_x, col_y, col_w, col_h, fill=CARD, line=BORDER)
add_text(s, col1_x + Inches(0.35), col_y + Inches(0.22), col_w - Inches(0.7), Inches(0.35),
          "TRADITIONAL AI", size=12, color=MUTED, bold=True, letter_spacing=True)
add_multirun_text(s, col1_x + Inches(0.35), col_y + Inches(0.68), col_w - Inches(0.7), Inches(0.6),
    [("ASK", MUTED, True, 22), ("   →   ", MUTED_DIM, False, 20), ("ANSWER", MUTED, True, 22)])

rounded_rect(s, col2_x, col_y, col_w, col_h, fill=RGBColor(0x0D,0x18,0x1E), line=CYAN, line_w=1.25, glow_color=CYAN)
add_text(s, col2_x + Inches(0.35), col_y + Inches(0.22), col_w - Inches(0.7), Inches(0.35),
          "TASKPILOT", size=12, color=CYAN, bold=True, letter_spacing=True)
add_multirun_text(s, col2_x + Inches(0.35), col_y + Inches(0.68), col_w - Inches(0.7), Inches(0.6),
    [("INTENT", CYAN, True, 16.5), (" → ", MUTED_DIM, False, 16), ("PLAN", CYAN, True, 16.5),
     (" → ", MUTED_DIM, False, 16), ("ACT", CYAN, True, 16.5), (" → ", MUTED_DIM, False, 16),
     ("VERIFY", CYAN, True, 16.5)])

# center arrow between columns
straight_arrow(s, col1_x + col_w + Inches(0.08), col_y + col_h/2, col2_x - Inches(0.08), col_y + col_h/2,
               color=VIOLET, width=1.5)

# pain points row — 4 compact cards
pains = [
    ("01", "Tasks are scattered across multiple apps"),
    ("02", "Users manually coordinate workflows"),
    ("03", "AI responses don't complete actions"),
    ("04", "Multi-step tasks need repeated effort"),
]
pw = Inches(2.95)
ph = Inches(1.95)
gap = Inches(0.22)
start_x = Inches(0.6)
py = Inches(4.35)
for i, (num, text) in enumerate(pains):
    x = start_x + i * (pw + gap)
    rounded_rect(s, x, py, pw, ph, fill=CARD, line=BORDER_SOFT)
    add_text(s, x + Inches(0.25), py + Inches(0.22), pw - Inches(0.5), Inches(0.5),
              num, size=20, color=VIOLET, bold=True)
    add_text(s, x + Inches(0.25), py + Inches(0.85), pw - Inches(0.5), Inches(1.0),
              text, size=12.5, color=WHITE, line_spacing=1.15)

footer(s, 2)

# ============================================================================
# SLIDE 3 — THE SOLUTION
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "02 — SOLUTION")
kicker(s, Inches(0.6), Inches(0.5), "The Solution", color=CYAN)
add_text(s, Inches(0.55), Inches(0.78), Inches(9.0), Inches(0.65),
          "Meet TaskPilot", size=34, color=WHITE, bold=True)
add_text(s, Inches(0.55), Inches(1.42), Inches(9.6), Inches(0.65),
          "TaskPilot transforms natural-language goals into executable, verified workflows —\nnot just answers.",
          size=13.5, color=MUTED, line_spacing=1.25)

# Vertical flow: 5 stages
stage_labels = ["USER REQUEST", "AI AGENT", "TOOLS", "REAL-WORLD ACTION", "VERIFICATION"]
n = len(stage_labels)
flow_top = Inches(2.35)
flow_h = Inches(0.62)
flow_gap = Inches(0.22)
flow_w = Inches(4.6)
flow_x = Inches(0.6)
for i, label in enumerate(stage_labels):
    y = flow_top + i * (flow_h + flow_gap)
    is_agent = (label == "AI AGENT")
    rounded_rect(s, flow_x, y, flow_w, flow_h,
                 fill=RGBColor(0x0D,0x18,0x1E) if is_agent else CARD,
                 line=CYAN if is_agent else BORDER, line_w=1.25 if is_agent else 0.75,
                 glow_color=CYAN if is_agent else None)
    add_text(s, flow_x + Inches(0.3), y, flow_w - Inches(0.6), flow_h, label,
              size=14, color=CYAN if is_agent else WHITE, bold=True,
              anchor=MSO_ANCHOR.MIDDLE)
    if i < n - 1:
        ay = y + flow_h + Emu(int(flow_gap/2))
        straight_arrow(s, flow_x + flow_w/2, y + flow_h + Emu(30000), flow_x + flow_w/2,
                        y + flow_h + flow_gap - Emu(30000), color=VIOLET, width=1.25)

# Right side: tool cards (Calendar / Gmail / Web)
tools = [
    ("CAL", "Calendar", "Reads availability, creates events"),
    ("MAIL", "Gmail", "Reads & drafts relevant emails"),
    ("WEB", "Web Search", "Looks up needed information"),
]
tx = Inches(5.75)
tw = Inches(6.95)
th = Inches(1.15)
ty0 = Inches(2.35)
tgap = Inches(0.28)
for i, (glyph, name, desc) in enumerate(tools):
    y = ty0 + i * (th + tgap)
    rounded_rect(s, tx, y, tw, th, fill=CARD, line=BORDER_SOFT)
    icon_circle(s, tx + Inches(0.65), y + th/2, Inches(0.62), glyph,
                fill=CARD_LIGHT, border=VIOLET, symbol_color=VIOLET, symbol_size=10.5)
    add_text(s, tx + Inches(1.15), y + Inches(0.18), tw - Inches(1.4), Inches(0.4),
              name, size=15, color=WHITE, bold=True)
    add_text(s, tx + Inches(1.15), y + Inches(0.62), tw - Inches(1.4), Inches(0.45),
              desc, size=11.5, color=MUTED)

footer(s, 3)

# ============================================================================
# SLIDE 4 — HOW THE AGENT THINKS
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "03 — AGENT LOOP")
kicker(s, Inches(0.6), Inches(0.45), "How The Agent Thinks", color=CYAN)
add_text(s, Inches(0.55), Inches(0.72), Inches(11.5), Inches(0.65),
          "From Intent to Execution", size=32, color=WHITE, bold=True)

steps = [
    ("1", "UNDER-\nSTAND", "Parses the user's natural-language goal"),
    ("2", "DECOM-\nPOSE", "Breaks the goal into discrete subtasks"),
    ("3", "PLAN", "Sequences subtasks into an execution plan"),
    ("4", "SELECT\nTOOL", "Chooses the right tool for each step"),
    ("5", "APPROVE", "Requests human sign-off when needed"),
    ("6", "EXECUTE", "Performs the approved actions"),
    ("7", "VERIFY", "Confirms the outcome actually occurred"),
]
n = len(steps)
top = Inches(2.05)
card_w = Inches(1.62)
card_h = Inches(3.55)
gap = Inches(0.135)
total_w = n*card_w + (n-1)*gap
start_x = (SLIDE_W - total_w) / 2
for i, (num, label, desc) in enumerate(steps):
    x = Emu(int(start_x + i*(card_w+gap)))
    is_approve = (label == "APPROVE")
    rounded_rect(s, x, top, card_w, card_h,
                 fill=RGBColor(0x14,0x12,0x1E) if is_approve else CARD,
                 line=VIOLET if is_approve else BORDER, line_w=1.25 if is_approve else 0.75,
                 glow_color=VIOLET if is_approve else None)
    ccx = x + card_w/2
    icon_circle(s, ccx, top + Inches(0.55), Inches(0.62), num,
                fill=RGBColor(0x0D,0x18,0x1E), border=CYAN, symbol_color=CYAN, symbol_size=16)
    add_text(s, x + Inches(0.08), top + Inches(1.1), card_w - Inches(0.16), Inches(0.7),
              label, size=12.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER, line_spacing=0.95)
    add_text(s, x + Inches(0.14), top + Inches(1.95), card_w - Inches(0.28), Inches(1.5),
              desc, size=9.5, color=MUTED, align=PP_ALIGN.CENTER, line_spacing=1.15)
    if i < n-1:
        mid_y = top + Inches(0.55)
        straight_arrow(s, x+card_w+Emu(15000), mid_y, x+card_w+gap-Emu(15000), mid_y, color=CYAN_DIM, width=1.1)

add_text(s, Inches(0.6), Inches(5.95), Inches(12.1), Inches(0.4),
          "A closed action loop — not a single reply.", size=13, color=MUTED, italic=True, align=PP_ALIGN.CENTER)
footer(s, 4)

# ============================================================================
# SLIDE 5 — SYSTEM ARCHITECTURE
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "04 — ARCHITECTURE")
kicker(s, Inches(0.6), Inches(0.4), "System Architecture", color=CYAN)
add_text(s, Inches(0.55), Inches(0.67), Inches(9.0), Inches(0.6),
          "Inside TaskPilot", size=30, color=WHITE, bold=True)

layer_x = Inches(0.6)
layer_w = Inches(12.13)

def layer(y, h, text, fill=CARD, line=BORDER, tcolor=WHITE, size=13.5, sub=None):
    rounded_rect(s, layer_x, y, layer_w, h, fill=fill, line=line, line_w=0.85)
    add_text(s, layer_x, y, layer_w, h, text, size=size, color=tcolor, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if sub:
        add_text(s, layer_x, y+h-Inches(0.22), layer_w, Inches(0.22), sub, size=9,
                  color=MUTED_DIM, align=PP_ALIGN.CENTER)

y = Inches(1.42)
h1 = Inches(0.44)
layer(y, h1, "USER")
y2 = y + h1 + Inches(0.10)
straight_arrow(s, SLIDE_W/2, y+h1, SLIDE_W/2, y2, color=CYAN_DIM, width=1.0)

h2 = Inches(0.44)
layer(y2, h2, "NEXT.JS / REACT FRONTEND", fill=RGBColor(0x0D,0x16,0x1C))
y3 = y2 + h2 + Inches(0.10)
straight_arrow(s, SLIDE_W/2, y2+h2, SLIDE_W/2, y3, color=CYAN_DIM, width=1.0)

h3 = Inches(0.44)
layer(y3, h3, "FASTAPI BACKEND", fill=RGBColor(0x0D,0x16,0x1C))
y4 = y3 + h3 + Inches(0.14)
straight_arrow(s, SLIDE_W/2, y3+h3, SLIDE_W/2, y4, color=CYAN, width=1.25)

# Agent orchestrator big box with 5 internal modules
h4 = Inches(1.55)
rounded_rect(s, layer_x, y4, layer_w, h4, fill=RGBColor(0x0D,0x18,0x1E), line=CYAN, line_w=1.25, glow_color=CYAN)
add_text(s, layer_x+Inches(0.3), y4+Inches(0.12), layer_w-Inches(0.6), Inches(0.32),
          "AGENT ORCHESTRATOR", size=13.5, color=CYAN, bold=True)
mods = ["LLM", "PLANNER", "TOOL SELECTOR", "APPROVAL MANAGER", "VERIFICATION ENGINE"]
mw = Inches(2.25)
mh = Inches(0.78)
mgap = (layer_w - Inches(0.6) - mw*len(mods)) / (len(mods)-1)
for i, m in enumerate(mods):
    mx = layer_x + Inches(0.3) + i*(mw+mgap)
    my = y4 + Inches(0.56)
    rounded_rect(s, mx, my, mw, mh, fill=CARD_LIGHT, line=BORDER, line_w=0.75, radius=0.14)
    add_text(s, mx+Inches(0.08), my, mw-Inches(0.16), mh, m, size=10.5, color=WHITE, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.95)

y5 = y4 + h4 + Inches(0.14)
straight_arrow(s, SLIDE_W/2, y4+h4, SLIDE_W/2, y5, color=CYAN, width=1.25)

# Tools row
h5 = Inches(0.85)
rounded_rect(s, layer_x, y5, layer_w, h5, fill=CARD, line=BORDER)
add_text(s, layer_x+Inches(0.3), y5+Inches(0.08), layer_w-Inches(0.6), Inches(0.26), "TOOLS",
          size=10.5, color=MUTED, bold=True, letter_spacing=True)
tool_names = ["Google Calendar", "Gmail", "Web Search"]
tw2 = Inches(2.6)
tgap2 = Inches(0.3)
tstart = layer_x + (layer_w - (tw2*3 + tgap2*2))/2
for i, tname in enumerate(tool_names):
    tx2 = tstart + i*(tw2+tgap2)
    ty2 = y5 + Inches(0.38)
    pill(s, tx2, ty2, tw2, Inches(0.4), tname, fill=RGBColor(0x14,0x0F,0x1E), line=VIOLET, text_color=VIOLET, size=11)

y6 = y5 + h5 + Inches(0.14)
straight_arrow(s, SLIDE_W/2, y5+h5, SLIDE_W/2, y6, color=CYAN_DIM, width=1.0)

h6 = Inches(0.44)
layer(y6, h6, "DATABASE  —  SUPABASE / POSTGRESQL", fill=RGBColor(0x0D,0x16,0x1C), size=12.5)

footer(s, 5)

# ============================================================================
# SLIDE 6 — HUMAN-IN-THE-LOOP
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "05 — CONTROL")
kicker(s, Inches(0.6), Inches(0.5), "Human-in-the-Loop", color=VIOLET)
add_text(s, Inches(0.55), Inches(0.78), Inches(11.5), Inches(0.65),
          "Autonomy With Control", size=32, color=WHITE, bold=True)

# horizontal 4-stage flow
stages = ["AI PROPOSES\nACTION", "USER\nAPPROVES", "AI\nEXECUTES", "SYSTEM\nVERIFIES"]
n = len(stages)
sw = Inches(2.55)
sh = Inches(1.15)
sgap = Inches(0.55)
total = n*sw + (n-1)*sgap
sx0 = (SLIDE_W - total)/2
sy = Inches(1.85)
for i, label in enumerate(stages):
    x = sx0 + i*(sw+sgap)
    highlight = (i == 1)
    rounded_rect(s, x, sy, sw, sh,
                 fill=RGBColor(0x14,0x12,0x1E) if highlight else CARD,
                 line=VIOLET if highlight else BORDER, line_w=1.25 if highlight else 0.75,
                 glow_color=VIOLET if highlight else None)
    add_text(s, x, sy, sw, sh, label, size=13, color=VIOLET if highlight else WHITE, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    if i < n-1:
        midy = sy + sh/2
        straight_arrow(s, x+sw+Emu(20000), midy, x+sw+sgap-Emu(20000), midy, color=CYAN, width=1.4)

# Example approval card mockup
card_x = Inches(3.3)
card_y = Inches(3.55)
card_w = Inches(6.7)
card_h = Inches(2.85)
rounded_rect(s, card_x, card_y, card_w, card_h, fill=RGBColor(0x10,0x14,0x1B), line=CYAN, line_w=1.25, glow_color=CYAN)
icon_circle(s, card_x+Inches(0.55), card_y+Inches(0.5), Inches(0.5), "AI",
            fill=CARD_LIGHT, border=CYAN, symbol_color=CYAN, symbol_size=11)
add_text(s, card_x+Inches(0.95), card_y+Inches(0.27), card_w-Inches(1.3), Inches(0.5),
          "TaskPilot wants to create 4 calendar events.", size=15, color=WHITE, bold=True,
          anchor=MSO_ANCHOR.MIDDLE)
add_text(s, card_x+Inches(0.55), card_y+Inches(1.0), card_w-Inches(1.1), Inches(0.4),
          "Mon 9–11am · Tue 2–4pm · Wed 9–11am · Thu 2–4pm  —  Study Sessions",
          size=11.5, color=MUTED)

btn_y = card_y + Inches(1.75)
btn_h = Inches(0.7)
btns = [("APPROVE", GREEN), ("EDIT", AMBER), ("CANCEL", MUTED)]
bw = Inches(1.85)
bgap = Inches(0.3)
bstart = card_x + (card_w - (bw*3+bgap*2))/2
for i, (label, color) in enumerate(btns):
    bx = bstart + i*(bw+bgap)
    rounded_rect(s, bx, btn_y, bw, btn_h, fill=None, line=color, line_w=1.25, radius=0.28)
    add_text(s, bx, btn_y, bw, btn_h, label, size=13, color=color, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.6), Inches(6.75), Inches(12.1), Inches(0.4),
          "The agent automates workflows without removing user control.",
          size=13, color=MUTED, italic=True, align=PP_ALIGN.CENTER)
footer(s, 6)

# ============================================================================
# SLIDE 7 — LIVE DEMO
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "06 — DEMO")
kicker(s, Inches(0.6), Inches(0.45), "Live Demo", color=CYAN)
add_text(s, Inches(0.55), Inches(0.72), Inches(11.5), Inches(0.6),
          "One Request. Multiple Actions.", size=30, color=WHITE, bold=True)

# Chat-window mockup
win_x, win_y, win_w, win_h = Inches(1.75), Inches(1.65), Inches(9.83), Inches(5.15)
rounded_rect(s, win_x, win_y, win_w, win_h, fill=RGBColor(0x0D,0x11,0x18), line=BORDER, line_w=1.0, radius=0.05)
# title bar
rounded_rect(s, win_x, win_y, win_w, Inches(0.5), fill=RGBColor(0x11,0x16,0x1F), line=None, radius=0.28)
for i, c in enumerate([RGBColor(0xE0,0x5D,0x5D), RGBColor(0xE0,0xB4,0x5D), RGBColor(0x5D,0xC0,0x7A)]):
    oval(s, win_x+Inches(0.28)+i*Inches(0.26), win_y+Inches(0.19), Inches(0.13), fill=c, line=None)
add_text(s, win_x, win_y+Inches(0.06), win_w, Inches(0.38), "TaskPilot — Agent Console",
          size=11, color=MUTED, bold=True, align=PP_ALIGN.CENTER)

# user message bubble (right aligned)
um_w = Inches(6.6)
um_h = Inches(0.7)
um_x = win_x + win_w - um_w - Inches(0.4)
um_y = win_y + Inches(0.75)
rounded_rect(s, um_x, um_y, um_w, um_h, fill=RGBColor(0x14,0x1B,0x26), line=BORDER, line_w=0.75, radius=0.22)
add_text(s, um_x+Inches(0.3), um_y, um_w-Inches(0.6), um_h,
          "\u201cPlan my study schedule for this week and add it to my calendar.\u201d",
          size=12.5, color=WHITE, italic=True, anchor=MSO_ANCHOR.MIDDLE)

# agent response panel — checklist
panel_x = win_x + Inches(0.4)
panel_y = um_y + um_h + Inches(0.3)
panel_w = win_w - Inches(0.8)
panel_h = Inches(3.15)
rounded_rect(s, panel_x, panel_y, panel_w, panel_h, fill=CARD, line=BORDER_SOFT, radius=0.06)

checklist = [
    ("\u2713", GREEN, "Calendar analyzed"),
    ("\u2713", GREEN, "Free slots identified"),
    ("\u2713", GREEN, "Study plan generated"),
    ("\u2713", GREEN, "4 events proposed"),
    ("\u26a0", AMBER, "Approval required"),
    ("\u2713", GREEN, "Events created"),
    ("\u2713", GREEN, "Results verified"),
]
row_h = panel_h / len(checklist)
for i, (glyph, color, text) in enumerate(checklist):
    ry = panel_y + i*row_h
    add_text(s, panel_x+Inches(0.35), ry, Inches(0.5), row_h, glyph, size=16, color=color,
              bold=True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, panel_x+Inches(0.85), ry, panel_w-Inches(1.2), row_h, text, size=13.5,
              color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if i < len(checklist)-1:
        plain_line(s, panel_x+Inches(0.35), ry+row_h, panel_x+panel_w-Inches(0.35), ry+row_h,
                   color=BORDER_SOFT, width=0.5)

footer(s, 7)

# ============================================================================
# SLIDE 8 — TECH STACK
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "07 — TECH STACK")
kicker(s, Inches(0.6), Inches(0.5), "Tech Stack", color=CYAN)
add_text(s, Inches(0.55), Inches(0.78), Inches(11.5), Inches(0.65),
          "Built for Speed. Designed to Scale.", size=30, color=WHITE, bold=True)

stack = [
    ("FE", "Frontend", "React · Next.js · Tailwind CSS"),
    ("BE", "Backend", "Python · FastAPI"),
    ("AI", "AI Layer", "LLM · Tool Calling · Orchestration"),
    ("DB", "Data", "Supabase · PostgreSQL"),
    ("INT", "Integrations", "Calendar · Gmail · Web Search"),
    ("DEP", "Deployment", "Vercel · Cloud Backend"),
]
cols = 3
rows = 2
cw = Inches(3.85)
ch = Inches(1.75)
cgapx = Inches(0.28)
cgapy = Inches(0.28)
total_w = cols*cw + (cols-1)*cgapx
start_x = (SLIDE_W - total_w)/2
start_y = Inches(2.05)
for i, (glyph, name, desc) in enumerate(stack):
    r, c = divmod(i, cols)
    x = start_x + c*(cw+cgapx)
    y = start_y + r*(ch+cgapy)
    rounded_rect(s, x, y, cw, ch, fill=CARD, line=BORDER)
    icon_circle(s, x+Inches(0.62), y+Inches(0.55), Inches(0.62), glyph,
                fill=CARD_LIGHT, border=CYAN, symbol_color=CYAN, symbol_size=10.5)
    add_text(s, x+Inches(1.1), y+Inches(0.28), cw-Inches(1.35), Inches(0.4), name,
              size=15.5, color=WHITE, bold=True)
    add_text(s, x+Inches(0.35), y+Inches(1.15), cw-Inches(0.7), Inches(0.5), desc,
              size=11.5, color=MUTED, line_spacing=1.15)

footer(s, 8)

# ============================================================================
# SLIDE 9 — WHY TASKPILOT
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "08 — DIFFERENTIATION")
kicker(s, Inches(0.6), Inches(0.5), "Why TaskPilot", color=CYAN)
add_text(s, Inches(0.55), Inches(0.78), Inches(11.5), Inches(0.65),
          "Beyond the Chatbot", size=32, color=WHITE, bold=True)

col_y = Inches(1.85)
col_h = Inches(4.85)
col_w = Inches(5.7)
col1_x = Inches(0.7)
col2_x = Inches(6.95)

rounded_rect(s, col1_x, col_y, col_w, col_h, fill=CARD, line=BORDER)
add_text(s, col1_x+Inches(0.45), col_y+Inches(0.35), col_w-Inches(0.9), Inches(0.4),
          "CHATBOT", size=15, color=MUTED, bold=True, letter_spacing=True)
chat_pts = ["Responds", "Gives information", "User performs actions", "Single interaction"]
for i, pt in enumerate(chat_pts):
    yy = col_y+Inches(1.05)+i*Inches(0.62)
    oval(s, col1_x+Inches(0.48), yy+Inches(0.08), Inches(0.09), fill=MUTED, line=None)
    add_text(s, col1_x+Inches(0.75), yy, col_w-Inches(1.1), Inches(0.5), pt, size=14, color=WHITE)

rounded_rect(s, col2_x, col_y, col_w, col_h, fill=RGBColor(0x0D,0x18,0x1E), line=CYAN, line_w=1.25, glow_color=CYAN)
add_text(s, col2_x+Inches(0.45), col_y+Inches(0.35), col_w-Inches(0.9), Inches(0.4),
          "TASKPILOT", size=15, color=CYAN, bold=True, letter_spacing=True)
tp_pts = ["Plans multi-step workflows", "Selects & uses tools", "Executes real actions",
          "Verifies results", "Human approval for key actions", "Handles multi-app tasks"]
for i, pt in enumerate(tp_pts):
    yy = col_y+Inches(1.05)+i*Inches(0.55)
    oval(s, col2_x+Inches(0.48), yy+Inches(0.08), Inches(0.09), fill=CYAN, line=None)
    add_text(s, col2_x+Inches(0.75), yy, col_w-Inches(1.1), Inches(0.5), pt, size=13.5, color=WHITE)

add_text(s, Inches(0.6), Inches(6.9), Inches(12.1), Inches(0.35),
          "A difference in capability, not a judgment of either approach.",
          size=11.5, color=MUTED_DIM, italic=True, align=PP_ALIGN.CENTER)
footer(s, 9)

# ============================================================================
# SLIDE 10 — FUTURE / CLOSING
# ============================================================================
s = new_slide()
bg_fill(s)
grid_lines(s, n_v=8, n_h=6, alpha=100)
section_number(s, "09 — ROADMAP")
kicker(s, Inches(0.6), Inches(0.5), "Roadmap", color=VIOLET)
add_text(s, Inches(0.55), Inches(0.78), Inches(11.5), Inches(0.65),
          "Where TaskPilot Goes Next", size=32, color=WHITE, bold=True)

phases = [
    ("NOW", "Calendar + Gmail + Web", CYAN),
    ("NEXT", "More integrations · Browser automation ·\nPersistent memory · Multi-agent workflows", VIOLET),
    ("FUTURE", "Personal AI operating layer ·\nCross-app task execution · Proactive assistance", MUTED),
]
pw = Inches(3.85)
ph = Inches(2.55)
pgap = Inches(0.3)
total = pw*3 + pgap*2
pstart = (SLIDE_W - total)/2
py = Inches(1.9)
for i, (label, desc, color) in enumerate(phases):
    x = pstart + i*(pw+pgap)
    rounded_rect(s, x, py, pw, ph, fill=CARD if i>0 else RGBColor(0x0D,0x18,0x1E),
                 line=color, line_w=1.25, glow_color=color if i==0 else None)
    pill(s, x+Inches(0.35), py+Inches(0.3), Inches(1.5), Inches(0.42), label,
         fill=None, line=color, text_color=color, size=12)
    add_text(s, x+Inches(0.35), py+Inches(1.0), pw-Inches(0.7), ph-Inches(1.3),
              desc, size=12.5, color=WHITE, line_spacing=1.3)
    if i < 2:
        midy = py+ph/2
        straight_arrow(s, x+pw+Emu(20000), midy, x+pw+pgap-Emu(20000), midy, color=MUTED_DIM, width=1.25)

add_text(s, Inches(0), Inches(5.05), Inches(13.333), Inches(0.6),
          "TASKPILOT", size=26, color=WHITE, bold=True, align=PP_ALIGN.CENTER, letter_spacing=True)
add_text(s, Inches(0), Inches(5.68), Inches(13.333), Inches(0.5),
          "\u201cGive AI a goal. Let it handle the workflow.\u201d",
          size=16, color=CYAN, italic=True, align=PP_ALIGN.CENTER)

footer(s, 10)

# ----------------------------------------------------------------------------
prs.save("/home/claude/taskpilot/TaskPilot_AI_Hackathon_2026.pptx")
print("Saved.")
