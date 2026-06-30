# -*- coding: utf-8 -*-
"""Rebuild BLC_Commercial_Deck in real Braze styling (extracted from approved
slideware: palette, Arial, BrazeAI logo). Canvas 10 x 5.625in. Content verbatim."""
import os
from pptx import Presentation
from pptx.util import Inches as _IN, Pt as _PT, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Design space is 13.333 x 7.5; emit at 0.75 scale -> 10 x 5.625 (Braze size)
SCALE = 0.75
def IN(v): return _IN(v * SCALE)
def PT(v): return _PT(v * SCALE)

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "ref", "brazeai_logo.png")   # 1200x513, aspect 2.339
LOGO_AR = 1200 / 513

# ---------- Braze brand kit (from approved theme) ----------
PURPLE   = RGBColor(0x80, 0x1E, 0xD7)   # accent2 — brand purple
PURPLE_D = RGBColor(0x30, 0x02, 0x66)   # dk2 — deep purple
GOLD     = RGBColor(0xFF, 0xA5, 0x24)   # accent5 — gold
LAV      = RGBColor(0xEF, 0xED, 0xFF)   # lt2 — light lavender
PERI     = RGBColor(0xC9, 0xC4, 0xFF)   # accent6 — periwinkle
INK      = RGBColor(0x1A, 0x1A, 0x24)
SLATE    = RGBColor(0x55, 0x52, 0x66)
MUTE     = RGBColor(0x8A, 0x86, 0x99)
LINE     = RGBColor(0xE3, 0xDF, 0xF5)
ONLAV    = RGBColor(0xD9, 0xD3, 0xFB)   # light text on purple
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
PAPER    = RGBColor(0xFA, 0xF9, 0xFE)

F_HEAD = "Arial"
F_BODY = "Arial"

ML = 0.6
CW = 13.333 - 2 * ML  # 12.133 design in

prs = Presentation()
prs.slide_width = IN(13.333)
prs.slide_height = IN(7.5)
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)

def _spc(run, val): run.font._rPr.set('spc', str(val))
def _supersc(run, val=30000): run.font._rPr.set('baseline', str(val))

def tb(s, l, t, w, h, anchor=None):
    box = s.shapes.add_textbox(IN(l), IN(t), IN(w), IN(h))
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    if anchor: tf.vertical_anchor = anchor
    return box, tf

def para(tf, runs, align=PP_ALIGN.LEFT, first=False, before=None, after=None, line=None):
    if first and len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.alignment = align
    if before is not None: p.space_before = PT(before)
    if after is not None: p.space_after = PT(after)
    if line is not None: p.line_spacing = line
    if isinstance(runs, dict): runs = [runs]
    for spec in runs:
        r = p.add_run(); r.text = spec['text']
        r.font.size = PT(spec.get('size', 12))
        r.font.bold = spec.get('bold', False)
        r.font.italic = spec.get('italic', False)
        r.font.name = spec.get('font', F_BODY)
        r.font.color.rgb = spec.get('color', INK)
        if 'spc' in spec: _spc(r, spec['spc'])
        if spec.get('super'): _supersc(r)
    return p

def shape(s, kind, l, t, w, h, fill=None, line_c=None, line_w=0.75, radius=None):
    sp = s.shapes.add_shape(kind, IN(l), IN(t), IN(w), IN(h))
    sp.shadow.inherit = False
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line_c is None: sp.line.fill.background()
    else: sp.line.color.rgb = line_c; sp.line.width = PT(line_w)
    if radius is not None:
        try: sp.adjustments[0] = radius
        except Exception: pass
    sp.text_frame.paragraphs[0].text = ""
    return sp

def card(s, l, t, w, h, fill=LAV, line_c=LINE, line_w=1.0, radius=0.06):
    return shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h, fill=fill,
                 line_c=line_c, line_w=line_w, radius=radius)

def rect(s, l, t, w, h, fill, line_c=None, line_w=0.75):
    return shape(s, MSO_SHAPE.RECTANGLE, l, t, w, h, fill=fill, line_c=line_c, line_w=line_w)

def logo(s, h=0.40, right=ML, top=0.34):
    w = h * LOGO_AR
    x = 13.333 - right - w
    s.shapes.add_picture(LOGO, IN(x), IN(top), height=IN(h))

def pagenum(s, n):
    box, tf = tb(s, 13.333 - 0.6 - 0.8, 7.5 - 0.42, 0.8, 0.3)
    para(tf, {'text': str(n), 'size': 9, 'color': MUTE, 'font': F_BODY},
         align=PP_ALIGN.RIGHT, first=True)

def header(s, title, sub=None, kicker=None, tsize=24):
    y = 0.5
    if kicker:
        box, tf = tb(s, ML, y, CW - 2.4, 0.3)
        para(tf, {'text': kicker.upper(), 'size': 11, 'color': GOLD, 'bold': True,
                  'font': F_HEAD, 'spc': 200}, first=True)
        y += 0.34
    box, tf = tb(s, ML, y, CW - 0.2, 0.7)
    para(tf, {'text': title, 'size': tsize, 'color': PURPLE, 'bold': True,
              'font': F_HEAD}, first=True, line=1.0)
    y += 0.16 + tsize / 50.0
    if sub:
        box, tf = tb(s, ML, y, CW, 0.5)
        para(tf, {'text': sub, 'size': 13, 'color': GOLD, 'bold': True,
                  'font': F_HEAD}, first=True, line=1.05)
        y += 0.4
    rect(s, ML, y + 0.05, 1.3, 0.045, GOLD)
    return y + 0.3

def footnote(s, text, y=7.02):
    box, tf = tb(s, ML, y, CW, 0.42)
    para(tf, {'text': text, 'size': 7.5, 'color': MUTE, 'font': F_BODY,
              'italic': True}, first=True, line=1.02)

def chrome(s, n):
    logo(s); pagenum(s, n)


# ====================== SLIDE 1 — Title ======================
s = slide()
rect(s, 0, 0, 13.333, 7.5, WHITE)
rect(s, 0, 0, 0.16, 7.5, PURPLE)
logo(s, h=0.46)
box, tf = tb(s, ML + 0.2, 2.45, 9.5, 0.4)
para(tf, {'text': 'AI DECISIONING', 'size': 13, 'color': GOLD, 'bold': True,
          'font': F_HEAD, 'spc': 240}, first=True)
box, tf = tb(s, ML + 0.18, 2.9, 11.4, 1.6)
para(tf, {'text': 'Commercials for Blue Light Card', 'size': 40, 'color': PURPLE,
          'bold': True, 'font': F_HEAD}, first=True, line=1.0)
rect(s, ML + 0.22, 4.35, 2.2, 0.06, GOLD)
box, tf = tb(s, ML + 0.2, 4.6, 10.0, 0.4)
para(tf, {'text': 'Year-one investment, cost structure & path to ROI', 'size': 15,
          'color': SLATE, 'font': F_BODY}, first=True)
box, tf = tb(s, ML + 0.2, 6.75, 6.0, 0.35)
para(tf, {'text': 'Confidential', 'size': 10.5, 'color': MUTE, 'bold': True,
          'font': F_BODY, 'spc': 60}, first=True)

# ====================== SLIDE 2 — Mutual action plan ======================
s = slide()
cy = header(s, 'Mutual action plan to go-live',
            sub='Foundations are complete. Next is commercial alignment and executive sign-off to lock the Nov 2026 implementation slot.',
            tsize=23)
chrome(s, 2)
rows = [
    ('1', 'Priority use case identified — Engagement (absorbs re-engagement + cross-cat)', 'BLC + Braze', '27 May', '✓ Completed', PURPLE),
    ('2', 'Business case & impact model built (£4.37M target run-rate)', 'Braze', '27 May', '✓ Completed', PURPLE),
    ('3', 'Data validation & audience sizing (4.2M MAU)', 'BLC + Braze', '27 May', '✓ Completed', PURPLE),
    ('4', 'Customer Readiness Assessment (CRA) — in-flight', 'Braze', '16 Jun', '● In flight', GOLD),
    ('5', 'Commercial proposal & cost structure shared', 'Braze', '17 Jun', '● Today', PURPLE_D),
    ('6', 'Lifecycle team review model & price → internal alignment', 'BLC (Marie)', 'w/c 23 Jun', 'Upcoming', MUTE),
    ('7', 'Executive review — Director of Data, CPO, Head of Engineering', 'BLC + Braze', 'w/c 23 Jun', 'Upcoming', MUTE),
    ('8', 'Finance approval (if incremental budget required)', 'BLC', 'early Jul', 'Upcoming', MUTE),
    ('9', 'Go / No-Go decision', 'BLC', 'by mid-Jul', 'Upcoming', MUTE),
    ('10', 'Order form finalised & signature', 'BLC + Braze', 'late Jul', 'Upcoming', MUTE),
    ('11', 'Implementation kickoff', 'Braze + BLC', 'November 2026', 'Upcoming', MUTE),
    ('12', 'Go live — Engagement engine in production', 'Braze + BLC', 'February 2027', 'Upcoming', MUTE),
]
tbl_top = cy + 0.02; tbl_h = 4.30
gtbl = s.shapes.add_table(len(rows) + 1, 5, IN(ML), IN(tbl_top), IN(CW), IN(tbl_h)).table
gtbl.first_row = False; gtbl.horz_banding = False
for i, wd in enumerate([0.5, 6.43, 1.7, 1.5, 2.0]): gtbl.columns[i].width = IN(wd)
for j, htext in enumerate(['#', 'Milestone', 'Owner', 'Target', 'Status']):
    c = gtbl.cell(0, j); c.fill.solid(); c.fill.fore_color.rgb = PURPLE
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.margin_left = IN(0.07); c.margin_right = IN(0.05); c.margin_top = IN(0.02); c.margin_bottom = IN(0.02)
    para(c.text_frame, {'text': htext, 'size': 9, 'color': WHITE, 'bold': True, 'font': F_HEAD},
         align=PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT, first=True)
gtbl.rows[0].height = IN(0.34)
for ri, (num, mile, owner, targ, stat, sc) in enumerate(rows, start=1):
    fill = WHITE if ri % 2 else LAV
    vals = [(num, INK, PP_ALIGN.CENTER, False), (mile, INK, PP_ALIGN.LEFT, False),
            (owner, SLATE, PP_ALIGN.LEFT, False), (targ, SLATE, PP_ALIGN.LEFT, False),
            (stat, sc, PP_ALIGN.LEFT, True)]
    for j, (txt, col, al, bold) in enumerate(vals):
        c = gtbl.cell(ri, j); c.fill.solid(); c.fill.fore_color.rgb = fill
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = IN(0.07); c.margin_right = IN(0.05); c.margin_top = IN(0.01); c.margin_bottom = IN(0.01)
        para(c.text_frame, {'text': txt, 'size': 8.2, 'color': col, 'bold': bold, 'font': F_BODY},
             align=al, first=True)
    gtbl.rows[ri].height = IN((tbl_h - 0.34) / len(rows))
footnote(s, "Note: BLC's current Braze contract was renewed Jan 2026 ($1M+). Decisioning Studio is net-new on top of that contract.", y=tbl_top + tbl_h + 0.06)

# ====================== SLIDE 3 — Year-one investment ======================
s = slide()
cy = header(s, 'Year-one investment — Engagement proof of value',
            sub='Priced at the 4M-MAU tier — sign by 17 July to lock the early-sign rate',
            kicker='Commercials', tsize=22)
chrome(s, 3)
top = cy + 0.05
hx, hw = ML, 6.05
card(s, hx, top, hw, 3.95, fill=PURPLE, line_c=None, radius=0.05)
box, tf = tb(s, hx + 0.35, top + 0.3, hw - 0.7, 0.4)
para(tf, {'text': '35% PARTNERSHIP DISCOUNT', 'size': 11, 'color': GOLD, 'bold': True,
          'font': F_HEAD, 'spc': 120}, first=True)
box, tf = tb(s, hx + 0.35, top + 0.72, hw - 0.7, 0.35)
para(tf, [{'text': 'List price  ', 'size': 12, 'color': ONLAV, 'font': F_BODY},
          {'text': '$409,000', 'size': 12, 'color': ONLAV, 'font': F_BODY, 'bold': True}], first=True)
box, tf = tb(s, hx + 0.35, top + 1.05, hw - 0.7, 1.2)
para(tf, {'text': '$264,000', 'size': 52, 'color': WHITE, 'bold': True, 'font': F_HEAD}, first=True)
box, tf = tb(s, hx + 0.37, top + 2.15, hw - 0.7, 0.35)
para(tf, {'text': 'USD   ≈ £200,000 / year', 'size': 14, 'color': GOLD, 'bold': True, 'font': F_HEAD}, first=True)
box, tf = tb(s, hx + 0.37, top + 2.62, hw - 0.72, 1.2)
para(tf, {'text': 'All-in annual fee — technology subscription plus a dedicated forward-deployed AI expert team for year 1. Sign by 17 July to lock this rate.',
          'size': 11, 'color': ONLAV, 'font': F_BODY}, first=True, line=1.18)
bx = hx + hw + 0.4; bw = 13.333 - ML - bx
box, tf = tb(s, bx, top + 0.02, bw, 0.3)
para(tf, {'text': 'Price breakdown', 'size': 13, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, first=True)
bd = [('AI Expert Services', 'forward-deployed eng.', '$100,000', INK, False),
      ('Use Case', 'Engagement', '$175,000', INK, False),
      ('Audience entitlement', '4M-MAU tier', '$134,000', INK, False),
      ('List price', '', '$409,000', SLATE, False),
      ('Partnership discount (35%)', '', '–$145,000', GOLD, False),
      ('Year-one investment', '', '$264,000', PURPLE, True)]
ry = top + 0.42; rh = 0.55
for i, (lab, sub, val, col, strong) in enumerate(bd):
    last = (i == len(bd) - 1)
    if last: card(s, bx, ry, bw, rh + 0.04, fill=LAV, line_c=PURPLE, line_w=1.25, radius=0.10)
    box, tf = tb(s, bx + 0.18, ry + 0.07, bw - 1.9, rh - 0.1, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, [{'text': lab, 'size': 11.5 if not strong else 12.5, 'color': col,
               'bold': strong or last, 'font': F_HEAD}], first=True, line=1.0)
    if sub: para(tf, {'text': sub, 'size': 8.5, 'color': MUTE, 'font': F_BODY}, before=0)
    box, tf = tb(s, bx + bw - 1.85, ry + 0.07, 1.7, rh - 0.1, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, {'text': val, 'size': 13 if strong else 12, 'color': col, 'bold': True, 'font': F_HEAD},
         align=PP_ALIGN.RIGHT, first=True)
    if not last: rect(s, bx, ry + rh, bw, 0.012, LINE)
    ry += rh

# ====================== SLIDE 4 — Cost structure ======================
s = slide()
cy = header(s, 'Your cost structure, broken down',
            sub='One indivisible all-in fee — three components',
            kicker='08 · Commercial', tsize=23)
chrome(s, 4)
top = cy + 0.1
comp = [('1', 'AI Expert Services', 'Embedded ML experts who design, configure, tune and monitor your use case — at launch and ongoing.'),
        ('2', 'Use Case', 'One fully-configured, live RL use case optimised 1:1 across your audience. Additional cases are land-and-expand (~$300K each).'),
        ('3', 'Audience Entitlement', 'Annual entitlement on a monthly-average audience; flexes month to month and is revisited at renewal — not metered with immediate overages.')]
gap = 0.35; cwm = (CW - 2 * gap) / 3; ch = 2.95
for i, (num, ttl, desc) in enumerate(comp):
    x = ML + i * (cwm + gap)
    card(s, x, top, cwm, ch, fill=WHITE, line_c=LINE, line_w=1.25, radius=0.045)
    rect(s, x, top, cwm, 0.12, GOLD)
    shape(s, MSO_SHAPE.OVAL, x + 0.28, top + 0.35, 0.62, 0.62, fill=LAV, line_c=PURPLE, line_w=1.25)
    box, tf = tb(s, x + 0.28, top + 0.36, 0.62, 0.6, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, {'text': num, 'size': 22, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, align=PP_ALIGN.CENTER, first=True)
    box, tf = tb(s, x + 0.3, top + 1.12, cwm - 0.6, 0.6)
    para(tf, {'text': ttl, 'size': 15.5, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, first=True, line=1.0)
    box, tf = tb(s, x + 0.3, top + 1.62, cwm - 0.58, 1.2)
    para(tf, {'text': desc, 'size': 11, 'color': SLATE, 'font': F_BODY}, first=True, line=1.2)
ay = top + ch + 0.22
card(s, ML, ay, CW, 0.86, fill=LAV, line_c=LINE, line_w=1.0, radius=0.05)
rect(s, ML, ay, 0.1, 0.86, PURPLE)
box, tf = tb(s, ML + 0.3, ay + 0.12, CW - 0.6, 0.66, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [{'text': 'Commercial anchor:  ', 'size': 11.5, 'color': PURPLE, 'bold': True, 'font': F_HEAD},
          {'text': "a single, indivisible all-in fee per use case. BLC's 4.2M MAU sits at the 4M-MAU tier ($134K). Additional use cases are land-and-expand ~$300K each; audience entitlement is shared.",
           'size': 11.5, 'color': INK, 'font': F_BODY}], first=True, line=1.18)

# ====================== SLIDE 5 — AI Expert Services ======================
s = slide()
cy = header(s, 'What your AI Expert Services team does',
            sub='Forward-deployed ML experts embedded alongside your team — required at launch AND for the entire life of the use case.',
            tsize=23)
chrome(s, 5)
top = cy + 0.1; colw = (CW - 0.4) / 2
two = [('At launch', ['Design the use case: action bank, reward signal, exploration parameters',
                      'Set up the holdout and integrate your data feeds',
                      'Configure, train and tune the model; run pre-go-live testing',
                      'Stand up live reporting and agree the success metric']),
       ('Ongoing', ['Continuously tune the model as behaviour and seasonality shift',
                    'Monitor controls and dimensions; troubleshoot and course-correct',
                    'Add and test new dimensions, actions and creative to unlock fresh value',
                    'Quarterly reviews and expansion: value keeps compounding years in, not just at go-live'])]
ch = 3.35
for i, (ttl, bullets) in enumerate(two):
    x = ML + i * (colw + 0.4)
    card(s, x, top, colw, ch, fill=WHITE, line_c=LINE, line_w=1.25, radius=0.04)
    rect(s, x, top, 0.1, ch, GOLD if i else PURPLE)
    box, tf = tb(s, x + 0.32, top + 0.25, colw - 0.6, 0.4)
    para(tf, {'text': ttl, 'size': 15, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, first=True)
    box, tf = tb(s, x + 0.32, top + 0.78, colw - 0.62, ch - 1.0)
    for j, b in enumerate(bullets):
        para(tf, [{'text': '•  ', 'size': 12, 'color': GOLD, 'bold': True, 'font': F_BODY},
                  {'text': b, 'size': 11.5, 'color': SLATE, 'font': F_BODY}], first=(j == 0), line=1.16, after=8)
ay = top + ch + 0.2
box, tf = tb(s, ML, ay, CW, 0.5)
para(tf, [{'text': 'One use case, one annual fee.  ', 'size': 11, 'color': PURPLE, 'bold': True, 'font': F_HEAD},
          {'text': 'AI Expert Services, the Use Case and Audience are three SKUs but a single, indivisible all-in price required at launch and ongoing.',
           'size': 11, 'color': SLATE, 'font': F_BODY}], first=True, line=1.15)

# ====================== SLIDE 6 — Audience measured & billed ======================
s = slide()
cy = header(s, 'How your audience is measured & billed',
            sub='Annual entitlement, priced on a monthly average — designed to flex, not to penalise.',
            tsize=23)
chrome(s, 6)
top = cy + 0.1
two = [("How it's measured", ['Measured daily — audience size tracked every day, per use case',
                              'Calculated as a monthly average — averaged over the year to give the audience size',
                              'Charged annually — that average monthly audience is charged once a year; runs over and under with seasonality']),
       ("How it's billed", ['Annual entitlement — invoiced annually, not a monthly charge',
                            'Not like message credits — no immediate overage when a month runs high',
                            'Right-sized at renewal — we review and reset the tier together',
                            'Standard policy — overages (1.25× rate) only if you run materially above, over time'])]
ch = 3.35
for i, (ttl, bullets) in enumerate(two):
    x = ML + i * (colw + 0.4)
    card(s, x, top, colw, ch, fill=WHITE, line_c=LINE, line_w=1.25, radius=0.04)
    rect(s, x, top, 0.1, ch, PURPLE if i == 0 else GOLD)
    box, tf = tb(s, x + 0.32, top + 0.25, colw - 0.6, 0.4)
    para(tf, {'text': ttl, 'size': 15, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, first=True)
    box, tf = tb(s, x + 0.32, top + 0.78, colw - 0.62, ch - 1.0)
    for j, b in enumerate(bullets):
        para(tf, [{'text': '•  ', 'size': 12, 'color': GOLD, 'bold': True, 'font': F_BODY},
                  {'text': b, 'size': 11.5, 'color': SLATE, 'font': F_BODY}], first=(j == 0), line=1.16, after=7)
ay = top + ch + 0.2
box, tf = tb(s, ML, ay, CW, 0.5)
para(tf, [{'text': 'Built-in headroom:  ', 'size': 11, 'color': PURPLE, 'bold': True, 'font': F_HEAD},
          {'text': "your 4M tier sits at the audience level of BLC's 4.2M MAU base — comfortable for normal seasonality.",
           'size': 11, 'color': SLATE, 'font': F_BODY}], first=True, line=1.15)

# ====================== SLIDE 7 — Path to ROI (chart) ======================
s = slide()
cy = header(s, 'Engagement — path to ROI, month by month',
            sub='Cumulative incremental value clears break-even by May 2027, inside Year 1',
            kicker='Engagement', tsize=22)
chrome(s, 7)
months = ['Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct']
mcode = ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12']
cumv = [0, 0, 0, 34, 102, 205, 341, 477, 614, 750, 887, 1023]
roi = ['–', '–', '–', '0.1×', '0.4×', '0.8×', '1.3×', '1.8×', '2.3×', '2.8×', '3.4×', '3.9×']
valmo = ['$0', '$0', '$0', '$34k', '$68k', '$102k', '$136k', '$136k', '$136k', '$136k', '$136k', '$136k']
costmo = ['$22k'] * 12
base_y = 4.95; plot_x = ML + 0.05; plot_w = CW - 0.1; maxv = 1023.0; top_h = 2.30
col_pitch = plot_w / 12.0; bar_w = col_pitch * 0.52
be_y = base_y - (264.0 / maxv) * top_h
rect(s, plot_x, be_y, plot_w, 0.018, GOLD)
box, tf = tb(s, plot_x, be_y - 0.30, 4.0, 0.25)
para(tf, {'text': 'Break-even · $264k subscription', 'size': 9, 'color': GOLD, 'bold': True, 'font': F_HEAD}, first=True)
for i in range(12):
    cx = plot_x + i * col_pitch + (col_pitch - bar_w) / 2
    h = (cumv[i] / maxv) * top_h
    if cumv[i] == 0:
        rect(s, cx, base_y - 0.03, bar_w, 0.03, LINE)
    else:
        col = PURPLE if cumv[i] >= 264 else PERI
        rect(s, cx, base_y - h, bar_w, h, col)
        box, tf = tb(s, cx - 0.15, base_y - h - 0.26, bar_w + 0.3, 0.24)
        para(tf, {'text': roi[i], 'size': 8.5, 'color': INK, 'bold': True, 'font': F_HEAD}, align=PP_ALIGN.CENTER, first=True)
    box, tf = tb(s, cx - 0.2, base_y + 0.05, bar_w + 0.4, 0.42)
    para(tf, {'text': months[i], 'size': 9, 'color': INK, 'bold': True, 'font': F_BODY}, align=PP_ALIGN.CENTER, first=True)
    para(tf, {'text': mcode[i], 'size': 7.5, 'color': MUTE, 'font': F_BODY}, align=PP_ALIGN.CENTER, before=0)
rect(s, plot_x, base_y, plot_w, 0.012, LINE)
tdy = base_y + 0.62
data_rows = [('Value / mo', valmo, INK), ('ROI', roi, PURPLE), ('Cost / mo', costmo, SLATE)]
dt = s.shapes.add_table(4, 13, IN(ML), IN(tdy), IN(CW), IN(1.05)).table
dt.first_row = False; dt.horz_banding = False
dt.columns[0].width = IN(1.25)
for j in range(1, 13): dt.columns[j].width = IN((CW - 1.25) / 12)
for j, t in enumerate([''] + mcode):
    c = dt.cell(0, j); c.fill.solid(); c.fill.fore_color.rgb = PURPLE if j else WHITE
    c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.03); c.margin_right = IN(0.03); c.margin_top = IN(0.0); c.margin_bottom = IN(0.0)
    para(c.text_frame, {'text': t, 'size': 7.5, 'color': WHITE if j else INK, 'bold': True, 'font': F_HEAD}, align=PP_ALIGN.CENTER, first=True)
for ri, (lab, vals, col) in enumerate(data_rows, start=1):
    c = dt.cell(ri, 0); c.fill.solid(); c.fill.fore_color.rgb = LAV
    c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.06); c.margin_right = IN(0.02)
    para(c.text_frame, {'text': lab, 'size': 8, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, align=PP_ALIGN.LEFT, first=True)
    for j, v in enumerate(vals, start=1):
        c = dt.cell(ri, j); c.fill.solid(); c.fill.fore_color.rgb = WHITE if ri % 2 else PAPER
        c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.02); c.margin_right = IN(0.02); c.margin_top = IN(0.0); c.margin_bottom = IN(0.0)
        para(c.text_frame, {'text': v, 'size': 7.6, 'color': col, 'bold': (ri == 2), 'font': F_BODY}, align=PP_ALIGN.CENTER, first=True)
for r in range(4): dt.rows[r].height = IN(0.24)
footnote(s, 'Cumulative value in USD at the target lift ($1.64M annual run-rate ÷ 12 ≈ $136k/month). Audience ramps 25%→50%→75%→100% across Feb–May 2027. ROI vs $264k early-sign subscription; at $317k standard rate, year-end ROI is 3.2×. FX £1 = $1.32 (30 Jun 2026).', y=6.95)

# ====================== SLIDE 8 — ROI compounds ======================
s = slide()
cy = header(s, 'ROI compounds as the model trains',
            sub='Year-1 ramps to full run-rate; the subscription stays flat (Engagement · Target)', tsize=23)
chrome(s, 8)
top = cy + 0.12
years = [('Year 1  (ramp)', '$1.09M', '≈ £825,000', 'In-year 2027 value', 'Feb go-live, partial year', False),
         ('Year 2', '$1.64M', '≈ £1.24M', 'Full run-rate incremental', 'value over BAU', True),
         ('Year 3', '$1.64M', '≈ £1.24M', 'Full run-rate continues;', 'subscription unchanged', False)]
gap = 0.35; cwm = (CW - 2 * gap) / 3; ch = 2.55
for i, (yr, usd, gbp, l1, l2, hot) in enumerate(years):
    x = ML + i * (cwm + gap)
    card(s, x, top, cwm, ch, fill=(PURPLE if hot else WHITE), line_c=(None if hot else LINE), line_w=1.25, radius=0.05)
    if not hot: rect(s, x, top, cwm, 0.12, GOLD)
    box, tf = tb(s, x + 0.32, top + 0.3, cwm - 0.6, 0.35)
    para(tf, {'text': yr, 'size': 13, 'color': (GOLD if hot else PURPLE), 'bold': True, 'font': F_HEAD}, first=True)
    box, tf = tb(s, x + 0.3, top + 0.72, cwm - 0.6, 0.7)
    para(tf, {'text': usd, 'size': 34, 'color': (WHITE if hot else INK), 'bold': True, 'font': F_HEAD}, first=True)
    box, tf = tb(s, x + 0.33, top + 1.42, cwm - 0.6, 0.3)
    para(tf, {'text': gbp, 'size': 13, 'color': GOLD, 'bold': True, 'font': F_HEAD}, first=True)
    box, tf = tb(s, x + 0.33, top + 1.82, cwm - 0.62, 0.7)
    para(tf, {'text': l1, 'size': 10, 'color': (ONLAV if hot else SLATE), 'font': F_BODY}, first=True, line=1.12)
    para(tf, {'text': l2, 'size': 10, 'color': (ONLAV if hot else SLATE), 'font': F_BODY}, before=0, line=1.12)
ay = top + ch + 0.22
stats = [('Subscription · year 1', '$264,000'), ('Year 1 ROI', '4.1×'), ('Year 2 / 3 ROI', '6.2×')]
sw = (CW - 2 * gap) / 3
for i, (lab, val) in enumerate(stats):
    x = ML + i * (sw + gap)
    card(s, x, ay, sw, 0.92, fill=LAV, line_c=LINE, line_w=1.0, radius=0.06)
    box, tf = tb(s, x + 0.3, ay + 0.14, sw - 0.6, 0.66, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, [{'text': val + '   ', 'size': 19, 'color': PURPLE, 'bold': True, 'font': F_HEAD},
              {'text': lab, 'size': 10.5, 'color': SLATE, 'font': F_BODY}], first=True, line=1.0)
footnote(s, 'ROI = annual incremental value ÷ year-1 subscription. Shown at $264k early-sign rate. At $317k standard rate (after 17 July): Year 1 = 3.4×, Year 2/3 = 5.2×. Low case ($493k run-rate): Year 2 ROI 1.9× / 1.6×. RL models improve at maximising the reward over time.', y=6.95)

# ====================== SLIDE 9 — How we get to 2.3x ======================
s = slide()
cy = header(s, 'How we get to 2.3× ROI by July 2027',
            sub='ROI multiple = cumulative incremental value to date ÷ annual subscription. Build, month by month.', tsize=23)
chrome(s, 9)
top = cy + 0.1
card(s, ML, top, CW, 0.62, fill=LAV, line_c=LINE, line_w=1.0, radius=0.06)
box, tf = tb(s, ML + 0.3, top, CW - 0.6, 0.62, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [{'text': 'ROI multiple   =   ', 'size': 13, 'color': PURPLE, 'bold': True, 'font': F_HEAD},
          {'text': 'cumulative incremental value to date   ÷   annual subscription ($264,000)', 'size': 13, 'color': INK, 'font': F_BODY}], first=True)
m9 = ['Nov M1', 'Dec M2', 'Jan M3', 'Feb M4', 'Mar M5', 'Apr M6', 'May M7', 'Jun M8', 'Jul M9']
valm = ['$0', '$0', '$0', '$32k', '$65k', '$97k', '$129k', '$129k', '$129k']
cumm = ['$0', '$0', '$0', '$32k', '$97k', '$193k', '$323k', '$451k', '$580k']
ty = top + 0.85
nt = s.shapes.add_table(3, 10, IN(ML), IN(ty), IN(CW), IN(1.5)).table
nt.first_row = False; nt.horz_banding = False
nt.columns[0].width = IN(1.85)
for j in range(1, 10): nt.columns[j].width = IN((CW - 1.85) / 9)
for ri, (lab, vals, col) in enumerate([('Month', m9, INK), ('Value that month', valm, INK), ('Cumulative value', cumm, PURPLE)]):
    c = nt.cell(ri, 0); c.fill.solid(); c.fill.fore_color.rgb = PURPLE if ri == 0 else LAV
    c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.08)
    para(c.text_frame, {'text': lab, 'size': 9, 'color': WHITE if ri == 0 else PURPLE, 'bold': True, 'font': F_HEAD}, first=True)
    for j, v in enumerate(vals, start=1):
        c = nt.cell(ri, j); c.fill.solid(); c.fill.fore_color.rgb = PURPLE if ri == 0 else (WHITE if ri % 2 else PAPER)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.02); c.margin_right = IN(0.02)
        para(c.text_frame, {'text': v, 'size': 9, 'color': WHITE if ri == 0 else col, 'bold': (ri >= 2 or ri == 0), 'font': F_BODY}, align=PP_ALIGN.CENTER, first=True)
for r in range(3): nt.rows[r].height = IN(0.42)
ry = ty + 1.55
card(s, ML, ry, CW, 1.0, fill=PURPLE, radius=0.05)
box, tf = tb(s, ML, ry + 0.16, CW, 0.5, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [{'text': '$614,000', 'size': 23, 'color': WHITE, 'bold': True, 'font': F_HEAD},
          {'text': '   ÷   ', 'size': 20, 'color': ONLAV, 'font': F_HEAD},
          {'text': '$264,000', 'size': 23, 'color': WHITE, 'bold': True, 'font': F_HEAD},
          {'text': '   =   ', 'size': 20, 'color': ONLAV, 'font': F_HEAD},
          {'text': '2.3×', 'size': 26, 'color': GOLD, 'bold': True, 'font': F_HEAD}], align=PP_ALIGN.CENTER, first=True)
box, tf = tb(s, ML, ry + 0.66, CW, 0.3)
para(tf, {'text': 'cumulative value to July 2027 (month 9)   ÷   annual subscription', 'size': 10, 'color': ONLAV, 'font': F_BODY}, align=PP_ALIGN.CENTER, first=True)
footnote(s, 'Run-rate value is $136,000/month (≈£103,000) ($1.64M target ÷ 12 — engagement use case, 27 May model). ROI shown at $264k early-sign price; at $317k standard rate, July 2027 ROI is 1.9×. Audience ramps 25%/50%/75% over Feb–Apr 2027.', y=7.0)

# ====================== SLIDE 10 — Cumulative cost vs value + ask ======================
s = slide()
cy = header(s, 'Cumulative cost vs. value — and the ask',
            sub='Value compounds while the subscription accrues evenly — break-even May 2027, 3.9× year-end at the early-sign rate', tsize=22)
chrome(s, 10)
top = cy + 0.06
cumval = ['$0', '$0', '$0', '$34k', '$102k', '$205k', '$341k', '$477k', '$614k', '$750k', '$887k', '$1,023k']
cumcost = ['$22k', '$44k', '$66k', '$88k', '$110k', '$132k', '$154k', '$176k', '$198k', '$220k', '$242k', '$264k']
roirow = ['–', '–', '–', '0.1×', '0.4×', '0.8×', '1.3×', '1.8×', '2.3×', '2.8×', '3.4×', '3.9×']
mlbl = ['Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct']
t10 = s.shapes.add_table(4, 13, IN(ML), IN(top), IN(CW), IN(1.65)).table
t10.first_row = False; t10.horz_banding = False
t10.columns[0].width = IN(1.55)
for j in range(1, 13): t10.columns[j].width = IN((CW - 1.55) / 12)
for j, tx in enumerate([''] + [f'{a} {b}' for a, b in zip(mlbl, mcode)]):
    c = t10.cell(0, j); c.fill.solid(); c.fill.fore_color.rgb = PURPLE if j else WHITE
    c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.02); c.margin_right = IN(0.02); c.margin_top = IN(0.0); c.margin_bottom = IN(0.0)
    para(c.text_frame, {'text': tx, 'size': 6.6, 'color': WHITE if j else INK, 'bold': True, 'font': F_HEAD}, align=PP_ALIGN.CENTER, first=True)
for ri, (lab, vals, col) in enumerate([('Cum. value', cumval, PURPLE), ('Cum. cost', cumcost, SLATE), ('ROI ×', roirow, GOLD)], start=1):
    c = t10.cell(ri, 0); c.fill.solid(); c.fill.fore_color.rgb = LAV
    c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.06)
    para(c.text_frame, {'text': lab, 'size': 8, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, first=True)
    for j, v in enumerate(vals, start=1):
        c = t10.cell(ri, j); c.fill.solid(); c.fill.fore_color.rgb = WHITE if ri % 2 else PAPER
        c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.margin_left = IN(0.01); c.margin_right = IN(0.01); c.margin_top = IN(0.0); c.margin_bottom = IN(0.0)
        para(c.text_frame, {'text': v, 'size': 7.4, 'color': col, 'bold': (ri == 3 or ri == 1), 'font': F_BODY}, align=PP_ALIGN.CENTER, first=True)
for r in range(4): t10.rows[r].height = IN(0.34)
hy = top + 1.85
hl = [('BREAK-EVEN', 'May 2027', 'Month 7 — cum value crosses $264k'),
      ('YEAR-END', '3.9× ROI', '$1.02M cum value ÷ $264k subscription'),
      ('YEAR 2 / 3', '6.2× ROI', '$1.64M run-rate ÷ $264k subscription')]
gap = 0.3; hw = (CW - 2 * gap) / 3
for i, (k, big, sub) in enumerate(hl):
    x = ML + i * (hw + gap)
    card(s, x, hy, hw, 1.18, fill=WHITE, line_c=LINE, line_w=1.25, radius=0.06)
    rect(s, x, hy, 0.09, 1.18, GOLD if i == 0 else PURPLE)
    box, tf = tb(s, x + 0.28, hy + 0.16, hw - 0.5, 0.9)
    para(tf, {'text': k, 'size': 9, 'color': GOLD, 'bold': True, 'font': F_HEAD, 'spc': 120}, first=True)
    para(tf, {'text': big, 'size': 19, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, before=2)
    para(tf, {'text': sub, 'size': 9, 'color': SLATE, 'font': F_BODY}, before=2, line=1.05)
ax = hy + 1.32
card(s, ML, ax, CW, 0.74, fill=PURPLE, radius=0.05)
box, tf = tb(s, ML + 0.35, ax, CW - 0.7, 0.74, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [{'text': 'THE ASK TODAY    ', 'size': 12, 'color': GOLD, 'bold': True, 'font': F_HEAD, 'spc': 100},
          {'text': 'Sign order form by 17 July to lock $264,000 → November 2026 kickoff → February 2027 go-live.', 'size': 12.5, 'color': WHITE, 'bold': True, 'font': F_HEAD}], first=True, line=1.05)
footnote(s, 'All values in USD. Cum. value is incremental over BAU at the target lift (4.2M MAU × 20% × £3.00/conv × 5.1% lift = $1.64M annual run-rate). Cost shown linearly at $22k/mo. ROI at $264k early-sign; at $317k standard rate: year-end 3.2×, Year 2+ 5.2×. FX £1 = $1.32 (Morningstar, 30 Jun 2026).', y=6.05)

# ====================== SLIDE 11 — Closing ======================
s = slide()
rect(s, 0, 0, 13.333, 7.5, LAV)
rect(s, 0, 0, 0.16, 7.5, PURPLE)
rect(s, 0, 7.5 - 0.16, 13.333, 0.16, GOLD)
logo(s, h=0.5)
box, tf = tb(s, ML + 0.2, 2.85, 11.5, 1.4)
para(tf, {'text': 'Let’s lock the November slot.', 'size': 34, 'color': PURPLE, 'bold': True, 'font': F_HEAD}, first=True, line=1.0)
rect(s, ML + 0.22, 3.95, 2.2, 0.06, GOLD)
box, tf = tb(s, ML + 0.2, 4.2, 11.0, 0.5)
para(tf, {'text': 'Sign by 17 July to lock $264,000 → Feb 2027 go-live.', 'size': 15, 'color': SLATE, 'font': F_BODY}, first=True)
box, tf = tb(s, ML + 0.2, 6.75, 7.0, 0.35)
para(tf, {'text': 'BrazeAI Decisioning Studio™  ·  Confidential', 'size': 10.5, 'color': MUTE, 'font': F_BODY}, first=True)

out = os.path.join(HERE, "BLC_Commercial_Deck_Braze.pptx")
prs.save(out)
print("saved", out)
