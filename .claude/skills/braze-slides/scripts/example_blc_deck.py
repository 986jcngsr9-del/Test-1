# -*- coding: utf-8 -*-
"""
example_blc_deck.py — canonical worked example for the braze-slides skill.

Shows both paths:
  • archetype builders (title/table/cards/two_column/closing) for standard slides
  • blank() + primitives for bespoke slides that still read as on-brand

Run:  python3 example_blc_deck.py [out.pptx]
This rebuilds the 11-slide BLC commercial deck in Braze styling. Use it as a
template: copy a slide block, swap the content, keep the structure.
"""
import sys
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as _RAW   # raw inches for bespoke tables (apply *0.75 scale)
from braze_deck import (BrazeDeck, ML, CW, DW,
                        PURPLE, PURPLE_D, GOLD, LAV, PERI, INK, SLATE, MUTE,
                        LINE, ONLAV, WHITE, PAPER)

d = BrazeDeck()

# 1 — title (archetype)
d.title("Commercials for Blue Light Card", kicker="AI Decisioning",
        subtitle="Year-one investment, cost structure & path to ROI")

# 2 — milestone table (archetype with status coloring)
def status_color(ri, ci, val):
    if ci != 4: return None
    if val.startswith('✓'): return PURPLE
    if 'In flight' in val: return GOLD
    if 'Today' in val: return PURPLE_D
    return MUTE
rows2 = [
    ['1', 'Priority use case identified — Engagement (absorbs re-engagement + cross-cat)', 'BLC + Braze', '27 May', '✓ Completed'],
    ['2', 'Business case & impact model built (£4.37M target run-rate)', 'Braze', '27 May', '✓ Completed'],
    ['3', 'Data validation & audience sizing (4.2M MAU)', 'BLC + Braze', '27 May', '✓ Completed'],
    ['4', 'Customer Readiness Assessment (CRA) — in-flight', 'Braze', '16 Jun', '● In flight'],
    ['5', 'Commercial proposal & cost structure shared', 'Braze', '17 Jun', '● Today'],
    ['6', 'Lifecycle team review model & price → internal alignment', 'BLC (Marie)', 'w/c 23 Jun', 'Upcoming'],
    ['7', 'Executive review — Director of Data, CPO, Head of Engineering', 'BLC + Braze', 'w/c 23 Jun', 'Upcoming'],
    ['8', 'Finance approval (if incremental budget required)', 'BLC', 'early Jul', 'Upcoming'],
    ['9', 'Go / No-Go decision', 'BLC', 'by mid-Jul', 'Upcoming'],
    ['10', 'Order form finalised & signature', 'BLC + Braze', 'late Jul', 'Upcoming'],
    ['11', 'Implementation kickoff', 'Braze + BLC', 'November 2026', 'Upcoming'],
    ['12', 'Go live — Engagement engine in production', 'Braze + BLC', 'February 2027', 'Upcoming'],
]
d.table('Mutual action plan to go-live',
        ['#', 'Milestone', 'Owner', 'Target', 'Status'], rows2,
        sub='Foundations are complete. Next is commercial alignment and executive sign-off to lock the Nov 2026 implementation slot.',
        col_widths=[0.5, 6.43, 1.7, 1.5, 2.0], cell_color=status_color, page=2,
        note="Note: BLC's current Braze contract was renewed Jan 2026 ($1M+). Decisioning Studio is net-new on top of that contract.")

# 3 — year-one investment (BESPOKE: price hero + breakdown)
s = d.blank()
cy = d.header(s, 'Year-one investment — Engagement proof of value',
              sub='Priced at the 4M-MAU tier — sign by 17 July to lock the early-sign rate',
              kicker='Commercials', tsize=22, page=3)
top = cy + 0.05; hx, hw = ML, 6.05
d.card(s, hx, top, hw, 3.95, fill=PURPLE, line_c=None)
_, tf = d.box(s, hx + 0.35, top + 0.3, hw - 0.7, 0.4)
d.para(tf, {'text': '35% PARTNERSHIP DISCOUNT', 'size': 11, 'color': GOLD, 'bold': True, 'spc': 120}, first=True)
_, tf = d.box(s, hx + 0.35, top + 0.72, hw - 0.7, 0.35)
d.para(tf, [{'text': 'List price  ', 'size': 12, 'color': ONLAV},
            {'text': '$409,000', 'size': 12, 'color': ONLAV, 'bold': True}], first=True)
_, tf = d.box(s, hx + 0.35, top + 1.05, hw - 0.7, 1.2)
d.para(tf, {'text': '$264,000', 'size': 52, 'color': WHITE, 'bold': True}, first=True)
_, tf = d.box(s, hx + 0.37, top + 2.15, hw - 0.7, 0.35)
d.para(tf, {'text': 'USD   ≈ £200,000 / year', 'size': 14, 'color': GOLD, 'bold': True}, first=True)
_, tf = d.box(s, hx + 0.37, top + 2.62, hw - 0.72, 1.2)
d.para(tf, {'text': 'All-in annual fee — technology subscription plus a dedicated forward-deployed AI expert team for year 1. Sign by 17 July to lock this rate.',
            'size': 11, 'color': ONLAV}, first=True, line=1.18)
bx = hx + hw + 0.4; bw = DW - ML - bx
_, tf = d.box(s, bx, top + 0.02, bw, 0.3)
d.para(tf, {'text': 'Price breakdown', 'size': 13, 'color': PURPLE, 'bold': True}, first=True)
bd = [('AI Expert Services', 'forward-deployed eng.', '$100,000', INK, False),
      ('Use Case', 'Engagement', '$175,000', INK, False),
      ('Audience entitlement', '4M-MAU tier', '$134,000', INK, False),
      ('List price', '', '$409,000', SLATE, False),
      ('Partnership discount (35%)', '', '–$145,000', GOLD, False),
      ('Year-one investment', '', '$264,000', PURPLE, True)]
ry = top + 0.42; rh = 0.55
for i, (lab, sub, val, col, strong) in enumerate(bd):
    last = (i == len(bd) - 1)
    if last: d.card(s, bx, ry, bw, rh + 0.04, fill=LAV, line_c=PURPLE, radius=0.10)
    _, tf = d.box(s, bx + 0.18, ry + 0.07, bw - 1.9, rh - 0.1, anchor=MSO_ANCHOR.MIDDLE)
    d.para(tf, [{'text': lab, 'size': 12.5 if strong else 11.5, 'color': col, 'bold': strong or last}], first=True, line=1.0)
    if sub: d.para(tf, {'text': sub, 'size': 8.5, 'color': MUTE}, before=0)
    _, tf = d.box(s, bx + bw - 1.85, ry + 0.07, 1.7, rh - 0.1, anchor=MSO_ANCHOR.MIDDLE)
    d.para(tf, {'text': val, 'size': 13 if strong else 12, 'color': col, 'bold': True}, align=PP_ALIGN.RIGHT, first=True)
    if not last: d.divider(s, bx, ry + rh, bw)
    ry += rh

# 4 — cost structure (archetype cards + anchor)
d.cards('Your cost structure, broken down',
        [('1', 'AI Expert Services', 'Embedded ML experts who design, configure, tune and monitor your use case — at launch and ongoing.'),
         ('2', 'Use Case', 'One fully-configured, live RL use case optimised 1:1 across your audience. Additional cases are land-and-expand (~$300K each).'),
         ('3', 'Audience Entitlement', 'Annual entitlement on a monthly-average audience; flexes month to month and is revisited at renewal — not metered with immediate overages.')],
        sub='One indivisible all-in fee — three components', kicker='08 · Commercial', page=4,
        anchor=("Commercial anchor:", "a single, indivisible all-in fee per use case. BLC's 4.2M MAU sits at the 4M-MAU tier ($134K). Additional use cases are land-and-expand ~$300K each; audience entitlement is shared."))

# 5 — AI Expert Services (archetype two_column)
d.two_column('What your AI Expert Services team does',
             ('At launch', ['Design the use case: action bank, reward signal, exploration parameters',
                            'Set up the holdout and integrate your data feeds',
                            'Configure, train and tune the model; run pre-go-live testing',
                            'Stand up live reporting and agree the success metric']),
             ('Ongoing', ['Continuously tune the model as behaviour and seasonality shift',
                          'Monitor controls and dimensions; troubleshoot and course-correct',
                          'Add and test new dimensions, actions and creative to unlock fresh value',
                          'Quarterly reviews and expansion: value keeps compounding years in, not just at go-live']),
             sub='Forward-deployed ML experts embedded alongside your team — required at launch AND for the entire life of the use case.', page=5,
             note=[{'text': 'One use case, one annual fee.  ', 'size': 11, 'color': PURPLE, 'bold': True},
                   {'text': 'AI Expert Services, the Use Case and Audience are three SKUs but a single, indivisible all-in price required at launch and ongoing.', 'size': 11, 'color': SLATE}])

# 6 — audience measured & billed (archetype two_column)
d.two_column('How your audience is measured & billed',
             ("How it's measured", ['Measured daily — audience size tracked every day, per use case',
                                    'Calculated as a monthly average — averaged over the year to give the audience size',
                                    'Charged annually — that average monthly audience is charged once a year; runs over and under with seasonality']),
             ("How it's billed", ['Annual entitlement — invoiced annually, not a monthly charge',
                                  'Not like message credits — no immediate overage when a month runs high',
                                  'Right-sized at renewal — we review and reset the tier together',
                                  'Standard policy — overages (1.25× rate) only if you run materially above, over time']),
             sub='Annual entitlement, priced on a monthly average — designed to flex, not to penalise.', page=6,
             note=[{'text': 'Built-in headroom:  ', 'size': 11, 'color': PURPLE, 'bold': True},
                   {'text': "your 4M tier sits at the audience level of BLC's 4.2M MAU base — comfortable for normal seasonality.", 'size': 11, 'color': SLATE}])

# 7 — path to ROI (BESPOKE: bar chart + data table)
s = d.blank()
d.header(s, 'Engagement — path to ROI, month by month',
         sub='Cumulative incremental value clears break-even by May 2027, inside Year 1', kicker='Engagement', tsize=22, page=7)
months = ['Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct']
mcode = ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12']
cumv = [0, 0, 0, 34, 102, 205, 341, 477, 614, 750, 887, 1023]
roi = ['–', '–', '–', '0.1×', '0.4×', '0.8×', '1.3×', '1.8×', '2.3×', '2.8×', '3.4×', '3.9×']
valmo = ['$0', '$0', '$0', '$34k', '$68k', '$102k', '$136k', '$136k', '$136k', '$136k', '$136k', '$136k']
costmo = ['$22k'] * 12
base_y = 4.9; px = ML + 0.05; pw = CW - 0.1; mx = 1023.0; top_h = 2.25
pitch = pw / 12; bw = pitch * 0.52; be_y = base_y - (264.0 / mx) * top_h
d.rect(s, px, be_y, pw, 0.018, GOLD)
_, tf = d.box(s, px, be_y - 0.30, 4.0, 0.25)
d.para(tf, {'text': 'Break-even · $264k subscription', 'size': 9, 'color': GOLD, 'bold': True}, first=True)
for i in range(12):
    cx = px + i * pitch + (pitch - bw) / 2; h = (cumv[i] / mx) * top_h
    if cumv[i] == 0:
        d.rect(s, cx, base_y - 0.03, bw, 0.03, LINE)
    else:
        d.rect(s, cx, base_y - h, bw, h, PURPLE if cumv[i] >= 264 else PERI)
        _, tf = d.box(s, cx - 0.15, base_y - h - 0.26, bw + 0.3, 0.24)
        d.para(tf, {'text': roi[i], 'size': 8.5, 'color': INK, 'bold': True}, align=PP_ALIGN.CENTER, first=True)
    _, tf = d.box(s, cx - 0.2, base_y + 0.05, bw + 0.4, 0.42)
    d.para(tf, {'text': months[i], 'size': 9, 'color': INK, 'bold': True}, align=PP_ALIGN.CENTER, first=True)
    d.para(tf, {'text': mcode[i], 'size': 7.5, 'color': MUTE}, align=PP_ALIGN.CENTER, before=0)
d.divider(s, px, base_y, pw)
tdy = base_y + 0.62
dt = s.shapes.add_table(4, 13, _RAW(ML * 0.75), _RAW(tdy * 0.75), _RAW(CW * 0.75), _RAW(1.05 * 0.75)).table
dt.first_row = False; dt.horz_banding = False
dt.columns[0].width = _RAW(1.25 * 0.75)
for j in range(1, 13): dt.columns[j].width = _RAW(((CW - 1.25) / 12) * 0.75)
def _cell(tbl, ri, ci, txt, size, color, fill, bold=False, al=PP_ALIGN.CENTER):
    c = tbl.cell(ri, ci); c.fill.solid(); c.fill.fore_color.rgb = fill
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.margin_left = _RAW(0.03 * 0.75); c.margin_right = _RAW(0.03 * 0.75)
    c.margin_top = _RAW(0); c.margin_bottom = _RAW(0)
    d.para(c.text_frame, {'text': txt, 'size': size, 'color': color, 'bold': bold}, align=al, first=True)
for j, t in enumerate([''] + mcode):
    _cell(dt, 0, j, t, 7.5, WHITE if j else INK, PURPLE if j else WHITE, bold=True)
for ri, (lab, vals, col) in enumerate([('Value / mo', valmo, INK), ('ROI', roi, PURPLE), ('Cost / mo', costmo, SLATE)], start=1):
    _cell(dt, ri, 0, lab, 8, PURPLE, LAV, bold=True, al=PP_ALIGN.LEFT)
    for j, v in enumerate(vals, start=1):
        _cell(dt, ri, j, v, 7.6, col, WHITE if ri % 2 else PAPER, bold=(ri == 2))
for r in range(4): dt.rows[r].height = _RAW(0.24 * 0.75)
d.footnote(s, 'Cumulative value in USD at the target lift ($1.64M annual run-rate ÷ 12 ≈ $136k/month). Audience ramps 25%→50%→75%→100% across Feb–May 2027. ROI vs $264k early-sign subscription; at $317k standard rate, year-end ROI is 3.2×. FX £1 = $1.32 (30 Jun 2026).', y=6.9)

# 8 — ROI compounds (BESPOKE: year cards + stat band)
s = d.blank()
cy = d.header(s, 'ROI compounds as the model trains',
              sub='Year-1 ramps to full run-rate; the subscription stays flat (Engagement · Target)', page=8)
top = cy + 0.12; gap = 0.35; cwm = (CW - 2 * gap) / 3; ch = 2.55
years = [('Year 1  (ramp)', '$1.09M', '≈ £825,000', 'In-year 2027 value', 'Feb go-live, partial year', False),
         ('Year 2', '$1.64M', '≈ £1.24M', 'Full run-rate incremental', 'value over BAU', True),
         ('Year 3', '$1.64M', '≈ £1.24M', 'Full run-rate continues;', 'subscription unchanged', False)]
for i, (yr, usd, gbp, l1, l2, hot) in enumerate(years):
    x = ML + i * (cwm + gap)
    d.card(s, x, top, cwm, ch, fill=(PURPLE if hot else WHITE), line_c=(None if hot else LINE))
    if not hot: d.rect(s, x, top, cwm, 0.12, GOLD)
    _, tf = d.box(s, x + 0.32, top + 0.3, cwm - 0.6, 0.35)
    d.para(tf, {'text': yr, 'size': 13, 'color': GOLD if hot else PURPLE, 'bold': True}, first=True)
    _, tf = d.box(s, x + 0.3, top + 0.72, cwm - 0.6, 0.7)
    d.para(tf, {'text': usd, 'size': 34, 'color': WHITE if hot else INK, 'bold': True}, first=True)
    _, tf = d.box(s, x + 0.33, top + 1.42, cwm - 0.6, 0.3)
    d.para(tf, {'text': gbp, 'size': 13, 'color': GOLD, 'bold': True}, first=True)
    _, tf = d.box(s, x + 0.33, top + 1.82, cwm - 0.62, 0.7)
    d.para(tf, {'text': l1, 'size': 10, 'color': ONLAV if hot else SLATE}, first=True, line=1.12)
    d.para(tf, {'text': l2, 'size': 10, 'color': ONLAV if hot else SLATE}, before=0, line=1.12)
d.stat_row(None, [('$264,000', 'Subscription · year 1'), ('4.1×', 'Year 1 ROI'), ('6.2×', 'Year 2 / 3 ROI')],
           slide=s, y=top + ch + 0.22)
d.footnote(s, 'ROI = annual incremental value ÷ year-1 subscription. Shown at $264k early-sign rate. At $317k standard rate (after 17 July): Year 1 = 3.4×, Year 2/3 = 5.2×. Low case ($493k run-rate): Year 2 ROI 1.9× / 1.6×. RL models improve at maximising the reward over time.', y=6.95)

# 9 — how we get to 2.3x (BESPOKE: formula + table + equation)
s = d.blank()
cy = d.header(s, 'How we get to 2.3× ROI by July 2027',
              sub='ROI multiple = cumulative incremental value to date ÷ annual subscription. Build, month by month.', page=9)
top = cy + 0.1
d.card(s, ML, top, CW, 0.62, fill=LAV, line_w=1.0)
_, tf = d.box(s, ML + 0.3, top, CW - 0.6, 0.62, anchor=MSO_ANCHOR.MIDDLE)
d.para(tf, [{'text': 'ROI multiple   =   ', 'size': 13, 'color': PURPLE, 'bold': True},
            {'text': 'cumulative incremental value to date   ÷   annual subscription ($264,000)', 'size': 13, 'color': INK}], first=True)
m9 = ['Nov M1', 'Dec M2', 'Jan M3', 'Feb M4', 'Mar M5', 'Apr M6', 'May M7', 'Jun M8', 'Jul M9']
valm = ['$0', '$0', '$0', '$32k', '$65k', '$97k', '$129k', '$129k', '$129k']
cumm = ['$0', '$0', '$0', '$32k', '$97k', '$193k', '$323k', '$451k', '$580k']
ty = top + 0.85
nt = s.shapes.add_table(3, 10, _RAW(ML * 0.75), _RAW(ty * 0.75), _RAW(CW * 0.75), _RAW(1.5 * 0.75)).table
nt.first_row = False; nt.horz_banding = False
nt.columns[0].width = _RAW(1.85 * 0.75)
for j in range(1, 10): nt.columns[j].width = _RAW(((CW - 1.85) / 9) * 0.75)
for ri, (lab, vals, col) in enumerate([('Month', m9, INK), ('Value that month', valm, INK), ('Cumulative value', cumm, PURPLE)]):
    _cell(nt, ri, 0, lab, 9, WHITE if ri == 0 else PURPLE, PURPLE if ri == 0 else LAV, bold=True, al=PP_ALIGN.LEFT)
    for j, v in enumerate(vals, start=1):
        _cell(nt, ri, j, v, 9, WHITE if ri == 0 else col, PURPLE if ri == 0 else (WHITE if ri % 2 else PAPER), bold=(ri >= 2 or ri == 0))
for r in range(3): nt.rows[r].height = _RAW(0.42 * 0.75)
ry = ty + 1.55
d.card(s, ML, ry, CW, 1.0, fill=PURPLE, line_c=None)
_, tf = d.box(s, ML, ry + 0.16, CW, 0.5, anchor=MSO_ANCHOR.MIDDLE)
d.para(tf, [{'text': '$614,000', 'size': 23, 'color': WHITE, 'bold': True},
            {'text': '   ÷   ', 'size': 20, 'color': ONLAV}, {'text': '$264,000', 'size': 23, 'color': WHITE, 'bold': True},
            {'text': '   =   ', 'size': 20, 'color': ONLAV}, {'text': '2.3×', 'size': 26, 'color': GOLD, 'bold': True}],
       align=PP_ALIGN.CENTER, first=True)
_, tf = d.box(s, ML, ry + 0.66, CW, 0.3)
d.para(tf, {'text': 'cumulative value to July 2027 (month 9)   ÷   annual subscription', 'size': 10, 'color': ONLAV}, align=PP_ALIGN.CENTER, first=True)
d.footnote(s, 'Run-rate value is $136,000/month (≈£103,000) ($1.64M target ÷ 12 — engagement use case, 27 May model). ROI shown at $264k early-sign price; at $317k standard rate, July 2027 ROI is 1.9×. Audience ramps 25%/50%/75% over Feb–Apr 2027.', y=7.0)

# 10 — cumulative cost vs value + ask (BESPOKE: table + highlights + callout)
s = d.blank()
cy = d.header(s, 'Cumulative cost vs. value — and the ask',
              sub='Value compounds while the subscription accrues evenly — break-even May 2027, 3.9× year-end at the early-sign rate', tsize=22, page=10)
top = cy + 0.06
cumval = ['$0', '$0', '$0', '$34k', '$102k', '$205k', '$341k', '$477k', '$614k', '$750k', '$887k', '$1,023k']
cumcost = ['$22k', '$44k', '$66k', '$88k', '$110k', '$132k', '$154k', '$176k', '$198k', '$220k', '$242k', '$264k']
roirow = ['–', '–', '–', '0.1×', '0.4×', '0.8×', '1.3×', '1.8×', '2.3×', '2.8×', '3.4×', '3.9×']
mlbl = ['Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct']
t10 = s.shapes.add_table(4, 13, _RAW(ML * 0.75), _RAW(top * 0.75), _RAW(CW * 0.75), _RAW(1.65 * 0.75)).table
t10.first_row = False; t10.horz_banding = False
t10.columns[0].width = _RAW(1.55 * 0.75)
for j in range(1, 13): t10.columns[j].width = _RAW(((CW - 1.55) / 12) * 0.75)
for j, tx in enumerate([''] + [f'{a} {b}' for a, b in zip(mlbl, mcode)]):
    _cell(t10, 0, j, tx, 6.6, WHITE if j else INK, PURPLE if j else WHITE, bold=True)
for ri, (lab, vals, col) in enumerate([('Cum. value', cumval, PURPLE), ('Cum. cost', cumcost, SLATE), ('ROI ×', roirow, GOLD)], start=1):
    _cell(t10, ri, 0, lab, 8, PURPLE, LAV, bold=True, al=PP_ALIGN.LEFT)
    for j, v in enumerate(vals, start=1):
        _cell(t10, ri, j, v, 7.4, col, WHITE if ri % 2 else PAPER, bold=(ri == 3 or ri == 1))
for r in range(4): t10.rows[r].height = _RAW(0.34 * 0.75)
hy = top + 1.85; gap = 0.3; hw = (CW - 2 * gap) / 3
hl = [('BREAK-EVEN', 'May 2027', 'Month 7 — cum value crosses $264k'),
      ('YEAR-END', '3.9× ROI', '$1.02M cum value ÷ $264k subscription'),
      ('YEAR 2 / 3', '6.2× ROI', '$1.64M run-rate ÷ $264k subscription')]
for i, (k, big, sub) in enumerate(hl):
    x = ML + i * (hw + gap)
    d.card(s, x, hy, hw, 1.18, fill=WHITE)
    d.rect(s, x, hy, 0.09, 1.18, GOLD if i == 0 else PURPLE)
    _, tf = d.box(s, x + 0.28, hy + 0.16, hw - 0.5, 0.9)
    d.para(tf, {'text': k, 'size': 9, 'color': GOLD, 'bold': True, 'spc': 120}, first=True)
    d.para(tf, {'text': big, 'size': 19, 'color': PURPLE, 'bold': True}, before=2)
    d.para(tf, {'text': sub, 'size': 9, 'color': SLATE}, before=2, line=1.05)
d.callout(s, ML, hy + 1.32, CW, 0.74, 'THE ASK TODAY',
          'Sign order form by 17 July to lock $264,000 → November 2026 kickoff → February 2027 go-live.')
d.footnote(s, 'All values in USD. Cum. value is incremental over BAU at the target lift (4.2M MAU × 20% × £3.00/conv × 5.1% lift = $1.64M annual run-rate). Cost shown linearly at $22k/mo. ROI at $264k early-sign; at $317k standard rate: year-end 3.2×, Year 2+ 5.2×. FX £1 = $1.32 (Morningstar, 30 Jun 2026).', y=6.05)

# 11 — closing (archetype)
d.closing("Let’s lock the November slot.",
          subline="Sign by 17 July to lock $264,000 → Feb 2027 go-live.")

out = sys.argv[1] if len(sys.argv) > 1 else "BLC_Commercial_Deck_Braze.pptx"
d.save(out)
print("saved", out)
