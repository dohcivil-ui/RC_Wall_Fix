"""Render an existing VB6 result as real vectors; does not run or alter the design."""
from pathlib import Path
import argparse
import math
import re
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument('report', type=Path)
parser.add_argument('output', type=Path)
parser.add_argument('--width-cm', type=float, default=7)
parser.add_argument('--height-cm', type=float, default=6)
parser.add_argument('--total-trials', type=int, default=1)
args = parser.parse_args()
report = args.report.read_text()
if 'Validation: PASS_IMPLEMENTED_PROJECT_CHECKS' not in report:
    raise ValueError('A valid production result is required for this drawing')
if args.width_cm <= 0 or args.height_cm <= 0:
    raise ValueError('Physical image size must be positive')

def number(key):
    match = re.search(r'(?<![\w])' + re.escape(key) + r'=([-+\d.eE]+)', report)
    if not match:
        raise ValueError(key)
    return float(match[1])

H, H1, tt, tb, base_t, B, toe_l, heel_l, cover = [number(k) for k in
    ('H', 'H1', 'tt', 'tb', 'TBase', 'B', 'toe', 'heel', 'clear_cover')]
algorithm, trial = re.search(r'Algorithm=(\w+); Trial=(\d+)', report).groups()
if args.total_trials < int(trial):
    raise ValueError('Total trials must include the selected trial')
price = float(re.search(r'Concrete\+main-steel estimate=([\d.]+)', report)[1])
evaluation = int(number('BestEvaluation'))
bars = {}
for member in ('Stem', 'Toe', 'Heel'):
    db, spacing = re.search(member + r': DB(\d+) @ ([\d.]+) m', report).groups()
    bars[member] = int(db), float(spacing)

ET.register_namespace('', 'http://www.w3.org/2000/svg')
def tag(name):
    return '{http://www.w3.org/2000/svg}' + name

W, CH = 876, 640
# Include native-dialog-style chrome in the SVG, without rasterizing the drawing.
VW = max(W+28, (CH+88)*args.width_cm/args.height_cm)
VH = VW*args.height_cm/args.width_cm
svg = ET.Element(tag('svg'), {'width': f'{args.width_cm:g}cm', 'height': f'{args.height_cm:g}cm',
    'viewBox': f'0 0 {VW:g} {VH:g}', 'role': 'img', 'aria-labelledby': 'title description'})
ET.SubElement(svg, tag('title'), {'id': 'title'}).text = f'Best observed {algorithm} retaining wall popup, trial {trial}'
ET.SubElement(svg, tag('desc'), {'id': 'description'}).text = (
    'Static vector illustration of the VB6 popup, including title bar, OK button and close X. '
    'The controls in this image are illustrative; the actual VB6 controls remain functional. '
    'Drawing uses the saved native VB6 result, not a new optimization. '
    'Circular reinforcement markers are schematic; dimensions and spacing are in metres.')
def chrome(name, **attrs):
    return ET.SubElement(svg, tag(name), {k.replace('_','-'):str(v) for k,v in attrs.items()})
def chrome_text(x, y, label, anchor='start', size=12):
    t=chrome('text', x=x, y=y, text_anchor=anchor, font_family='Tahoma, Arial, sans-serif', font_size=size, fill='#111')
    t.text=label
    return t

defs=ET.SubElement(svg, tag('defs'))
grad=ET.SubElement(defs, tag('linearGradient'), {'id':'caption','x2':'0','y2':'1'})
ET.SubElement(grad, tag('stop'), {'offset':'0','stop-color':'#a5bfd7'})
ET.SubElement(grad, tag('stop'), {'offset':'1','stop-color':'#d5e1ed'})
chrome('rect', x=1,y=1,width=VW-2,height=VH-2,rx=5,fill='#f0f0f0',stroke='#4c667d',stroke_width=1)
chrome('rect', x=2,y=2,width=VW-4,height=25,rx=4,fill='url(#caption)')
chrome_text(10,18,'Best retaining wall - dimensions')
chrome('rect', x=VW-37,y=5,width=30,height=18,rx=3,fill='#c56851',stroke='#7c3b2d')
chrome('path', d=f'M {VW-27},9 l 10,10 M {VW-17},9 l -10,10',stroke='white',stroke_width=1.7,fill='none')
chrome('rect', x=14,y=38,width=VW-28,height=VH-88,fill='white')
chrome('rect', x=VW-118,y=VH-39,width=104,height=27,fill='#f8f8f8',stroke='#444')
chrome('rect', x=VW-115,y=VH-36,width=98,height=21,fill='none',stroke='#444',stroke_dasharray='1 1')
chrome_text(VW-66,VH-21,'OK','middle')
g = ET.SubElement(svg, tag('g'), {'transform': f'translate({(VW-W)/2:g},{38+(VH-88-CH)/2:g})',
    'stroke': '#232323', 'stroke-width': '1', 'fill': 'none', 'font-family': 'Tahoma, Arial, sans-serif'})

def element(name, **attrs):
    return ET.SubElement(g, tag(name), {k.replace('_', '-'): str(v) for k, v in attrs.items()})

def line(x1, y1, x2, y2, width=1, color='#232323'):
    element('line', x1=f'{x1:.4f}', y1=f'{y1:.4f}', x2=f'{x2:.4f}', y2=f'{y2:.4f}', stroke_width=width, stroke=color)

def text(x, y, label, anchor='middle', size=13.333, bold=False):
    e = element('text', x=f'{x:.4f}', y=f'{y+size:.4f}', text_anchor=anchor,
        font_size=size, font_weight='bold' if bold else 'normal', fill='#232728', stroke='none')
    e.text = label

def horizontal(left, right, face, y, label):
    tip = min(5, (right-left)/3)
    line(left, face, left, y+5, color='#6e6e64')
    line(right, face, right, y+5, color='#6e6e64')
    line(left, y, right, y)
    for end, direction in ((left, 1), (right, -1)):
        line(end+tip*direction, y-3, end, y)
        line(end, y, end+tip*direction, y+3)
    text((left+right)/2, y-19, label)

def vertical(top, bottom, face, x, label, right=False):
    tip = min(5, (bottom-top)/3)
    line(face, top, x, top, color='#6e6e64')
    line(face, bottom, x, bottom, color='#6e6e64')
    line(x, top, x, bottom)
    for end, direction in ((top, 1), (bottom, -1)):
        line(x-3, end+tip*direction, x, end)
        line(x, end, x+3, end+tip*direction)
    text(x+8 if right else x-8, (top+bottom-16)/2, label, 'start' if right else 'end')

def dots(ax, ay, bx, by):
    intervals = max(2, 2*round(math.hypot(bx-ax, by-ay)/44))
    for i in range(intervals+1):
        element('circle', cx=f'{ax+(bx-ax)*i/intervals:.4f}', cy=f'{ay+(by-ay)*i/intervals:.4f}',
            r='3', fill='black', stroke='black')

def callout(x, y, member, title, tx, ty, point_right=False):
    db, spacing = bars[member]
    label = f'DB{db} @ {spacing:.2f} m'
    text(x, y, title, 'start', bold=True)
    text(x, y+18, label, 'start')
    sx, sy = x+115 if point_right else x-8, y+25
    length = math.hypot(tx-sx, ty-sy)
    ux, uy = (tx-sx)/length, (ty-sy)/length
    line(sx, sy, tx, ty)
    line(tx-8*ux+3*uy, ty-8*uy-3*ux, tx, ty)
    line(tx, ty, tx-8*ux-3*uy, ty-8*uy+3*ux)

text(W/2, 14, f'FINAL RESULT  |  {algorithm}  |  Selected trial {trial} / {args.total_trials}', size=20, bold=True)
text(W/2, 44, f'COST  {price:,.2f} Baht/m  (concrete + main steel)', bold=True)
text(W/2, 69, f'Best cost first found at evaluation {evaluation}  |  Dimensions in metres')
scale = min((CH-275)/H, (W-440)/B)
x0 = (W-B*scale)/2
x1, toe, back = x0+B*scale, x0+toe_l*scale, x0+(toe_l+tb)*scale
top = back-tt*scale
y0 = CH-135
y_top, y_base, y_front = y0-H*scale, y0-base_t*scale, y0-H1*scale
line(x0-30, y_front, top+(toe-top)*(y_front-y_top)/(y_base-y_top), y_front, 2)
line(back, y_top, x1+20, y_top, 2)
points = [(x0,y0),(x0,y_base),(toe,y_base),(top,y_top),(back,y_top),(back,y_base),(x1,y_base),(x1,y0)]
element('polygon', points=' '.join(f'{x:.4f},{y:.4f}' for x,y in points), stroke_width='2')
ec = cover*scale
stem_x = back-(cover+bars['Stem'][0]/2000)*scale
toe_y = y0-(cover+bars['Toe'][0]/2000)*scale
heel_y = y_base+(cover+bars['Heel'][0]/2000)*scale
dots(stem_x, y_top+ec, stem_x, y_base-4)
dots(x0+ec, toe_y, toe-4, toe_y)
dots(back+4, heel_y, x1-ec, heel_y)
callout(x1+65, y_top+75, 'Stem', 'Stem - back face', stem_x, (y_top+ec+y_base-4)/2)
callout(x0-110, y_base-130, 'Toe', 'Toe - bottom', (x0+ec+toe-4)/2, toe_y, True)
callout(x1+65, y_base-80, 'Heel', 'Heel - top', (back+4+x1-ec)/2, heel_y)
horizontal(top, back, y_top, y_top-22, f'tt = {tt:.2f}')
horizontal(x0, toe, y0, y0+28, f'LToe = {toe_l:.2f}')
horizontal(toe, back, y0, y0+60, f'tb = {tb:.2f}')
horizontal(back, x1, y0, y0+28, f'LHeel = {heel_l:.2f}')
horizontal(x0, x1, y0, y0+92, f'Base = {B:.2f}')
vertical(y_top, y0, x0, x0-140, f'H = {H:.2f}')
vertical(y_front, y0, x0-30, x0-48, f'H1 = {H1:.2f}')
vertical(y_base, y0, x1, x1+26, f'TBase = {base_t:.2f}', True)
args.output.parent.mkdir(parents=True, exist_ok=True)
ET.indent(svg)
ET.ElementTree(svg).write(args.output, encoding='utf-8', xml_declaration=True)
assert not list(svg.iter(tag('image'))), 'SVG must contain vectors, not an embedded bitmap'
print(args.output)
