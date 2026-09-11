"""Plot only Passed and Better value at its original CSV No.; no search."""
from pathlib import Path
import argparse
import csv
import hashlib
import json
import math
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.lines import Line2D

ROOT = Path(__file__).resolve().parent.parent
BETTER = 'Passed and Better value'
COLORS = {'BA':'#1b5d87', 'HCA':'#df5400'}

def read_series(path):
    raw = path.read_bytes()
    rows = list(csv.DictReader(raw.decode('utf-8-sig').splitlines()))
    if not rows or list(rows[0]) != ['No.', 'Rejected', 'Passed', BETTER]:
        raise ValueError('Unexpected CSV columns')
    if [int(r['No.']) for r in rows] != list(range(len(rows))):
        raise ValueError('Missing No.0 or non-consecutive rows; cannot invent an initial point')
    initial = [(c,float(rows[0][c])) for c in ('Rejected','Passed',BETTER) if rows[0][c].strip()]
    if len(initial) != 1:
        raise ValueError('Missing or ambiguous initial price')
    points = [(int(r['No.']),float(r[BETTER])) for r in rows if r[BETTER].strip()]
    if not points or any(points[i][1]>points[i-1][1] for i in range(1,len(points))):
        raise ValueError('Missing or non-monotone feasible-best history')
    if not all(math.isfinite(y) and y>0 for _,y in points+[(0,initial[0][1])]):
        raise ValueError('Invalid price')
    final = points[-1][1]
    return {'file':str(path.resolve()),'sha256':hashlib.sha256(raw).hexdigest(),
            'rows':len(rows),'initial_no':0,'initial_price':initial[0][1],
            'initial_column':initial[0][0],'points':points,'price':final,
            'first_feasible_no':points[0][0], 'best_no':points[-1][0]}

def plot(series, output, cut, resume):
    rows = series['BA']['rows']
    assert rows == series['HCA']['rows']
    last_no = rows-1
    budget = math.ceil(last_no/100)*100 if last_no>=100 else max(1,last_no)
    assert max(s['best_no'] for s in series.values()) < cut < resume < budget
    assert not any(cut<x<resume for s in series.values() for x,y in s['points'])
    font_manager.fontManager.addfont('C:/Windows/Fonts/tahoma.ttf')
    plt.rcParams.update({'font.family':'Tahoma','font.size':7.2,'axes.linewidth':.9})
    fig = plt.figure(figsize=(7/2.54,4/2.54))
    grid = fig.add_gridspec(1,2,width_ratios=[5,1],left=.24,right=.94,bottom=.225,top=.96,wspace=.12)
    left = fig.add_subplot(grid[0,0]); right = fig.add_subplot(grid[0,1],sharey=left)
    low = min(s['price'] for s in series.values())
    high = max(y for s in series.values() for _,y in s['points'])
    unit = 1000 if high > 1000 else 100
    upper = math.ceil(high/unit)*unit
    upper = max(upper, low + unit)
    lower = max(0,low-.055*(upper-low))
    for ax in (left,right):
        for method in ('BA','HCA'):
            s = series[method]; pts = s['points']+[(last_no,s['price'])]
            ax.step([x for x,y in pts],[y for x,y in pts],where='post',color=COLORS[method],lw=1.3,zorder=3)
        ax.set_ylim(lower,upper)
        ax.grid(True,linestyle=(0,(2,3)),lw=.45,color='#d3d7db'); ax.set_axisbelow(True)
        ax.spines[['top','right']].set_visible(False)
        ax.tick_params(length=2.8,width=.8,pad=2,labelsize=7.1); ax.margins(x=0)
    divisions = 5 if cut in (100, 500, 1500) else 4
    left.set_xlim(0,cut); left.set_xticks([cut*i/divisions for i in range(divisions+1)])
    left.set_xticklabels([f'{cut*i/divisions:g}' for i in range(divisions+1)])
    step = max(unit, round((upper-low)/4/unit)*unit)
    ticks = [low]+[upper-i*step for i in range(3,-1,-1) if upper-i*step>low]
    left.set_yticks(ticks); left.set_yticklabels([f'{low:,.2f}']+[f'{x:,.0f}' for x in ticks[1:]])
    left.set_ylabel('ราคารวม (บาท/ม.)',fontsize=8,labelpad=5)
    right.set_xlim(resume,budget); right.set_xticks([budget]); right.spines['left'].set_visible(False)
    right.tick_params(axis='y',left=False,labelleft=False)
    right.get_xticklabels()[0].set_horizontalalignment('right')
    for method,marker in [('BA','o'),('HCA','s')]:
        s = series[method]; x,y = s['best_no'],s['price']
        label_y = lower+(upper-lower)*.40
        left.plot(x,y,marker=marker,markersize=4.5,color=COLORS[method],zorder=6)
        left.vlines(x,y,label_y-(upper-lower)*.025,color=COLORS[method],linestyle=(0,(2,2)),lw=.8,zorder=4)
        label_x = x
        left.text(max(.16*cut,min(.82*cut,label_x)),label_y,f'{method}: {x:,}',ha='center',va='bottom',color=COLORS[method],fontsize=7.3)
    handles = [Line2D([0],[0],color=COLORS[m],lw=1.3,marker=mk,markersize=3.6,label=m) for m,mk in [('BA','o'),('HCA','s')]]
    fig.legend(handles=handles,loc='upper right',bbox_to_anchor=(.94,.958),frameon=False,ncol=2,
               handlelength=1.65,columnspacing=1.1,handletextpad=.5,fontsize=7.4)
    fig.text(.765,.62,f'ละช่วง {cut:,.0f}–{resume:,.0f} รอบ',ha='center',color='#555555',fontsize=6.5)
    for ax,xpos in ((left,1),(right,0)):
        for ypos in {0,*((s['price']-lower)/(upper-lower) for s in series.values())}:
            for color,width in [('white',2.4),('#333333',.95)]:
                ax.plot([xpos],[ypos],transform=ax.transAxes,marker=[(-1,-.9),(1,.9)],markersize=2.8,
                        linestyle='none',color=color,mew=width,clip_on=False,zorder=10 if color=='white' else 11)
    fig.text(.60,.047,'จำนวนรอบ',ha='center',fontsize=8)
    fig.canvas.draw()
    assert abs(left.transData.transform((0,0))[0]-left.bbox.x0)<1e-8
    output.parent.mkdir(parents=True,exist_ok=True)
    fig.savefig(output,dpi=600,facecolor='white'); plt.close(fig)

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--height',type=int,default=3); parser.add_argument('--fc',type=int,default=240)
    parser.add_argument('--data',type=Path,default=ROOT/'result_csv')
    parser.add_argument('--output',type=Path)
    parser.add_argument('--break-start',type=float); parser.add_argument('--break-end',type=float)
    args = parser.parse_args()
    out = ROOT/'audit/plot-from-accept-20260911'
    output = args.output or out/f'price_H{args.height}_{args.fc}_from_csv.png'
    series = {m:read_series(args.data/f'accept-{m}-H{args.height}-{args.fc}.csv') for m in ('BA','HCA')}
    latest = max(s['best_no'] for s in series.values())
    unit = 100 if latest<=400 else 500
    cut = args.break_start or (100 if latest<100 else 500 if latest<500 else math.ceil((latest+1)/unit)*unit)
    last_no = series['BA']['rows']-1
    axis_end = math.ceil(last_no/100)*100 if last_no>=100 else max(1,last_no)
    resume = args.break_end or axis_end-cut/5
    plot(series,output,cut,resume)
    for s in series.values():
        assert hashlib.sha256(Path(s['file']).read_bytes()).hexdigest()==s['sha256']
    print(output.resolve())
    print(json.dumps({m:{k:v for k,v in s.items() if k not in ('points','file','sha256')} for m,s in series.items()}))

if __name__=='__main__':
    main()
