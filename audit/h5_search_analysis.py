"""Read the latest user sessions without modifying their CSV files."""
from pathlib import Path
import csv, collections, json, re
ROOT=Path(__file__).resolve().parent.parent
out=ROOT/'audit'/'h5-search'
out.mkdir(exist_ok=True)
sessions={}
for p in sorted((ROOT/'result_csv').glob('*/trial-summary.csv'),key=lambda p:p.stat().st_mtime,reverse=True):
    report=(p.parent/'run.txt').read_text()
    key=(re.search(r'Algorithm=(\w+);',report)[1],int(re.search(r'^H=(\d+);',report,re.M)[1]))
    if key not in sessions:sessions[key]=p
result=[]
for (method,height),path in sorted(sessions.items()):
    summaries=list(csv.DictReader(path.open()))
    best=None; reasons=collections.Counter(); tails=[]; valid=0; total=0
    for summary in summaries:
        rows=list(csv.DictReader((Path(summary['run_folder'])/'evaluations.csv').open()))
        local=None
        for i,row in enumerate(rows,1):
            assert int(row['evaluation'])==i
            total+=1; reasons[row['reason']]+=1
            if row['valid']=='True':
                valid+=1
                if local is None or float(row['cost'])<float(local['cost']):local=row
            assert abs(float(row['best_cost'])-(float(local['cost']) if local else 999999999))<1e-7
        assert len(rows)==int(summary['evaluations'])
        assert local and float(local['cost'])==float(summary['best_cost'])
        assert int(local['evaluation'])==int(summary['best_evaluation'])
        tails.append(len(rows)-int(local['evaluation']))
        if best is None or float(local['cost'])<float(best['candidate']['cost']):best={'summary':summary,'candidate':local}
    result.append({'method':method,'height':height,'session':str(path),'trials':len(summaries),'evaluations':total,
                   'valid_fraction':valid/total,'reasons':reasons.most_common(8),'mean_tail':sum(tails)/len(tails),'best':best})
(out/'latest-user-analysis.json').write_text(json.dumps(result,indent=2))
for r in result:
    print(r['method'],r['height'],'best',r['best']['candidate']['cost'],'seed',r['best']['summary']['seed'],
          'valid',round(r['valid_fraction'],3),'tail',round(r['mean_tail']),r['reasons'][:3])
