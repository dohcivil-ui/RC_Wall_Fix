"""Validate the released shared random policy and every reported native benchmark."""
from pathlib import Path
import csv,json,re
P=Path(__file__).resolve().parent/'h5-search'
release=P/'release'
assert (release/'parity.txt').read_text().strip()=='GENERATOR checks=5400; failures=0'
def read_records(path):
    text=path.read_text();assert 'FATAL=' not in text
    return [dict(height=int(h),method=m,seed=int(s),cost=float(c),evaluation=int(e),path=p.strip())
        for h,m,s,c,e,p in re.findall(r'H=(\d); method=(HCA|BA); seed=(\d+); cost=([\d.]+); evaluation=(\d+)\nRUN=([^\n]+)',text)]
benchmark=read_records(release/'bench.txt')
prototype=read_records(P/'weighted'/'bench.txt')
without_path=lambda rows:[{k:v for k,v in r.items() if k!='path'} for r in rows]
assert without_path(benchmark)==without_path(prototype)
holdout=read_records(release/'holdout.txt')
assert len(benchmark)==12 and len(holdout)==6
records=benchmark+holdout
traces={}
for r in records:
    p=Path(r['path'])
    with (p/'evaluations.csv').open(newline='') as f:rows=list(csv.DictReader(f))
    report=(p/'run.txt').read_text()
    assert 'SearchPolicy=MEMBER_WEIGHTED_V1' in report
    assert len(rows)==5000 and 'Evaluations=5000; Budget=5000' in report
    best=999999999.0;first=0
    for n,row in enumerate(rows,1):
        assert int(row['evaluation'])==n
        if row['valid']=='True' and float(row['cost'])<best:best=float(row['cost']);first=n
        assert abs(float(row['best_cost'])-best)<1e-7
    assert abs(best-r['cost'])<1e-7 and first==r['evaluation']
    if r['method']=='HCA':assert all(row['entry']=='neighbor' for row in rows[1:])
    with (p/f"accept-{r['method']}-H{r['height']}.csv").open(newline='') as f:
        reader=csv.DictReader(f);accept=list(reader)
        assert reader.fieldnames==['No.','Rejected','Passed','Passed and Better value']
    assert len(accept)==5000 and [int(row['No.']) for row in accept]==list(range(5000))
    traces[r['height'],r['method'],r['seed']]=rows
for height,seed in {(r['height'],r['seed']) for r in records}:
    assert traces[height,'HCA',seed][:21]==traces[height,'BA',seed][:21]
targets=[r for r in records if r['height']==5 and abs(r['cost']-8538.41892)<1e-7]
assert {r['method'] for r in targets}=={'HCA','BA'}
for r in targets:
    rows=traces[r['height'],r['method'],r['seed']]
    assert float(rows[0]['cost'])>r['cost'] and rows[0]['entry']=='initial'
assert any(r['method']=='HCA' and r['seed']==12374 and r['evaluation']==3162 for r in targets)
assert any(r['method']=='BA' and r['seed']==12345 and r['evaluation']==4196 for r in targets)
summary={'policy':'MEMBER_WEIGHTED_V1','generator_comparisons':5400,'records':records,'target_runs':targets,
         'BA_always_faster':False,'global_optimum_certified':False,'development_set_used_for_policy_selection':True}
(release/'verification.json').write_text(json.dumps(summary,indent=2))
print('Native release verified: 5400 generator comparisons, 18 exact-budget runs, both methods reach 8538.41892 from normal starts. No universal relative-speed or global-optimum claim.')
