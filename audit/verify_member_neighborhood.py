"""Validate native member-neighborhood results, budgets and shared policy."""
from pathlib import Path
import csv,json,re
P=Path(__file__).resolve().parent/'h5-search'
final=P/'final'
assert (final/'parity.txt').read_text().strip()=='GENERATOR checks=5400; failures=0'
def bench(folder):
    text=(folder/'bench.txt').read_text()
    assert 'FATAL=' not in text
    records=[]
    for match in re.finditer(r'H=(\d); method=(HCA|BA); seed=(\d+); cost=([\d.]+); evaluation=(\d+)\nRUN=([^\n]+)',text):
        height,method,seed,cost,evaluation,path=match.groups()
        records.append(dict(height=int(height),method=method,seed=int(seed),cost=float(cost),evaluation=int(evaluation),path=path))
    assert len(records)==12
    return records
records=bench(final)
prototype=bench(P/'member')
assert [{k:v for k,v in r.items() if k!='path'} for r in records]==[{k:v for k,v in r.items() if k!='path'} for r in prototype]
traces={}
for record in records:
    path=Path(record['path'])
    with (path/'evaluations.csv').open(newline='') as f:rows=list(csv.DictReader(f))
    report=(path/'run.txt').read_text()
    assert 'SearchPolicy=MEMBER_RANDOM_V1' in report
    assert len(rows)==5000 and 'Evaluations=5000; Budget=5000' in report
    lowest=999999999.0;first=0
    for n,row in enumerate(rows,1):
        assert int(row['evaluation'])==n
        if row['valid']=='True' and float(row['cost'])<lowest:lowest=float(row['cost']);first=n
        assert abs(float(row['best_cost'])-lowest)<1e-7
    assert abs(lowest-record['cost'])<1e-7 and first==record['evaluation']
    if record['method']=='HCA':assert all(r['entry']=='neighbor' for r in rows[1:])
    with (path/f"accept-{record['method']}-H{record['height']}.csv").open(newline='') as f:
        reader=csv.DictReader(f);accept=list(reader)
        assert reader.fieldnames==['No.','Rejected','Passed','Passed and Better value']
        assert len(accept)==5000 and [int(r['No.']) for r in accept]==list(range(5000))
    traces[record['height'],record['method'],record['seed']]=rows
for height,seed in {(r['height'],r['seed']) for r in records}:
    assert traces[height,'HCA',seed][:21]==traces[height,'BA',seed][:21]
target=next(r for r in records if (r['height'],r['method'],r['seed'])==(5,'HCA',12374))
assert target['cost']==8538.41892 and target['evaluation']==4866
assert traces[5,'HCA',12374][0]['entry']=='initial'
assert float(traces[5,'HCA',12374][0]['cost'])>target['cost']
for height,cost in [(3,2942.34192),(4,4985.56464)]:
    assert all(r['cost']==cost for r in records if r['height']==height)
summary={'policy':'MEMBER_RANDOM_V1','generator_comparisons':5400,'records':records,
         'normal_start_target':target,'BA_always_faster':False,'global_optimum_certified':False}
(final/'verification.json').write_text(json.dumps(summary,indent=2))
print('Native member policy verified: 5400 matching generator comparisons; 12 exact-budget runs; HCA H5 seed12374 reaches 8538.41892 at evaluation4866; no universal BA-speed advantage claimed.')
