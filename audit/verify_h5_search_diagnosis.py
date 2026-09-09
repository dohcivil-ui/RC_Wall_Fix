"""Check the persisted native findings without claiming heuristic global optimality."""
from pathlib import Path
import csv,json,re
P=Path(__file__).resolve().parent/'h5-search'
report=json.loads((P/'latest-user-analysis.json').read_text())
assert sum(r['evaluations'] for r in report)==900000
best={(r['method'],r['height']):float(r['best']['candidate']['cost']) for r in report}
assert best['BA',3]==best['HCA',3] and best['BA',4]==best['HCA',4]
assert best['BA',5]==8538.41892 and best['HCA',5]==8646.93688
native=(P/'before/probe.txt').read_text()
assert 'BA_WINNER_IN_HCA valid=True; cost=8538.41892' in native
assert 'REPLAY_12360 cost=8646.93688' in native
assert native.count('CHEAPER_PAIR')==3
heel={}
for line in (P/'transitions/probe.txt').read_text().splitlines():
    m=re.fullmatch(r'PAIR part=2; db=(\d+); sp=(\d+); valid=(True|False); cost=([\d.]+)',line)
    if m and m[3]=='True':heel[int(m[1]),int(m[2])]=float(m[4])
start=(101,110); target=(103,113)
assert heel[target]<heel[start]
assert abs(target[1]-start[1])>2
reachable={start}
while True:
    newer={b for a in reachable for b in heel if abs(a[0]-b[0])<=2 and abs(a[1]-b[1])<=2 and heel[b]<heel[a]}
    if newer<=reachable:break
    reachable|=newer
assert reachable=={start}
count=0
for folder in ('before','grouped','spacing3'):
    text=(P/folder/'probe.txt').read_text()
    assert 'FATAL=' not in text
    for name in re.findall(r'; RUN=([^\n]+)',text):
        path=Path(name.strip())
        rows=list(csv.DictReader((path/'evaluations.csv').open()))
        run=(path/'run.txt').read_text()
        assert len(rows)==int(re.search(r'Budget=(\d+)',run)[1])
        lowest=999999999.0
        for n,row in enumerate(rows,1):
            assert int(row['evaluation'])==n
            if row['valid']=='True':lowest=min(lowest,float(row['cost']))
            assert abs(lowest-float(row['best_cost']))<1e-7
        count+=1
print(f'Native diagnosis verified: {count} search traces; BA winner accepted by HCA; heel has no strictly cheaper +/-2 member-only path. No production fix claimed.')
