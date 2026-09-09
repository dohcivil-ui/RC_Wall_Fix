from pathlib import Path
import re,hashlib,json,csv
P=Path(__file__).resolve().parent.parent
B=P/'audit'/'baseline'
def read(p): return p.read_bytes().decode('latin1').replace('\r\n','\n')
def routine(s,name):
    m=re.search(rf'(?:Public |Private )?(?:Function|Sub) {name}\([\s\S]*?\nEnd (?:Function|Sub)',s)
    assert m,name
    return m.group()
result=[]
assert routine(read(P/'modBA.bas'),'GenerateNeighbor_BA') == routine(read(B/'modBA.bas'),'GenerateNeighbor_BA')
result.append('BA original neighbor routine unchanged')
ba_outside_neighbor = read(P/'modBA.bas').replace(routine(read(P/'modBA.bas'),'GenerateNeighbor_BA'), '')
ba_outside_neighbor = '\n'.join(line.split("'",1)[0] for line in ba_outside_neighbor.splitlines())
assert not re.search(r'\b(?:Rand\s*\(|Rnd\b|Randomize\b)', ba_outside_neighbor, re.I)
assert 'Midtb = Currenttb: MidTBase = CurrentTBase: MidBase = CurrentBase' in ba_outside_neighbor
result.append('BA recovery reopens bounds at current; no random calls outside the HCA-equivalent neighbor generator')
def normalized_neighbor(body):
    return [re.sub(r'\s+', '', line.split("'",1)[0]).lower() for line in body.splitlines()
            if line.split("'",1)[0].strip()]
ba_neighbor = routine(read(P/'modBA.bas'),'GenerateNeighbor_BA').replace('GenerateNeighbor_BA','GenerateNeighbor')
for old,new in [('Mintb','TB_MIN'),('Maxtb','FixedTbMax'),('MinTBase','TBASE_MIN'),('MaxTBase','FixedTBaseMax'),('MinBase','FixedBaseMin'),('MaxBase','FixedBaseMax')]:
    ba_neighbor = re.sub(r'\b'+old+r'\b',new,ba_neighbor)
assert normalized_neighbor(ba_neighbor) == normalized_neighbor(routine(read(P/'modHillClimbing.bas'),'GenerateNeighbor'))
assert 'MainSteelNeighbor' not in read(P/'modHillClimbing.bas') and 'refinementVisit' not in read(P/'modHillClimbing.bas')
result.append('HCA matches BA draw order, step sizes and repairs with full bounds; final steel sweep removed')
bounds=set(re.findall(r'Private (Min\w+|Max\w+) As Integer',read(P/'modBA.bas')))
assert bounds=={'Mintb','Maxtb','MinTBase','MaxTBase','MinBase','MaxBase'}
result.append('BA still has only tb, TBase, Base bisection bounds')
for fn in ['CalculateCostFull','CalculateSteelWeight','GetConcretePrice']:
    assert routine(read(P/'modShared.bas'),fn)==routine(read(B/'modShared.bas'),fn)
result.append('Original unit-price and bar-mass primitives unchanged; project detail quantities use modProjectChecks')
vbp=read(P/'RC_RT_HCA_v2.vbp')
modules=re.findall(r'^Module=[^;]+; (.+)$',vbp,re.M)+re.findall(r'^Form=(.+)$',vbp,re.M)
for module in modules:
    assert ':' not in module and (P/module).is_file(),module
result.append('All .vbp source modules resolve within this copy')
external=Path(r'C:\Users\moosu\Downloads\modShared.bas')
assert external.read_bytes()==(B/'modShared.ACTIVE-external.bas').read_bytes()
result.append('External Downloads/modShared.bas remains byte-identical to baseline')
changed=[]
for p in P.iterdir():
    if p.suffix.lower() in ('.bas','.frm','.vbp') and (not (B/p.name).exists() or p.read_bytes()!=(B/p.name).read_bytes()):
        data=p.read_bytes()
        assert not data.startswith(b'\xef\xbb\xbf')
        assert data.count(b'\r\n')==data.count(b'\n'),p.name
        changed.append(dict(file=p.name,sha256=hashlib.sha256(data).hexdigest()))
result.append('Changed VB6 files retain byte-preserving legacy text and CRLF, without UTF-8 BOM')
assert read(P/'modProjectChecks.bas') == (P/'audit'/'modProjectChecks.source').read_text(encoding='ascii')
result.append('Production project-check module matches its archived source')
assert read(P/'frmBestDesign.frm') == (P/'audit'/'best_design_form.source').read_text(encoding='ascii')
result.append('Best-design popup matches archived source; drawing is separate from calculation modules')
for handler in ('cmdBA_Click','cmdRun_Click'):
    body=routine(read(P/'Form1.frm'),handler)
    assert body.count('frmBestDesign.ShowBest')==1
    assert body.index('frmBestDesign.ShowBest')>body.index('Next trial')
    assert 'globalBestCost, cover, CLng(numTrials), globalBestIteration)' in body
    assert 'BestCostIteration < globalBestIteration' in body
    assert body.count('txtSeed.Text') == 1
    assert 'RandomSeed:=firstSeed + CLng(trial) - 1' in body
    assert 'Call ClearGraph(picGraph)' in body
result.append('BA/HCA popup is called once after the trial loop, with session winner and first-best evaluation')
# Independently inspect every persisted optimizer trace against its own valid candidates.
traces=0
trace_paths = list((P/'audit'/'results').glob('*/evaluations.csv')) + list((P/'result_csv').glob('*/evaluations.csv'))
for path in trace_paths:
    rows=list(csv.DictReader(path.open()))
    best=999999999.0
    for i,row in enumerate(rows,1):
        assert int(row['evaluation'])==i,path
        if row['valid']=='True': best=min(best,float(row['cost']))
        assert abs(best-float(row['best_cost']))<1e-7,path
    run=(path.parent/'run.txt').read_text()
    m=re.search(r'Evaluations=(\d+); Budget=(\d+)',run)
    assert m and len(rows)==int(m[1])==int(m[2]),path
    traces+=1
result.append(f'{traces} persisted traces: sequential evaluations, exact budgets, prefix-best from all valid entries')
(P/'audit'/'source-checks.txt').write_text('\n'.join(result),encoding='utf-8')
(P/'audit'/'changed-files.json').write_text(json.dumps(changed,indent=2))
print('\n'.join(result))
