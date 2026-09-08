"""Assess decision sensitivity to n using the same active/full-passive load case.
Archived native grid flags are compared, not substituted for a fresh VB6 run.
"""
from pathlib import Path
import math, json, gzip, hashlib
import numpy as np

P = Path(__file__).resolve().parent
paths = {
    'n9': P/'full-sizing-checks/20260908-212528-480/full-sizing-flags-1.bin.gz',
    'EsEc': P/'full-sizing-checks/20260908-210452-544/full-sizing-flags-1.bin.gz',
}
flags = {key: np.frombuffer(gzip.decompress(path.read_bytes()), dtype='S1') for key,path in paths.items()}
assert len(flags['n9']) == len(flags['EsEc']) == 520200
assert np.array_equal(flags['n9'] != b'0', flags['EsEc'] != b'0')
loss = np.flatnonzero((flags['n9'] == b'2') & (flags['EsEc'] != b'2'))
gain = np.flatnonzero((flags['n9'] != b'2') & (flags['EsEc'] == b'2'))
n_exact = 2040000/(15100*math.sqrt(320))

def section(moment, thickness, db, spacing, n):
    depth = 100*(thickness-.075-db/2000)
    area = math.pi*(db/10)**2/(4*spacing)
    x = (-n*area+math.sqrt((n*area)**2+200*n*area*depth))/100
    lever = depth-x/3
    return dict(n=n,As=area,d_cm=depth,k=x/depth,j=lever/depth,
                fc=moment*1e5/(50*x*lever),fs=moment*1e5/(area*lever))

moment = .1*4.4**3-.9*.6**3
example = [section(moment,.25,28,.1,n) for n in (9,n_exact)]
assert example[0]['fc'] < 144 < example[1]['fc']
assert max(row['fs'] for row in example) < 1700
assert int(loss[0])+1 == 2541
selected = [section(9.7262,.4,16,.1,n) for n in (9,n_exact)]
result = dict(
    decision='Use Es/Ec: fixed n changes flexural acceptance; this is engineering sensitivity, not statistical significance.',
    load_case='H=5, H1=1.2 from base underside; active and full passive; qa=30 allowable; fc_prime=320; fy=4000',
    provenance={key:dict(path=str(path.relative_to(P)),sha256=hashlib.sha256(path.read_bytes()).hexdigest()) for key,path in paths.items()},
    total_rows=520200,stable_rows=int((flags['n9']!=b'0').sum()),
    n9_screened=int((flags['n9']==b'2').sum()),exact_screened=int((flags['EsEc']==b'2').sum()),
    n9_only_count=len(loss),exact_only_count=len(gain),n9_only_first_rows=(loss[:20]+1).tolist(),
    example=dict(geometry_row=2541,H=5,H1=1.2,tt=.2,tb=.25,TBase=.6,Base=2.5,toe=.3,heel=1.95,DB=28,spacing=.1,M=moment,sections=example),
    selected_stem=selected,
    selected_fc_increase_percent=100*(selected[1]['fc']/selected[0]['fc']-1),
    n9_over_exact_percent=100*(9/n_exact-1))
(P/'modular-ratio-sensitivity.json').write_text(json.dumps(result,indent=2),encoding='ascii')
vb = ['    Dim savedSensitivityWSD As WSDParams, sensitivityFile As Integer',
      '    savedSensitivityWSD = currentWSD',
      '    Dim savedSensitivityQa As Double',
      '    savedSensitivityQa = qa: qa = 30',
      '    H = 5: H1 = 1.2: PassiveFactor = 1',
      '    d = Fixture(0.25, 0.6, 2.5, 0.3)',
      '    Call AssertTrue("n sensitivity example overturning", CheckFS_OT(d, ca))',
      '    Call AssertTrue("n sensitivity example sliding", CheckFS_SL(d, ca))',
      '    Call AssertTrue("n sensitivity example bearing", CheckFS_BC(d, ca, e, qt, qh))',
      f'    Call Near("n sensitivity stem moment", CalculateMomentStem(d), {moment:.12f})',
      '    sensitivityFile = FreeFile',
      '    Open App.Path & "\\modular-ratio-sensitivity-native.csv" For Output As #sensitivityFile',
      '    Print #sensitivityFile, "model,n,fc,fs,flexural_acceptance"']
for i,row in enumerate(example):
    vb.append('    currentWSD = CalculateWSDParameters(4000, 320)')
    if i == 0:
        vb.append('    currentWSD.n = 9#  \' Isolated comparison only; restore production parameters below.')
    vb.append('    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)')
    vb.append(f'    Call Near("n sensitivity case{i} concrete", ca, {row["fc"]:.12f})')
    vb.append(f'    Call Near("n sensitivity case{i} steel", sa, {row["fs"]:.12f})')
    vb.append('    ok = CheckSteelOK(CalculateMomentStem(d), SectionDepth(d.tb, 104), 104, 110)')
    vb.append(f'    Call AssertTrue("n sensitivity case{i} flexural decision", {"ok" if i==0 else "Not ok"})')
    vb.append(f'    Print #sensitivityFile, "{"fixed9" if i==0 else "EsEc"}," & CsvNumber(currentWSD.n) & "," & CsvNumber(ca) & "," & CsvNumber(sa) & "," & CStr(ok)')
vb.extend(['    Close #sensitivityFile','    currentWSD = savedSensitivityWSD', '    qa = savedSensitivityQa'])
(P/'modular_ratio_decision_checks.inc').write_text('\n'.join(vb)+'\n',encoding='ascii')
print(f'Same-load native grids: {len(loss)} n9-only flexural acceptances; {len(gain)} exact-only. Example concrete stress {example[0]["fc"]:.4f} -> {example[1]["fc"]:.4f}, limit 144.')
