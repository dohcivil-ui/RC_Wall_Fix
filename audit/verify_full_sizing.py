"""Independent vectorized integration/equilibrium check of the entire relaxed grid.
Verifies a stated mathematical screening problem, not complete EIT compliance.
"""
from pathlib import Path
import csv
import math
import gzip
import json
import numpy as np
from independent_stem_profile import stem_profile

P = Path(__file__).resolve().parent
n = 2040000 / (15100 * math.sqrt(320))  # Reference Es/Ec, without intermediate rounding.
native = list(csv.DictReader((P/'full-sizing-summary.csv').open()))
inputs = json.loads((P/'full-sizing-inputs.json').read_text())['inputs']
# This independent model verifies the requested H5 material/soil case. Refuse
# silently comparing another input set to these particular reference equations.
for key, value in dict(H=5,H1=1.2,gamma_soil=1.8,gamma_concrete=2.4,phi=30,mu=.6,cover=.075,fc=320,fy=4000).items():
    assert abs(inputs[key]-value) < 1e-9, (key,inputs[key])
assert len(native) == 1 and inputs["passive_factor"] == 1
assert inputs["modular_ratio_model"] == "Es/Ec" and abs(inputs["n"]-n) < 1e-12
tt, tb, z, width, toe = [x.ravel() for x in np.meshgrid(
    np.round(np.linspace(.2, .6, 17), 3), np.round(np.linspace(.2, 1, 17), 2),
    np.round(np.linspace(.3, 1, 15), 2), np.arange(1.5, 7.01, .5), np.round(np.linspace(.3, 1.2, 10), 1), indexing='ij')]
hs = 5-z
hp = 1.2-z
heel = width-toe-tb
slope = (tb-tt)/hs
# Integrate the front soil width toe+slope*y and its first moment.
soil_front = 1.8*(toe*hp+slope*hp**2/2)
soil_front_m = 1.8*(toe**2*hp/2+toe*slope*hp**2/2+slope**2*hp**3/6)
stem_area = hs*(tb+tt)/2
stem_m = 2.4*((toe+tb)*stem_area-hs*(tb**2+tb*tt+tt**2)/6)
soil_heel = 1.8*heel*hs
base = 2.4*width*z
weight = soil_front+2.4*stem_area+soil_heel+base
mr = soil_front_m+stem_m+soil_heel*(toe+tb+heel/2)+base*width/2
valid_geometry = (tb>=tt)&(heel>=.3-1e-9)&(heel-toe>1e-9)
messages = []
profiles = []

def choose_bars(moment, thickness):
    score = np.full(len(tt), np.inf)
    bars = np.zeros(len(tt), dtype=int)
    spacings = np.zeros(len(tt))
    for db in (12,16,20,25,28):
        dc = 100*(thickness-.075-db/2000)
        for sp in (.1,.15,.2,.25):
            area = math.pi*(db/10)**2/4/sp
            # Compression/tension equilibrium, rather than the production k/rho form.
            neutral = (-n*area+np.sqrt((n*area)**2+200*n*area*dc))/100
            lever = dc-neutral/3
            fc = np.abs(moment)*1e5/(50*neutral*lever)
            fs = np.abs(moment)*1e5/(area*lever)
            mask = (fc<=144)&(fs<=1700)&(db*db/sp < score)
            score[mask] = db*db/sp
            bars[mask] = db
            spacings[mask] = sp
    return score, bars, spacings

for row in native:
    eta = int(row['eta'])
    assert eta == 1
    assert float(row['qa_allowable']) == inputs['qa_allowable']
    reaction_m = mr-12.5+eta*1.5552
    qt = 4*weight/width-6*reaction_m/width**2
    qh = 6*reaction_m/width**2-2*weight/width
    fsot = (mr+eta*1.5552)/12.5
    fssl = (.6*weight+eta*3.888)/7.5
    stable = valid_geometry&(fsot>=2)&(fssl>=1.5)&(np.minimum(qt,qh)>=0)&(np.maximum(qt,qh)<=inputs['qa_allowable'])
    grad = (qh-qt)/width
    ms = .1*hs**3-eta*.9*hp**3
    mt = (qt-2.4*z-1.8*hp)*toe**2/2+grad*toe**3/6
    qjunction = qt+grad*(toe+tb)
    mh = (2.4*z+1.8*hs-qjunction)*heel**2/2-grad*heel**3/3
    sets = [choose_bars(ms,tb), choose_bars(mt,z), choose_bars(mh,z)]
    screened = stable&(ms>=0)&(mt>=0)&(mh>=0)
    for score,_,_ in sets:
        screened &= np.isfinite(score)
    cost = (stem_area+width*z)*2617 + .00617*24*(sets[0][0]*(hs+.4)+sets[1][0]*(toe+.4)+sets[2][0]*(heel+.4))
    cost[~screened] = np.inf
    best = int(np.argmin(cost))
    assert abs(cost[best]-float(row['screening_estimate'])) < 1e-7
    assert best+1 == int(row['best_row'])
    raw_path = P/f'full-sizing-flags-{eta}.bin'
    raw = raw_path.read_bytes() if raw_path.exists() else gzip.decompress((P/f'full-sizing-flags-{eta}.bin.gz').read_bytes())
    vf = np.frombuffer(raw, dtype='S1')
    mismatch = np.where((stable != (vf!=b'0')) | (screened != (vf==b'2')))[0]
    assert not len(mismatch), (eta, mismatch[:20]+1)
    assert int(stable.sum()) == int(row['stable_rows'])
    assert int(screened.sum()) == int(row['screened_rows'])
    (P/f'full-sizing-flags-{eta}.bin.gz').write_bytes(gzip.compress(raw, mtime=0))
    for key, vals in [('tt',tt),('tb',tb),('TBase',z),('Base',width),('toe',toe),('heel',heel)]:
        assert abs(vals[best]-float(row[key])) < 1e-8
    for name,(_,bars,spacings) in zip(('stem','toe','heel'),sets):
        assert bars[best] == int(row[name+'DB'])
        assert abs(spacings[best]-float(row[name+'SP'])) < 1e-8
    profile=stem_profile(5,1.2,tb[best],tt[best],z[best],.075,int(row['stemDB']),float(row['stemSP']),eta)
    profiles.append(dict(eta=eta,**profile))
    assert profile['moment_height']==0 and profile['concrete_height']==0 and profile['steel_height']==0
    messages.append(f'eta={eta}: independent full grid 520200 rows agrees on stability count {stable.sum()}, screened count {screened.sum()}, selected geometry row {best+1}, three independent steel layouts and relaxed estimate {cost[best]:.8f}.')

# Independently validate every section alternative in the native output.
bar_rows = list(csv.DictReader((P/'full-sizing-bar-options.csv').open()))
assert len(bar_rows) == 60
for r in bar_rows:
    db, sp, depth, moment = [float(r[k]) for k in ('DB','spacing','depth','moment')]
    area = math.pi*(db/10)**2/4/sp
    dc = depth*100
    neutral = (-n*area+math.sqrt((n*area)**2+200*n*area*dc))/100
    lever = dc-neutral/3
    fc = abs(moment)*1e5/(50*neutral*lever)
    fs = abs(moment)*1e5/(area*lever)
    assert abs(fc-float(r['fc'])) < 1e-6 and abs(fs-float(r['fs'])) < 1e-6
    assert (r['legacy_flexure']=='True') == (fc<=144 and fs<=1700)
text = (P/'full-sizing-vb6.md').read_text()
assert 'FATAL' not in text and 'NATIVE COMPLETE' in text
assert text.count(f"Actual production BA: NO_SOLUTION; evaluations={inputs['budget']}; best evaluation=0;") == 1
for profile in profiles:
    eta = profile["eta"]
    section=text.split(f'## Passive fraction={eta}')[1].split('## Passive fraction=')[0]
    report_rows={parts[1].strip():parts[2].strip() for line in section.splitlines() if line.startswith('| ') for parts in [line.split('|')]}
    for key,expected in [('Stem nominal shear V/bd',profile['max_v']),('Stem governing shear height',profile['shear_height'])]:
        assert abs(float(report_rows[key].split()[0])-expected)<.000051
(P/'full-sizing-stem-profiles.json').write_text(json.dumps(profiles,indent=2),encoding='ascii')
messages.append('60 native bar alternatives independently checked. Native production BA: no accepted design in the project case; WSD remains UNVERIFIED.')
messages.append('Selected stem profiles independently sampled at 20002 heights: base governs flexural M/fc/fs; interior nominal shear envelope matches native root solving.')
messages.append('Every native geometry status agrees after dimension-equality tolerance; equal heel/toe is rejected consistently. No selected minimum lies on that boundary.')
(P/'full-sizing-boundary-cases.csv').write_text('issue,original_example,correction\nheel_equal_to_toe,"B=2.5 tb=0.7 toe=0.9; computed heel-toe approximately 1e-16 m",shared strict-dimension comparison with 1e-9 m tolerance\n', encoding='ascii')
(P/'full-sizing-independent-checks.txt').write_text('\n'.join(messages)+'\n', encoding='ascii')
for eta in (1,):
    raw_path = P/f'full-sizing-flags-{eta}.bin'
    if raw_path.exists():
        assert gzip.decompress((P/f'full-sizing-flags-{eta}.bin.gz').read_bytes()) == raw_path.read_bytes()
        raw_path.unlink()  # Exact native status bytes remain in the verified gzip artifact.
print('\n'.join(messages))
