"""Detailed independent force, member and cost reconciliation for selected rows."""
from pathlib import Path
import csv, json, math
from independent_checks import reference, integrate
from independent_stem_profile import stem_profile

P=Path(__file__).resolve().parent
n = 2040000 / (15100 * math.sqrt(320))  # Reference Es/Ec, without intermediate rounding.
inputs=json.loads((P/'full-sizing-inputs.json').read_text())['inputs']
rows=list(csv.DictReader((P/'full-sizing-summary.csv').open()))
assert len(rows)==1 and int(rows[0]["eta"])==1
details=[]
native_text=(P/'full-sizing-vb6.md').read_text()
for row in rows:
    eta=int(row['eta'])
    values={key:float(row[key]) for key in ('tt','tb','TBase','Base','toe','heel')}
    hs=inputs['H']-values['TBase']; hp=inputs['H1']-values['TBase']
    r=reference(inputs['H'],values['tb'],values['TBase'],values['Base'],values['toe'],tt=values['tt'],h1=inputs['H1'],p=eta,db=int(row['stemDB']),sp=float(row['stemSP']))
    pressure=lambda x:r['qt']+(r['qh']-r['qt'])*x/values['Base']
    volume_stem=(values['tt']+values['tb'])/2*hs
    volume_base=values['Base']*values['TBase']
    members=[]
    section=native_text.split(f'## Passive fraction={eta}')[1].split('## Passive fraction=')[0]
    parsed={parts[1].strip():parts[2].strip() for line in section.splitlines() if line.startswith('| ') for parts in [line.split('|')]}
    for part,key,moment,length,thickness in [('Stem','stem',r['net'],hs+.4,values['tb']),('Toe','toe',r['mt'],values['toe']+.4,values['TBase']),('Heel','heel',r['mh'],values['heel']+.4,values['TBase'])]:
        db=int(row[key+'DB']); spacing=float(row[key+'SP'])
        area=math.pi*(db/10)**2/4/spacing
        depth=thickness-inputs['cover']-db/2000
        dc=depth*100
        neutral=(-n*area+math.sqrt((n*area)**2+200*n*area*dc))/100
        lever=dc-neutral/3
        mc=144*50*neutral*lever/1e5
        ms=1700*area*lever/1e5
        fc=abs(moment)*1e5/(50*neutral*lever)
        fs=abs(moment)*1e5/(area*lever)
        assert abs(moment)<=min(mc,ms)
        if part=='Stem':
            profile=stem_profile(inputs['H'],inputs['H1'],values['tb'],values['tt'],values['TBase'],inputs['cover'],db,spacing,eta)
            v=profile['max_v']
        elif part=='Toe':
            force=integrate(lambda x:pressure(x)-2.4*values['TBase']-1.8*hp,0,max(0,values['toe']-depth))
            v=abs(force)/(10*depth)
        else:
            force=integrate(lambda x:2.4*values['TBase']+1.8*hs-pressure(x),min(values['Base'],values['toe']+values['tb']+depth),values['Base'])
            v=abs(force)/(10*depth)
        for label,expected in [('signed M',moment),('concrete stress',fc),('steel stress',fs),('nominal shear V/bd',v)]:
            assert abs(float(parsed[part+' '+label].split()[0])-expected)<.000051, (eta,part,label)
        weight=.00617*db*db/spacing*length
        members.append(dict(member=part,DB=db,spacing=spacing,As=area,depth=depth,M=moment,fc=fc,fs=fs,
                            legacy_Mc=mc,legacy_Ms=ms,legacy_Mallow=min(mc,ms),nominal_v=v,
                            bar_length_allowance_included=length,steel_kg_per_m=weight,steel_cost_per_m=weight*24))
    steel_weight=sum(m['steel_kg_per_m'] for m in members)
    concrete_cost=(volume_stem+volume_base)*2617
    steel_cost=steel_weight*24
    estimate=concrete_cost+steel_cost
    assert abs(estimate-float(row['screening_estimate']))<1e-7
    details.append(dict(eta=eta,geometry=values,Pa=.5*(1/3)*1.8*inputs['H']**2,Pp=eta*.5*3*1.8*inputs['H1']**2,
                        active_arm=inputs['H']/3,passive_arm=inputs['H1']/3,MO=12.5-eta*1.5552,MR=r['mr'],W=r['w'],e=r['e'],
                        qtoe=r['qt'],qheel=r['qh'],FSot=r['fsot'],FSsl=r['fssl'],FSbc=inputs['qa_allowable']/max(r['qt'],r['qh']),
                        stem_profile=profile,members=members,volume_stem=volume_stem,volume_base=volume_base,
                        steel_kg=steel_weight,concrete_cost=concrete_cost,steel_cost=steel_cost,screening_estimate=estimate,
                        status='INDETERMINATE_WSD; not a complete construction price'))
(P/'recalculation-details.json').write_text(json.dumps(dict(inputs=inputs,cases=details),indent=2),encoding='ascii')
print('Independent reconciliation: 12 printed member values and the project quantity/cost total agree; legacy flexural capacities checked; WSD remains UNVERIFIED.')
for d in details:
    print('eta=',d['eta'],'cost=',round(d['screening_estimate'],5),'FS=',d['FSot'],d['FSsl'],d['FSbc'])
    for m in d['members']: print(m['member'],m['DB'],m['spacing'],'M=',round(m['M'],6),'legacy capacity=',round(m['legacy_Mallow'],6))
