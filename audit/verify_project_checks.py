"""Independent equilibrium, force integration, local stresses and quantity audit."""
from pathlib import Path
import csv, json, math, sys, itertools
import numpy as np
from independent_checks import reference, integrate
from independent_stem_profile import stem_profile
P=Path(__file__).resolve().parent
E=Path(sys.argv[1])
psi=.45359237/(2.54**2)
comparisons=0
def near(actual, expected, tol=1e-7):
    global comparisons
    comparisons+=1
    assert abs(float(actual)-expected)<tol,(actual,expected)
def peak(f, length):
    xx=np.linspace(0,length,201)
    ii=max(range(len(xx)),key=lambda i:f(float(xx[i])))
    lo=float(xx[max(0,ii-1)]);hi=float(xx[min(len(xx)-1,ii+1)])
    for _ in range(65):
        left=lo+(hi-lo)/3;right=hi-(hi-lo)/3
        if f(left)<f(right):lo=left
        else:hi=right
    return max(f(0),f(length),f((lo+hi)/2))
def ld(db,sp,top):
    # Inches/psi calculation throughout, convert the resulting length once.
    diameter=db/25.4
    c=min(75/25.4+diameter/2,sp/0.0254/2)
    value=3/40*(4000/psi)/math.sqrt(320/psi)*(1.3 if top else 1)*(.8 if db<=19 else 1)/min(c/diameter,2.5)*diameter
    return max(12,value)*.0254
def mass(db,sp,length):return .00617*db**2/sp*length
summaries=[]
fixtures=list(csv.DictReader((E/'project-fixtures.csv').open()))
for row in fixtures:
    if row['case'] not in ('H3','H4','H5'):continue
    assert row['valid']=='True'
    h,tt,tb,tbase,B,toe=[float(row[key]) for key in ('H','tt','tb','TBase','B','toe')]
    heel=B-toe-tb;hs=h-tbase;hp=1.2-tbase
    raw=list(csv.reader((E/f"project-{row['case']}-detail.csv").open()))
    members=[dict(zip(raw[0],line)) for line in raw[1:4]]
    extra={line[0]:[float(x) for x in line[1:]] for line in raw[4:]}
    independent=reference(h,tb,tbase,B,toe,tt=tt,p=1,db=int(members[0]['DB']),sp=float(members[0]['SP']))
    ot=[];sl=[];bc=[];mom=[0,0];shear=[0,0]
    for fac in itertools.product((.85,1),repeat=4):
        weight=sum(a*b for a,b in zip(independent['weights'],fac))
        mr=sum(a*b*x for a,b,x in zip(independent['weights'],fac,independent['xs']))
        pa=.3*h*h;pp=2.7*1.2**2;net=pa*h/3-pp*1.2/3
        eccentric=B/2-(mr-net)/weight
        qt=weight/B*(1+6*eccentric/B);qh=weight/B*(1-6*eccentric/B)
        pressure=lambda x:qt+(qh-qt)*x/B
        assert min(qt,qh)>=0
        near(integrate(pressure,0,B),weight)
        near(integrate(lambda x:pressure(x)*x,0,B),mr-net)
        ot.append((mr+pp*1.2/3)/(pa*h/3));sl.append((.6*weight+pp)/pa);bc.append(30/max(qt,qh))
        toe_net=lambda x:pressure(x)-2.4*tbase*fac[3]-1.8*hp*fac[0]
        heel_net=lambda x:2.4*tbase*fac[3]+1.8*hs*fac[1]-pressure(B-x)
        for i,(load,length) in enumerate(((toe_net,toe),(heel_net,heel))):
            # Independent Simpson integration with continuously varied cut location.
            moment=lambda z:integrate(lambda x:load(x)*(z-x),0,z)
            force=lambda z:abs(integrate(load,0,z))
            mom[i]=max(mom[i],peak(moment,length))
            depth=float(members[i+1]['depth'])
            shear[i]=max(shear[i],peak(force,length)/(10*depth))
    totals=extra['totals']
    for a,b in zip(totals[:3],(min(ot),min(sl),min(bc))):near(a,b)
    assert min(ot)>=2 and min(sl)>=1.5 and min(bc)>=1
    main_weight=0
    for i,member in enumerate(members):
        db=int(member['DB']);sp=float(member['SP']);depth=(tb if i==0 else tbase)-.075-db/2000
        area=math.pi*(db/10)**2/4/sp
        near(member['depth'],depth);near(member['As'],area)
        minimum=max(3*math.sqrt(320/psi),200)/(4000/psi)*10000*depth if i==0 else .002*10000*tbase
        near(member['minimum'],minimum);assert area>=minimum
        development=ld(db,sp,i==2 and depth>12*.0254)
        near(member['ld'],development)
        length=(hs if i==0 else toe if i==1 else heel)-.075+development
        near(member['length'],length);main_weight+=mass(db,sp,length)
        if i==0:
            profile=stem_profile(h,1.2,tb,tt,tbase,.075,db,sp,1)
            near(member['M'],profile['max_M'],1e-6);near(member['v'],profile['max_v'])
            assert float(member['fc_bound'])+1e-7>=profile['max_fc']
            assert float(member['fs_bound'])+1e-7>=profile['max_fs']
            assert profile['max_v']<=.5*1.1*math.sqrt(320/psi)*psi
        else:
            near(member['M'],mom[i-1]);near(member['v'],shear[i-1])
        assert float(member['fc_bound'])<=144 and float(member['fs_bound'])<=1700
    front_db,front_sp,front_ld=extra['front'];hor_db,hor_sp,hor_lap=extra['horizontal'];base_db,base_sp,base_lap=extra['base']
    near(front_ld,ld(front_db,front_sp,False));near(hor_lap,1.3*ld(hor_db,hor_sp,True));near(base_lap,1.3*ld(base_db,base_sp,True))
    extra_weight=mass(front_db,front_sp,hs-.075+front_ld)+2*mass(hor_db,hor_sp,hs-.15)*12/(12-hor_lap)+2*mass(base_db,base_sp,B-.15)*12/(12-base_lap)
    concrete=(tt+tb)/2*hs+B*tbase
    near(totals[3],concrete);near(totals[4],main_weight);near(totals[5],extra_weight)
    cost=concrete*2617+(main_weight+extra_weight)*24;near(totals[6],cost,1e-6)
    summaries.append(dict(case=row['case'],geometry=dict(H=h,tt=tt,tb=tb,TBase=tbase,B=B,toe=toe,heel=heel),
                          fs=[min(ot),min(sl),min(bc)],cost=cost,main_weight=main_weight,extra_weight=extra_weight,
                          note='Fixed verification fixture; not an optimized or fully code-certified design'))
assert len(summaries)==3
reason={row['case']:row['reason'] for row in fixtures}
assert reason['H3_reverse_heel']=='REVERSED_BASE_FACE'
assert reason['DB12weak']=='STEM_STRESS_OR_INTERVAL_BOUND'
assert reason['short_anchorage']=='STEM_STRAIGHT_ANCHORAGE'
report=dict(comparisons=comparisons,fixtures=summaries,scope='Independent arithmetic verified against actual VB6 outputs; no 30-trial run')
(E/'independent-project.json').write_text(json.dumps(report,indent=2),encoding='ascii')
(P/'project-checks-latest.json').write_text(json.dumps(dict(evidence=str(E.relative_to(P)),**report),indent=2),encoding='ascii')
print(f'Independent project checks: {comparisons} numerical comparisons; H3/H4/H5 verified; no 30-trial run.')
