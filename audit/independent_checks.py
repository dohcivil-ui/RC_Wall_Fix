"""Independent force integration and cracked-section reference, NOT a VB6 run.
Generates numeric fixtures for the separately compiled native VB6 test project.
The synthetic shear/minimum settings used by tests are NOT EIT requirements.
"""
from pathlib import Path
import math, json
P=Path(__file__).resolve().parent
def polygon(points):
    cross=[points[i][0]*points[(i+1)%len(points)][1]-points[(i+1)%len(points)][0]*points[i][1] for i in range(len(points))]
    a=sum(cross)/2
    x=sum((points[i][0]+points[(i+1)%len(points)][0])*v for i,v in enumerate(cross))/(6*a)
    return abs(a),x
def integrate(f,a,b):
    # Simpson integration: exact for all pressure/lever-arm polynomials here.
    return (b-a)*(f(a)+4*f((a+b)/2)+f(b))/6
def reference(h,tb,tbase,b,toe,tt=.2,h1=1.2,p=0,db=28,sp=.1):
    hs=h-tbase; hp=h1-tbase; heel=b-toe-tb
    ka=math.tan(math.radians(45-30/2))**2; kp=1/ka
    active=integrate(lambda z:1.8*ka*(hs-z)*z,0,hs)
    passive=integrate(lambda z:p*1.8*kp*(hp-z)*z,0,hp)
    stem_area,stem_x=polygon([(toe,0),(toe+tb,0),(toe+tb,hs),(toe+tb-tt,hs)])
    soil_area,soil_x=polygon([(0,0),(toe,0),(toe+(tb-tt)*hp/hs,hp),(0,hp)])
    weights=[soil_area*1.8,heel*hs*1.8,stem_area*2.4,b*tbase*2.4]
    xs=[soil_x,toe+tb+heel/2,stem_x,b/2]
    w=sum(weights); mr=sum(a*x for a,x in zip(weights,xs))
    pa=integrate(lambda z:1.8*ka*(h-z),0,h)
    pp=integrate(lambda z:p*1.8*kp*(h1-z),0,h1)
    driving=integrate(lambda z:1.8*ka*(h-z)*z,0,h)
    resisting=integrate(lambda z:p*1.8*kp*(h1-z)*z,0,h1)
    reaction_m=mr-driving+resisting
    qt=4*w/b-6*reaction_m/b**2
    qh=6*reaction_m/b**2-2*w/b
    e=b/2-reaction_m/w
    q=lambda x:qt+(qh-qt)*x/b
    mt=integrate(lambda x:(q(x)-2.4*tbase-1.8*hp)*(toe-x),0,toe)
    mh=integrate(lambda x:(2.4*tbase+1.8*hs-q(x))*(x-toe-tb),toe+tb,b)
    depth=tb-.075-db/2000
    base_depth=tbase-.075-db/2000
    vs=abs(integrate(lambda z:1.8*ka*(hs-z),0,hs)-integrate(lambda z:p*1.8*kp*(hp-z),0,hp))
    vt=abs(integrate(lambda x:q(x)-2.4*tbase-1.8*hp,0,toe-base_depth)) if toe>base_depth else 0
    vh=abs(integrate(lambda x:2.4*tbase+1.8*hs-q(x),toe+tb+base_depth,b)) if heel>base_depth else 0
    area=math.pi*(db/10)**2/4/sp
    dc=100*depth
    neutral=(-9*area+math.sqrt((9*area)**2+200*9*area*dc))/100
    lever=dc-neutral/3
    fs=abs(active-passive)*1e5/(area*lever)
    fc=abs(active-passive)*1e5/(50*neutral*lever)
    return dict(h=h,tb=tb,tbase=tbase,b=b,toe=toe,tt=tt,h1=h1,p=p,hs=hs,active=active,net=active-passive,
                weights=weights,xs=xs,w=w,mr=mr,e=e,qt=qt,qh=qh,mt=mt,mh=mh,fs=fs,fc=fc,depth=depth,
                fsot=(mr+resisting)/driving,fssl=(pp+.6*w)/pa,vs=vs,vt=vt,vh=vh)
fixtures=[reference(h,{3:.35,4:.45,5:.6}[h],{3:.4,4:.5,5:.7}[h],{3:2,4:2.5,5:3.5}[h],{3:.3,4:.4,5:.5}[h]) for h in (3,4,5)]
fixtures.append(reference(3,.6,.4,4,.2,h1=2,p=1))
fixtures.append(reference(5,.6,.3,4,1))
weak=reference(5,.2,.3,3.5,.5,db=12,sp=.25)
assert math.isclose(weak['hs'],4.7)
assert abs(weak['active']-10.3823)<1e-9
net=reference(5,.2,.3,3.5,.5,p=1,db=12,sp=.25)
assert abs(net['net']-9.7262)<1e-9
assert weak['fs']>1700 and weak['fc']>144
assert fixtures[3]['e']<0 and fixtures[3]['qh']>fixtures[3]['qt']
assert fixtures[4]['vt']>0 and min(fixtures[4]['qt'],fixtures[4]['qh'])>=0
for f in fixtures[:3]:
    assert min(f['qt'],f['qh'])>=0 and f['fsot']>=2 and f['fssl']>=1.5
    assert f['fs']<=1700 and f['fc']<=144
(P/'independent-results.json').write_text(json.dumps(dict(fixtures=fixtures,weak=weak,full_passive=net),indent=2))
print('Independent numeric checks completed; not VB6 evidence.')
print('Weak H5: active=',weak['active'],'net(full passive)=',net['net'],'d=',weak['depth'],'fc=',weak['fc'],'fs=',weak['fs'])
def vbnum(x): return f'{x:.12f}'
vb=[]
for i,f in enumerate(fixtures):
    vb.append(f'''    H = {f['h']}: H1 = {f['h1']}: PassiveFactor = {f['p']}
    d = Fixture({f['tb']}, {f['tbase']}, {f['b']}, {f['toe']})
    Call Near("case{i} stem moment", CalculateMomentStem(d), {vbnum(f['net'])})
    Call Near("case{i} W1", CalculateW1(d, x), {vbnum(f['weights'][0])})
    Call Near("case{i} x1", x, {vbnum(f['xs'][0])})
    Call Near("case{i} W3", CalculateW3(d, x), {vbnum(f['weights'][2])})
    Call Near("case{i} x3", x, {vbnum(f['xs'][2])})
    Call AssertTrue("case{i} contact", BearingEdges(d, e, qt, qh))
    Call Near("case{i} signed e", e, {vbnum(f['e'])})
    Call Near("case{i} q toe", qt, {vbnum(f['qt'])})
    Call Near("case{i} q heel", qh, {vbnum(f['qh'])})
    Call Near("case{i} toe moment", CalculateMomentToe(d), {vbnum(f['mt'])})
    Call Near("case{i} heel moment", CalculateMomentHeel(d), {vbnum(f['mh'])})
    Call Near("case{i} stem shear", SectionShear(d, 0, SectionDepth(d.tb, 104)), {vbnum(f['vs'])})
    Call Near("case{i} toe shear", SectionShear(d, 1, SectionDepth(d.TBase, 104)), {vbnum(f['vt'])})
    Call Near("case{i} heel shear", SectionShear(d, 2, SectionDepth(d.TBase, 104)), {vbnum(f['vh'])})
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("case{i} concrete stress", ca, {vbnum(f['fc'])})
    Call Near("case{i} steel stress", sa, {vbnum(f['fs'])})
''')
    if i<3:
        vb.append(f'''    Call AssertTrue("H{f['h']} synthetic fixture valid", Valid(d))
    d.tb = 0.2: d.LHeel = d.Base - d.LToe - d.tb: d.ASst_DB = 100: d.ASst_Sp = 113
    Call AssertTrue("H{f['h']} weak section rejected", Not Valid(d))
''')
    elif i==3:
        vb.append('    Call AssertTrue("reversed tension faces rejected", Not Valid(d))\n')
(P/'fixture_checks.inc').write_text('\n'.join(vb),encoding='ascii')
