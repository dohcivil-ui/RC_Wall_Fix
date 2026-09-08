"""Generate independent equilibrium/force-path fixtures for actual VB6 regression."""
from pathlib import Path
import json
from independent_checks import reference, integrate
from independent_stem_profile import stem_profile

P = Path(__file__).resolve().parent
cases = []
vb = ['''    ' Force path: same geometry, change passive only; no synthetic criteria used here.
    Dim chainW As Double, cw1 As Double, cw2 As Double, cw3 As Double, cw4 As Double
    Dim cx1 As Double, cx2 As Double, cx3 As Double, cx4 As Double
    Dim cot As Double, csl As Double, cbc As Double, cqmax As Double, cqmin As Double
    H = 5: H1 = 1.2: qa = 30: mu = 0.6
    currentMaterial.concretePrice = GetConcretePrice(320)
    f = FreeFile
    Open App.Path & "\\force-chain-native.csv" For Output As #f
    Print #f, "eta,Pa,Pp,active_arm,passive_arm,MO,MR,W,xR,e,qtoe,qheel,Mstem,Mtoe,Mheel,FSot,FSsl,FSbc,cost"
''']
for eta in (0, 1):
    r = reference(5,.45,.3,3,1.1,p=eta)
    pa = integrate(lambda y: 1.8/3*(5-y), 0,5)
    pp = integrate(lambda y: eta*1.8*3*(1.2-y), 0,1.2)
    mo = integrate(lambda y:1.8/3*(5-y)*y,0,5)-integrate(lambda y:eta*1.8*3*(1.2-y)*y,0,1.2)
    volume = (.2+.45)/2*4.7+3*.3
    steel = sum(.00617*db*db/sp*length for db,sp,length in [(25,.25,5.1),(20,.25,1.5),(20,.2,1.85)])
    cost = volume*2617+steel*24
    row = dict(eta=eta,Pa=pa,Pp=pp,active_arm=5/3,passive_arm=.4,MO=mo,MR=r['mr'],W=r['w'],xR=(r['mr']-mo)/r['w'],e=r['e'],qtoe=r['qt'],qheel=r['qh'],Mstem=r['net'],Mtoe=r['mt'],Mheel=r['mh'],FSot=r['fsot'],FSsl=r['fssl'],FSbc=30/max(r['qt'],r['qh']),cost=cost)
    cases.append(row)
    vb.append(f'''    PassiveFactor = {eta}
    d = Fixture(0.45, 0.3, 3, 1.1)
    d.ASst_DB = 103: d.ASst_Sp = 113
    d.AStoe_DB = 102: d.AStoe_Sp = 113
    d.ASheel_DB = 102: d.ASheel_Sp = 112
    chainW = CalculateWTotal(d, cw1, cw2, cw3, cw4, cx1, cx2, cx3, cx4)
    Call AssertTrue("force-chain eta={eta} full contact", BearingEdges(d, e, qt, qh))
    Call CheckFS_OT(d, cot): Call CheckFS_SL(d, csl)
    Call CheckFS_BC(d, cbc, x, cqmax, cqmin)
''')
    expressions = dict(Pa='CalculatePa()',Pp='CalculatePp()',active_arm='H / 3#',passive_arm='H1 / 3#',MO='CalculateMO(d)',MR='CalculateMR(d)',W='chainW',xR='d.Base / 2# - e',e='e',qtoe='qt',qheel='qh',Mstem='CalculateMomentStem(d)',Mtoe='CalculateMomentToe(d)',Mheel='CalculateMomentHeel(d)',FSot='cot',FSsl='csl',FSbc='cbc',cost='CalculateCost(d)')
    for key, expr in expressions.items():
        vb.append(f'    Call Near("force-chain eta={eta} {key}", {expr}, {row[key]:.12f})\n')
    vb.append(f'''    Call Near("force-chain eta={eta} vertical equilibrium", (qt + qh) * d.Base / 2#, {r['w']:.12f})
    Call Near("force-chain eta={eta} moment equilibrium", d.Base ^ 2 * (qt + 2# * qh) / 6#, {r['mr']-mo:.12f})
''')
    # Keep each physical quantity in the CSV visible, rather than only a test label.
    vb.append('    textA = CStr(PassiveFactor)\n')
    for expr in expressions.values():
        vb.append(f'    textA = textA & "," & CsvNumber({expr})\n')
    vb.append('    Print #f, textA\n')
vb.append('    Close #f\n    PassiveFactor = 0: qa = 20: currentMaterial = mat\n')
profiles=[]
vb.append('    f = FreeFile\n    Open App.Path & "\\stem-shear-native.csv" For Output As #f\n    Print #f, "case,face_v,max_v,height"\n')
for i,(tb,tt,tz,db,sp,eta) in enumerate([(.4,.2,.3,16,.1,1),(.2,.2,.3,12,.25,1),(.45,.2,.3,25,.25,0),(.4,.2,.45,16,.1,1)]):
    p=stem_profile(5,1.2,tb,tt,tz,.075,db,sp,eta)
    profiles.append(dict(case=i,tb=tb,tt=tt,TBase=tz,DB=db,spacing=sp,eta=eta,**p))
    dbidx={12:100,16:101,20:102,25:103,28:104}[db]
    vb.append(f'''    PassiveFactor = {eta}
    d = Fixture({tb}, {tz}, 3.5, 0.5): d.tt = {tt}
    val = StemShearStressEnvelope(d, SectionDepth(d.tb, {dbidx}), x)
    Call Near("stem envelope case{i} stress", val, {p['max_v']:.12f})
    Call Near("stem envelope case{i} height", x, {p['shear_height']:.12f})
    Call Near("stem envelope case{i} common checker", MemberShearStress(d, 0, SectionDepth(d.tb, {dbidx})), {p['max_v']:.12f})
    Print #f, "{i}," & CsvNumber(SectionShear(d, 0, SectionDepth(d.tb, {dbidx})) / (10# * SectionDepth(d.tb, {dbidx}))) & "," & CsvNumber(val) & "," & CsvNumber(x)
''')
vb.append('''    Close #f
    ' A synthetic 1.55 limit isolates the old false acceptance: the face and
    ' both base sections are below 1.55, while the interior stem is above it.
    PassiveFactor = 1: AllowableShear = 1.55
    d = Fixture(0.4, 0.45, 2.5, 0.8): d.ASst_DB = 101
    Call AssertTrue("stem face alone misses shear failure", SectionShear(d, 0, SectionDepth(d.tb, 101)) / (10# * SectionDepth(d.tb, 101)) < AllowableShear)
    Call AssertTrue("shear fixture toe independently below limit", MemberShearStress(d, 1, SectionDepth(d.TBase, 104)) < AllowableShear)
    Call AssertTrue("shear fixture heel independently below limit", MemberShearStress(d, 2, SectionDepth(d.TBase, 104)) < AllowableShear)
    Call AssertTrue("interior stem shear failure rejected", Not Valid(d) And LastValidationReason = "CONCRETE_SHEAR")
    On Error Resume Next
    val = StemShearStressEnvelope(d, 0.19, x)
    Call AssertTrue("invalid local stem depth raises", Err.Number <> 0)
    Err.Clear
    On Error GoTo Fatal
    PassiveFactor = 0: AllowableShear = 8
''')
(P/'force_chain_checks.inc').write_bytes(''.join(vb).encode('ascii'))
(P/'force-chain-independent.json').write_text(json.dumps(cases,indent=2),encoding='ascii')
(P/'stem-shear-independent.json').write_text(json.dumps(profiles,indent=2),encoding='ascii')
