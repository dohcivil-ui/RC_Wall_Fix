    ' Force path: same geometry, change passive only; no synthetic criteria used here.
    Dim chainW As Double, cw1 As Double, cw2 As Double, cw3 As Double, cw4 As Double
    Dim cx1 As Double, cx2 As Double, cx3 As Double, cx4 As Double
    Dim cot As Double, csl As Double, cbc As Double, cqmax As Double, cqmin As Double
    H = 5: H1 = 1.2: qa = 30: mu = 0.6
    currentMaterial.concretePrice = GetConcretePrice(320)
    f = FreeFile
    Open App.Path & "\force-chain-native.csv" For Output As #f
    Print #f, "eta,Pa,Pp,active_arm,passive_arm,MO,MR,W,xR,e,qtoe,qheel,Mstem,Mtoe,Mheel,FSot,FSsl,FSbc,cost"
    PassiveFactor = 0
    d = Fixture(0.45, 0.3, 3, 1.1)
    d.ASst_DB = 103: d.ASst_Sp = 113
    d.AStoe_DB = 102: d.AStoe_Sp = 113
    d.ASheel_DB = 102: d.ASheel_Sp = 112
    chainW = CalculateWTotal(d, cw1, cw2, cw3, cw4, cx1, cx2, cx3, cx4)
    Call AssertTrue("force-chain eta=0 full contact", BearingEdges(d, e, qt, qh))
    Call CheckFS_OT(d, cot): Call CheckFS_SL(d, csl)
    Call CheckFS_BC(d, cbc, x, cqmax, cqmin)
    Call Near("force-chain eta=0 Pa", CalculatePa(), 7.500000000000)
    Call Near("force-chain eta=0 Pp", CalculatePp(), 0.000000000000)
    Call Near("force-chain eta=0 active_arm", H / 3#, 1.666666666667)
    Call Near("force-chain eta=0 passive_arm", H1 / 3#, 0.400000000000)
    Call Near("force-chain eta=0 MO", CalculateMO(d), 12.500000000000)
    Call Near("force-chain eta=0 MR", CalculateMR(d), 37.227998030783)
    Call Near("force-chain eta=0 W", chainW, 19.913776595745)
    Call Near("force-chain eta=0 xR", d.Base / 2# - e, 1.241753311427)
    Call Near("force-chain eta=0 e", e, 0.258246688573)
    Call Near("force-chain eta=0 qtoe", qt, 10.066370107137)
    Call Near("force-chain eta=0 qheel", qh, 3.209480956692)
    Call Near("force-chain eta=0 Mstem", CalculateMomentStem(d), 10.382300000000)
    Call Near("force-chain eta=0 Mtoe", CalculateMomentToe(d), 4.167425055971)
    Call Near("force-chain eta=0 Mheel", CalculateMomentHeel(d), 5.115170161706)
    Call Near("force-chain eta=0 FSot", cot, 2.978239842463)
    Call Near("force-chain eta=0 FSsl", csl, 1.593102127660)
    Call Near("force-chain eta=0 FSbc", cbc, 2.980220246296)
    Call Near("force-chain eta=0 cost", CalculateCost(d), 9144.075500000001)
    Call Near("force-chain eta=0 vertical equilibrium", (qt + qh) * d.Base / 2#, 19.913776595745)
    Call Near("force-chain eta=0 moment equilibrium", d.Base ^ 2 * (qt + 2# * qh) / 6#, 24.727998030783)
    textA = CStr(PassiveFactor)
    textA = textA & "," & CsvNumber(CalculatePa())
    textA = textA & "," & CsvNumber(CalculatePp())
    textA = textA & "," & CsvNumber(H / 3#)
    textA = textA & "," & CsvNumber(H1 / 3#)
    textA = textA & "," & CsvNumber(CalculateMO(d))
    textA = textA & "," & CsvNumber(CalculateMR(d))
    textA = textA & "," & CsvNumber(chainW)
    textA = textA & "," & CsvNumber(d.Base / 2# - e)
    textA = textA & "," & CsvNumber(e)
    textA = textA & "," & CsvNumber(qt)
    textA = textA & "," & CsvNumber(qh)
    textA = textA & "," & CsvNumber(CalculateMomentStem(d))
    textA = textA & "," & CsvNumber(CalculateMomentToe(d))
    textA = textA & "," & CsvNumber(CalculateMomentHeel(d))
    textA = textA & "," & CsvNumber(cot)
    textA = textA & "," & CsvNumber(csl)
    textA = textA & "," & CsvNumber(cbc)
    textA = textA & "," & CsvNumber(CalculateCost(d))
    Print #f, textA
    PassiveFactor = 1
    d = Fixture(0.45, 0.3, 3, 1.1)
    d.ASst_DB = 103: d.ASst_Sp = 113
    d.AStoe_DB = 102: d.AStoe_Sp = 113
    d.ASheel_DB = 102: d.ASheel_Sp = 112
    chainW = CalculateWTotal(d, cw1, cw2, cw3, cw4, cx1, cx2, cx3, cx4)
    Call AssertTrue("force-chain eta=1 full contact", BearingEdges(d, e, qt, qh))
    Call CheckFS_OT(d, cot): Call CheckFS_SL(d, csl)
    Call CheckFS_BC(d, cbc, x, cqmax, cqmin)
    Call Near("force-chain eta=1 Pa", CalculatePa(), 7.500000000000)
    Call Near("force-chain eta=1 Pp", CalculatePp(), 3.888000000000)
    Call Near("force-chain eta=1 active_arm", H / 3#, 1.666666666667)
    Call Near("force-chain eta=1 passive_arm", H1 / 3#, 0.400000000000)
    Call Near("force-chain eta=1 MO", CalculateMO(d), 10.944800000000)
    Call Near("force-chain eta=1 MR", CalculateMR(d), 37.227998030783)
    Call Near("force-chain eta=1 W", chainW, 19.913776595745)
    Call Near("force-chain eta=1 xR", d.Base / 2# - e, 1.319849999543)
    Call Near("force-chain eta=1 e", e, 0.180150000457)
    Call Near("force-chain eta=1 qtoe", qt, 9.029570107137)
    Call Near("force-chain eta=1 qheel", qh, 4.246280956692)
    Call Near("force-chain eta=1 Mstem", CalculateMomentStem(d), 9.726200000000)
    Call Near("force-chain eta=1 Mtoe", CalculateMomentToe(d), 3.693492255971)
    Call Near("force-chain eta=1 Mheel", CalculateMomentHeel(d), 4.376435761706)
    Call Near("force-chain eta=1 FSot", cot, 3.102655842463)
    Call Near("force-chain eta=1 FSsl", csl, 2.111502127660)
    Call Near("force-chain eta=1 FSbc", cbc, 3.322417307141)
    Call Near("force-chain eta=1 cost", CalculateCost(d), 9144.075500000001)
    Call Near("force-chain eta=1 vertical equilibrium", (qt + qh) * d.Base / 2#, 19.913776595745)
    Call Near("force-chain eta=1 moment equilibrium", d.Base ^ 2 * (qt + 2# * qh) / 6#, 26.283198030783)
    textA = CStr(PassiveFactor)
    textA = textA & "," & CsvNumber(CalculatePa())
    textA = textA & "," & CsvNumber(CalculatePp())
    textA = textA & "," & CsvNumber(H / 3#)
    textA = textA & "," & CsvNumber(H1 / 3#)
    textA = textA & "," & CsvNumber(CalculateMO(d))
    textA = textA & "," & CsvNumber(CalculateMR(d))
    textA = textA & "," & CsvNumber(chainW)
    textA = textA & "," & CsvNumber(d.Base / 2# - e)
    textA = textA & "," & CsvNumber(e)
    textA = textA & "," & CsvNumber(qt)
    textA = textA & "," & CsvNumber(qh)
    textA = textA & "," & CsvNumber(CalculateMomentStem(d))
    textA = textA & "," & CsvNumber(CalculateMomentToe(d))
    textA = textA & "," & CsvNumber(CalculateMomentHeel(d))
    textA = textA & "," & CsvNumber(cot)
    textA = textA & "," & CsvNumber(csl)
    textA = textA & "," & CsvNumber(cbc)
    textA = textA & "," & CsvNumber(CalculateCost(d))
    Print #f, textA
    Close #f
    PassiveFactor = 0: qa = 20: currentMaterial = mat
    f = FreeFile
    Open App.Path & "\stem-shear-native.csv" For Output As #f
    Print #f, "case,face_v,max_v,height"
    PassiveFactor = 1
    d = Fixture(0.4, 0.3, 3.5, 0.5): d.tt = 0.2
    val = StemShearStressEnvelope(d, SectionDepth(d.tb, 101), x)
    Call Near("stem envelope case0 stress", val, 1.647525300350)
    Call Near("stem envelope case0 height", x, 0.571057207166)
    Call Near("stem envelope case0 common checker", MemberShearStress(d, 0, SectionDepth(d.tb, 101)), 1.647525300350)
    Print #f, "0," & CsvNumber(SectionShear(d, 0, SectionDepth(d.tb, 101)) / (10# * SectionDepth(d.tb, 101))) & "," & CsvNumber(val) & "," & CsvNumber(x)
    PassiveFactor = 1
    d = Fixture(0.2, 0.3, 3.5, 0.5): d.tt = 0.2
    val = StemShearStressEnvelope(d, SectionDepth(d.tb, 100), x)
    Call Near("stem envelope case1 stress", val, 4.095378151261)
    Call Near("stem envelope case1 height", x, 0.424999992337)
    Call Near("stem envelope case1 common checker", MemberShearStress(d, 0, SectionDepth(d.tb, 100)), 4.095378151261)
    Print #f, "1," & CsvNumber(SectionShear(d, 0, SectionDepth(d.tb, 100)) / (10# * SectionDepth(d.tb, 100))) & "," & CsvNumber(val) & "," & CsvNumber(x)
    PassiveFactor = 0
    d = Fixture(0.45, 0.3, 3.5, 0.5): d.tt = 0.2
    val = StemShearStressEnvelope(d, SectionDepth(d.tb, 103), x)
    Call Near("stem envelope case2 stress", val, 1.828137931034)
    Call Near("stem envelope case2 height", x, 0.000000000000)
    Call Near("stem envelope case2 common checker", MemberShearStress(d, 0, SectionDepth(d.tb, 103)), 1.828137931034)
    Print #f, "2," & CsvNumber(SectionShear(d, 0, SectionDepth(d.tb, 103)) / (10# * SectionDepth(d.tb, 103))) & "," & CsvNumber(val) & "," & CsvNumber(x)
    PassiveFactor = 1
    d = Fixture(0.4, 0.45, 3.5, 0.5): d.tt = 0.2
    val = StemShearStressEnvelope(d, SectionDepth(d.tb, 101), x)
    Call Near("stem envelope case3 stress", val, 1.615557602889)
    Call Near("stem envelope case3 height", x, 0.422944814704)
    Call Near("stem envelope case3 common checker", MemberShearStress(d, 0, SectionDepth(d.tb, 101)), 1.615557602889)
    Print #f, "3," & CsvNumber(SectionShear(d, 0, SectionDepth(d.tb, 101)) / (10# * SectionDepth(d.tb, 101))) & "," & CsvNumber(val) & "," & CsvNumber(x)
    Close #f
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
