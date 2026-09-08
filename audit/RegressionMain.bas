Attribute VB_Name = "RegressionMain"
Option Explicit
Private output As Integer
Private failures As Long
Private checks As Long

Private Sub AssertTrue(label As String, value As Boolean)
    checks = checks + 1
    If value Then
        Print #output, "OK: " & label
    Else
        Print #output, "FAIL: " & label
        failures = failures + 1
    End If
End Sub

Private Sub Near(label As String, actual As Double, expected As Double)
    Call AssertTrue(label, Abs(actual - expected) < 0.00001)
    Print #output, "  actual=" & CsvNumber(actual) & "; expected=" & CsvNumber(expected)
End Sub

Private Function Fixture(tb As Double, baseDepth As Double, width As Double, toe As Double) As Design
    Dim d As Design
    d.tt = 0.2: d.tb = tb: d.TBase = baseDepth: d.Base = width: d.LToe = toe
    d.LHeel = width - toe - tb
    d.ASst_DB = 104: d.ASst_Sp = 110
    d.AStoe_DB = 104: d.AStoe_Sp = 110
    d.ASheel_DB = 104: d.ASheel_Sp = 110
    Fixture = d
End Function

Private Function Valid(d As Design) As Boolean
    Dim ot As Double, sl As Double, bc As Double
    Valid = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, ot, sl, bc)
    Print #output, "  validator=" & LastValidationReason
End Function

Public Sub Main()
    Dim d As Design, d2 As Design, a As Design, b As Design
    Dim x As Double, e As Double, qt As Double, qh As Double
    Dim ca As Double, sa As Double, ja As Double, val As Double, ok As Boolean
    Dim mat As MaterialProperties, pathA As String, pathB As String, f As Integer, textA As String, textB As String
    On Error GoTo Fatal
    output = FreeFile
    Open App.Path & "\native-regression.txt" For Output As #output
    Print #output, "ACTUAL compiled VB6 regression. Synthetic criteria are TEST FIXTURES, NOT EIT 2562 compliance."
    InitializeArrays
    Call AssertTrue("project initialization includes full passive", PassiveFactor = 1#)
    gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 20: cover = 0.075
    mat = GetSD40Material(320, 2601, 24): currentMaterial = mat
    currentWSD = CalculateWSDParameters(4000, 320)
    Dim refParams As WSDParams, refFile As Integer
    refFile = FreeFile
    Open App.Path & "\wsd-reference-native.csv" For Output As #refFile
    Print #refFile, "fc_prime,fy,n,k_bal,j_bal,R_bal,rho_bal"
    refParams = CalculateWSDParameters(3000, 180)
    Call Near("project n9 fc180 n", refParams.n, 9.000000000000)
    Call Near("project n9 fc180 k", refParams.k, 0.327052489906)
    Call Near("project n9 fc180 j", refParams.j, 0.890982503365)
    Call Near("project n9 fc180 R", refParams.R, 11.801620870611)
    Call Near("project n9 fc180 balanced rho", CalculateRhoBalanced(refParams), 0.008830417227)
    Print #refFile, "180,3000," & CsvNumber(refParams.n) & "," & CsvNumber(refParams.k) & "," & CsvNumber(refParams.j) & "," & CsvNumber(refParams.R) & "," & CsvNumber(CalculateRhoBalanced(refParams))
    Call SectionStresses(CalculateMomentCapacity(refParams.R, 0.3), 0.3, CalculateRhoBalanced(refParams) * 100# * 30#, refParams.n, ca, sa, ja)
    Call Near("project n9 fc180 actual balanced ca", ca, 81.000000000000)
    Call Near("project n9 fc180 actual balanced sa", sa, 1500.000000000000)
    Call Near("project n9 fc180 actual balanced ja", ja, 0.890982503365)
    refParams = CalculateWSDParameters(4000, 210)
    Call Near("project n9 fc210 n", refParams.n, 9.000000000000)
    Call Near("project n9 fc210 k", refParams.k, 0.333464026661)
    Call Near("project n9 fc210 j", refParams.j, 0.888845324446)
    Call Near("project n9 fc210 R", refParams.R, 14.004802710786)
    Call Near("project n9 fc210 balanced rho", CalculateRhoBalanced(refParams), 0.009268338388)
    Print #refFile, "210,4000," & CsvNumber(refParams.n) & "," & CsvNumber(refParams.k) & "," & CsvNumber(refParams.j) & "," & CsvNumber(refParams.R) & "," & CsvNumber(CalculateRhoBalanced(refParams))
    Call SectionStresses(CalculateMomentCapacity(refParams.R, 0.3), 0.3, CalculateRhoBalanced(refParams) * 100# * 30#, refParams.n, ca, sa, ja)
    Call Near("project n9 fc210 actual balanced ca", ca, 94.500000000000)
    Call Near("project n9 fc210 actual balanced sa", sa, 1700.000000000000)
    Call Near("project n9 fc210 actual balanced ja", ja, 0.888845324446)
    refParams = CalculateWSDParameters(4000, 240)
    Call Near("project n9 fc240 n", refParams.n, 9.000000000000)
    Call Near("project n9 fc240 k", refParams.k, 0.363772455090)
    Call Near("project n9 fc240 j", refParams.j, 0.878742514970)
    Call Near("project n9 fc240 R", refParams.R, 17.261765391373)
    Call Near("project n9 fc240 balanced rho", CalculateRhoBalanced(refParams), 0.011555125044)
    Print #refFile, "240,4000," & CsvNumber(refParams.n) & "," & CsvNumber(refParams.k) & "," & CsvNumber(refParams.j) & "," & CsvNumber(refParams.R) & "," & CsvNumber(CalculateRhoBalanced(refParams))
    Call SectionStresses(CalculateMomentCapacity(refParams.R, 0.3), 0.3, CalculateRhoBalanced(refParams) * 100# * 30#, refParams.n, ca, sa, ja)
    Call Near("project n9 fc240 actual balanced ca", ca, 108.000000000000)
    Call Near("project n9 fc240 actual balanced sa", sa, 1700.000000000000)
    Call Near("project n9 fc240 actual balanced ja", ja, 0.878742514970)
    refParams = CalculateWSDParameters(4000, 280)
    Call Near("project n9 fc280 n", refParams.n, 9.000000000000)
    Call Near("project n9 fc280 k", refParams.k, 0.400141143260)
    Call Near("project n9 fc280 j", refParams.j, 0.866619618913)
    Call Near("project n9 fc280 R", refParams.R, 21.846520400281)
    Call Near("project n9 fc280 balanced rho", CalculateRhoBalanced(refParams), 0.014828760015)
    Print #refFile, "280,4000," & CsvNumber(refParams.n) & "," & CsvNumber(refParams.k) & "," & CsvNumber(refParams.j) & "," & CsvNumber(refParams.R) & "," & CsvNumber(CalculateRhoBalanced(refParams))
    Call SectionStresses(CalculateMomentCapacity(refParams.R, 0.3), 0.3, CalculateRhoBalanced(refParams) * 100# * 30#, refParams.n, ca, sa, ja)
    Call Near("project n9 fc280 actual balanced ca", ca, 126.000000000000)
    Call Near("project n9 fc280 actual balanced sa", sa, 1700.000000000000)
    Call Near("project n9 fc280 actual balanced ja", ja, 0.866619618913)
    refParams = CalculateWSDParameters(4000, 320)
    Call Near("project n9 fc320 n", refParams.n, 9.000000000000)
    Call Near("project n9 fc320 k", refParams.k, 0.432576769025)
    Call Near("project n9 fc320 j", refParams.j, 0.855807743658)
    Call Near("project n9 fc320 R", refParams.R, 26.654583503416)
    Call Near("project n9 fc320 balanced rho", CalculateRhoBalanced(refParams), 0.018320898453)
    Print #refFile, "320,4000," & CsvNumber(refParams.n) & "," & CsvNumber(refParams.k) & "," & CsvNumber(refParams.j) & "," & CsvNumber(refParams.R) & "," & CsvNumber(CalculateRhoBalanced(refParams))
    Call SectionStresses(CalculateMomentCapacity(refParams.R, 0.3), 0.3, CalculateRhoBalanced(refParams) * 100# * 30#, refParams.n, ca, sa, ja)
    Call Near("project n9 fc320 actual balanced ca", ca, 144.000000000000)
    Call Near("project n9 fc320 actual balanced sa", sa, 1700.000000000000)
    Call Near("project n9 fc320 actual balanced ja", ja, 0.855807743658)
    Close #refFile
    refParams = CalculateWSDParameters(4000, 320)
    Call SectionStresses(9.7262, 0.317, CalculateAsProv(101, 110), refParams.n, ca, sa, ja)
    Call Near("actual DB16 section ca", ca, 74.912752644519)
    Call Near("actual DB16 section sa", sa, 1686.553404336288)
    Call Near("actual DB16 section ja", ja, 0.904802911182)
    Call AssertTrue("actual bar j differs from balanced j", Abs(ja - refParams.j) > 0.01)
    On Error Resume Next
    refParams = CalculateWSDParameters(4000, 0)
    Call AssertTrue("zero concrete strength rejected", Err.Number <> 0)
    Err.Clear
    On Error GoTo Fatal

    H = 5: H1 = 1.2: PassiveFactor = 0
    d = Fixture(0.2, 0.3, 3.5, 0.5)
    Call Near("stem height H-TBase", H - d.TBase, 4.7)
    Call Near("requested active moment", CalculateMomentStem(d), 10.3823)
    PassiveFactor = 1
    Call Near("requested full passive net moment", CalculateMomentStem(d), 9.7262)
    PassiveFactor = 0
    Call Near("actual DB12 effective depth", SectionDepth(0.2, 100), 0.119)
    Call AssertTrue("requested DB12 rejected in flexure", Not CheckSteelOK(CalculateMomentStem(d), SectionDepth(0.2, 100), 100, 113))
    Call SectionStresses(5, 0.119, CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call AssertTrue("concrete-only failure fixture", ca > currentWSD.fc And sa < currentWSD.fs)
    Call AssertTrue("concrete compression checked independently", Not CheckSteelOK(5, 0.119, 104, 110))
    d.ASst_DB = 100: d.ASst_Sp = 113
    Call AssertTrue("requested full design rejected", Not Valid(d))
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    Call AssertTrue("unverified criteria never accepted", Not Valid(d))
    Call AssertTrue("explicit unverified status", LastValidationReason = "WSD_CRITERIA_UNVERIFIED")
    ' Diagnostic reports use production UNSET criteria, never synthetic limits.
    WSDReviewed = False: WSDSource = ""
    AllowableShear = 0: MinStemRatio = 0: MinBaseRatio = 0
    f = FreeFile
    Open App.Path & "\h5-vb6-checks.md" For Output As #f
    Print #f, "# H5 actual VB6 per-check reports"
    Print #f, "Generated by compiled Regression.exe using the actual project modules."
    Print #f, "No synthetic shear/minimum thresholds. No EIT compliance claim."
    Print #f, ""
    d = Fixture(0.2, 0.3, 3.5, 0.5)
    d.ASst_DB = 100: d.ASst_Sp = 113
    textA = BuildDesignCheckReport(d, True)
    Print #f, "## Reference DB12 at 0.25 m - active only"
    Print #f, textA
    Call AssertTrue("weak audit necessary yield rejection", InStr(textA, "AUDIT RESULT: FAIL_NECESSARY_YIELD_BOUND") > 0)
    Call AssertTrue("weak audit continues to shear and minimum after failure", InStr(textA, "Stem nominal shear") > 0 And InStr(textA, "Stem minimum steel") > 0)
    Call AssertTrue("unset shear is not a zero limit failure", InStr(textA, "UNSET | UNVERIFIED") > 0)
    Call AssertTrue("stem-only audit does not invent base steel", InStr(textA, "Toe bars") = 0 And InStr(textA, "OUT_OF_SCOPE") > 0)
    PassiveFactor = 1
    textA = BuildDesignCheckReport(d, True)
    Print #f, "## Reference DB12 at 0.25 m - full passive sensitivity"
    Print #f, textA
    Call AssertTrue("full passive weak audit still rejected", InStr(textA, "FAIL_NECESSARY_YIELD_BOUND") > 0)
    PassiveFactor = 0
    ' Fixed candidate from the previous synthetic GUI BA search; re-evaluated
    ' here with all unsourced criteria UNSET. This is NOT a new accepted design.
    d = Fixture(0.45, 0.45, 3, 0.6): d.tt = 0.25
    d.ASst_DB = 103: d.ASst_Sp = 112
    d.AStoe_DB = 103: d.AStoe_Sp = 113
    d.ASheel_DB = 102: d.ASheel_Sp = 111
    textA = BuildDesignCheckReport(d)
    Print #f, "## Prior BA candidate - rechecked with unverified criteria"
    Print #f, textA
    Close #f
    Call AssertTrue("BA diagnostic indeterminate, not certified", InStr(textA, "AUDIT RESULT: INDETERMINATE_WSD") > 0)
    Call AssertTrue("audit checks all three members", InStr(textA, "Stem steel stress") > 0 And InStr(textA, "Toe steel stress") > 0 And InStr(textA, "Heel steel stress") > 0)
    Call AssertTrue("audit shows missing detailing", InStr(textA, "anchorage, distribution steel") > 0)
    Call AssertTrue("audit does not enable criteria", Not WSDCriteriaReady() And AllowableShear = 0)
    Call AssertTrue("audit agrees candidate remains unaccepted", Not Valid(d) And LastValidationReason = "WSD_CRITERIA_UNVERIFIED")
    Call AssertTrue("shared display contains audit", InStr(FormatResults(d, mat), "AUDIT RESULT: INDETERMINATE_WSD") > 0)
    d.ASst_Sp = 0
    textA = BuildDesignCheckReport(d)
    Call AssertTrue("bad steel input diagnosed without losing heel checks", InStr(textA, "INVALID_DATA") > 0 And InStr(textA, "Heel steel stress") > 0)
    d = Fixture(0.6, 0.7, 1.5, 0.5)
    textA = BuildDesignCheckReport(d)
    Call AssertTrue("partial contact audit rejects base and retains stem", InStr(textA, "FAIL_MODEL") > 0 And InStr(textA, "Stem steel stress") > 0 And InStr(textA, "Toe bars") = 0)
    d = Fixture(0.6, 0.7, 3, 1.2)
    d.LHeel = d.LToe + 0.000000000000001
    Call AssertTrue("equal heel/toe roundoff rejected", Not CheckHeelLayout(d))
    d.LHeel = d.LToe + 0.025: d.Base = d.LToe + d.tb + d.LHeel
    Call AssertTrue("distinct heel/toe preserved", CheckHeelLayout(d))

    ' Synthetic thresholds exercise branches only; no normative claim.
    WSDReviewed = True: WSDSource = "SYNTHETIC TEST FIXTURE ONLY - not EIT"
    AllowableShear = 8: MinStemRatio = 0.0015: MinBaseRatio = 0.0015
    H = 3: H1 = 1.2: PassiveFactor = 0
    d = Fixture(0.35, 0.4, 2, 0.3)
    Call Near("case0 stem moment", CalculateMomentStem(d), 1.757600000000)
    Call Near("case0 W1", CalculateW1(d, x), 0.465230769231)
    Call Near("case0 x1", x, 0.161813186813)
    Call Near("case0 W3", CalculateW3(d, x), 1.716000000000)
    Call Near("case0 x3", x, 0.509090909091)
    Call AssertTrue("case0 contact", BearingEdges(d, e, qt, qh))
    Call Near("case0 signed e", e, 0.180339637107)
    Call Near("case0 q toe", qt, 8.028115828402)
    Call Near("case0 q heel", qh, 2.391114940828)
    Call Near("case0 toe moment", CalculateMomentToe(d), 0.240581960281)
    Call Near("case0 heel moment", CalculateMomentHeel(d), 1.804785171940)
    Call Near("case0 stem shear", SectionShear(d, 0, SectionDepth(d.tb, 104)), 2.028000000000)
    Call Near("case0 toe shear", SectionShear(d, 1, SectionDepth(d.TBase, 104)), 0.000000000000)
    Call Near("case0 heel shear", SectionShear(d, 2, SectionDepth(d.TBase, 104)), 1.854276367691)
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("case0 concrete stress", ca, 12.950557826766)
    Call Near("case0 steel stress", sa, 129.836787948195)

    Call AssertTrue("H3 synthetic fixture valid", Valid(d))
    d.tb = 0.2: d.LHeel = d.Base - d.LToe - d.tb: d.ASst_DB = 100: d.ASst_Sp = 113
    Call AssertTrue("H3 weak section rejected", Not Valid(d))

    H = 4: H1 = 1.2: PassiveFactor = 0
    d = Fixture(0.45, 0.5, 2.5, 0.4)
    Call Near("case1 stem moment", CalculateMomentStem(d), 4.287500000000)
    Call Near("case1 W1", CalculateW1(d, x), 0.535500000000)
    Call Near("case1 x1", x, 0.212745098039)
    Call Near("case1 W3", CalculateW3(d, x), 2.730000000000)
    Call Near("case1 x3", x, 0.679487179487)
    Call AssertTrue("case1 contact", BearingEdges(d, e, qt, qh))
    Call Near("case1 signed e", e, 0.245795444314)
    Call Near("case1 q toe", qt, 10.595472000000)
    Call Near("case1 q heel", qh, 2.732928000000)
    Call Near("case1 toe moment", CalculateMomentToe(d), 0.617290905600)
    Call Near("case1 heel moment", CalculateMomentHeel(d), 4.134541395600)
    Call Near("case1 stem shear", SectionShear(d, 0, SectionDepth(d.tb, 104)), 3.675000000000)
    Call Near("case1 toe shear", SectionShear(d, 1, SectionDepth(d.TBase, 104)), 0.000000000000)
    Call Near("case1 heel shear", SectionShear(d, 2, SectionDepth(d.TBase, 104)), 3.492410926435)
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("case1 concrete stress", ca, 18.164072365559)
    Call Near("case1 steel stress", sa, 224.407767128530)

    Call AssertTrue("H4 synthetic fixture valid", Valid(d))
    d.tb = 0.2: d.LHeel = d.Base - d.LToe - d.tb: d.ASst_DB = 100: d.ASst_Sp = 113
    Call AssertTrue("H4 weak section rejected", Not Valid(d))

    H = 5: H1 = 1.2: PassiveFactor = 0
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    Call Near("case2 stem moment", CalculateMomentStem(d), 7.950700000000)
    Call Near("case2 W1", CalculateW1(d, x), 0.470930232558)
    Call Near("case2 x1", x, 0.261800172265)
    Call Near("case2 W3", CalculateW3(d, x), 4.128000000000)
    Call Near("case2 x3", x, 0.883333333333)
    Call AssertTrue("case2 contact", BearingEdges(d, e, qt, qh))
    Call Near("case2 signed e", e, 0.225835623712)
    Call Near("case2 q toe", qt, 11.515272290593)
    Call Near("case2 q heel", qh, 5.087544985155)
    Call Near("case2 toe moment", CalculateMomentToe(d), 1.078648754744)
    Call Near("case2 heel moment", CalculateMomentHeel(d), 8.246189382260)
    Call Near("case2 stem shear", SectionShear(d, 0, SectionDepth(d.tb, 104)), 5.547000000000)
    Call Near("case2 toe shear", SectionShear(d, 1, SectionDepth(d.TBase, 104)), 0.000000000000)
    Call Near("case2 heel shear", SectionShear(d, 2, SectionDepth(d.TBase, 104)), 4.811893989654)
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("case2 concrete stress", ca, 18.785591577232)
    Call Near("case2 steel stress", sa, 288.204163288811)

    Call AssertTrue("H5 synthetic fixture valid", Valid(d))
    d.tb = 0.2: d.LHeel = d.Base - d.LToe - d.tb: d.ASst_DB = 100: d.ASst_Sp = 113
    Call AssertTrue("H5 weak section rejected", Not Valid(d))

    H = 3: H1 = 2: PassiveFactor = 1
    d = Fixture(0.6, 0.4, 4, 0.2)
    Call Near("case3 stem moment", CalculateMomentStem(d), -1.928800000000)
    Call Near("case3 W1", CalculateW1(d, x), 0.930461538462)
    Call Near("case3 x1", x, 0.169352869353)
    Call Near("case3 W3", CalculateW3(d, x), 2.496000000000)
    Call Near("case3 x3", x, 0.583333333333)
    Call AssertTrue("case3 contact", BearingEdges(d, e, qt, qh))
    Call Near("case3 signed e", e, -0.236082379882)
    Call Near("case3 q toe", qt, 3.591470414201)
    Call Near("case3 q heel", qh, 7.529760355030)
    Call Near("case3 toe moment", CalculateMomentToe(d), -0.003657828402)
    Call Near("case3 heel moment", CalculateMomentHeel(d), -4.298494485207)
    Call Near("case3 stem shear", SectionShear(d, 0, SectionDepth(d.tb, 104)), 4.884000000000)
    Call Near("case3 toe shear", SectionShear(d, 1, SectionDepth(d.TBase, 104)), 0.000000000000)
    Call Near("case3 heel shear", SectionShear(d, 2, SectionDepth(d.TBase, 104)), 1.350738661027)
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("case3 concrete stress", ca, 4.557290431555)
    Call Near("case3 steel stress", sa, 69.916886582497)

    Call AssertTrue("reversed tension faces rejected", Not Valid(d))

    H = 5: H1 = 1.2: PassiveFactor = 0
    d = Fixture(0.6, 0.3, 4, 1)
    Call Near("case4 stem moment", CalculateMomentStem(d), 10.382300000000)
    Call Near("case4 W1", CalculateW1(d, x), 1.682042553191)
    Call Near("case4 x1", x, 0.519384373910)
    Call Near("case4 W3", CalculateW3(d, x), 4.512000000000)
    Call Near("case4 x3", x, 1.383333333333)
    Call AssertTrue("case4 contact", BearingEdges(d, e, qt, qh))
    Call Near("case4 signed e", e, 0.052068087424)
    Call Near("case4 q toe", qt, 7.918132571299)
    Call Near("case4 q heel", qh, 6.770888705297)
    Call Near("case4 toe moment", CalculateMomentToe(d), 2.741264457900)
    Call Near("case4 heel moment", CalculateMomentHeel(d), 6.277428061928)
    Call Near("case4 stem shear", SectionShear(d, 0, SectionDepth(d.tb, 104)), 6.627000000000)
    Call Near("case4 toe shear", SectionShear(d, 1, SectionDepth(d.TBase, 104)), 4.311873673917)
    Call Near("case4 heel shear", SectionShear(d, 2, SectionDepth(d.TBase, 104)), 4.586385558500)
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("case4 concrete stress", ca, 24.530877461393)
    Call Near("case4 steel stress", sa, 376.346998945177)

    H = 5: H1 = 1.2: PassiveFactor = 0
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    AllowableShear = 0.0001
    Call AssertTrue("shear limit enforced", Not Valid(d))
    Call AssertTrue("shear rejection reason", LastValidationReason = "CONCRETE_SHEAR")
    AllowableShear = 8: MinStemRatio = 0.1
    Call AssertTrue("minimum steel enforced", Not Valid(d))
    Call AssertTrue("minimum steel rejection reason", LastValidationReason = "MINIMUM_STEEL")
    MinStemRatio = 0.0015
    d.tb = 0.07: d.tt = 0.07: d.LHeel = d.Base - d.LToe - d.tb
    Call AssertTrue("negative depth invalid, not clamped", Not Valid(d))
    On Error Resume Next
    val = CalculateEffectiveDepth(0.08, 0.075, 28)
    Call AssertTrue("depth helper raises error", Err.Number <> 0)
    Err.Clear
    On Error GoTo Fatal
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    d.LHeel = d.LHeel + 0.1
    Call AssertTrue("inconsistent geometry rejected", Not Valid(d))
    d = Fixture(0.6, 0.7, 3.5, 0.5): H1 = 0.2
    Call AssertTrue("front soil below top rejected", Not Valid(d))
    H1 = 1.2
    d = Fixture(0.6, 0.7, 1.5, 0.5)
    Call AssertTrue("partial contact rejected", Not BearingEdges(d, e, qt, qh))
    d = Fixture(0.6, 0.7, 3.5, 0.5): d.UseDoubleStem = True
    Call AssertTrue("unsupported double layers rejected", Not Valid(d))
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    Call BeginSearch(3, 123, "ENTRY-TEST")
    ok = EvaluateCandidate(d, "initial", val)
    Call AssertTrue("initial updates global best", ok And RunBestEvaluation = 1)
    d.ASst_DB = 103: d.AStoe_DB = 103: d.ASheel_DB = 103
    ok = EvaluateCandidate(d, "reset", val)
    Call AssertTrue("reset updates global best", ok And RunBestEvaluation = 2)
    d.ASst_DB = 102: d.AStoe_DB = 102: d.ASheel_DB = 102
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("neighbor updates global best", ok And RunBestEvaluation = 3)
    Call AssertTrue("all entries counted", EvaluationCount = 3)
    On Error Resume Next
    ok = EvaluateCandidate(d, "excess", val)
    Call AssertTrue("budget cannot be exceeded", Err.Number <> 0 And EvaluationCount = 3)
    Err.Clear
    On Error GoTo Fatal
    Call FinishSearch
    ' Cost selection is downstream of every safety/structural check.
    ' Synthetic structural thresholds here are branch fixtures only.
    H = 5: H1 = 1.2: mu = 0.6: qa = 20
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    Call BeginSearch(5, 123, "FEASIBLE-COST-TEST")
    ok = EvaluateCandidate(d, "initial", val)
    Call AssertTrue("feasible-cost baseline accepted in fixture", ok And RunBestEvaluation = 1)
    d = Fixture(0.6, 0.7, 2, 0.5)
    Call AssertTrue("OT rejection fixture is cheaper", CalculateCost(d) < RunBestCost)
    Call AssertTrue("OT rejection fixture violates overturning", Not CheckFS_OT(d, x))
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("OT failure not priced or selected", Not ok And val = NO_SOLUTION_COST And RunBestEvaluation = 1)
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    d.ASst_DB = 102: d.AStoe_DB = 102: d.ASheel_DB = 102
    Call AssertTrue("SL/BC rejection fixture is cheaper", CalculateCost(d) < RunBestCost)
    mu = 0.01
    Call AssertTrue("SL rejection fixture violates sliding", Not CheckFS_SL(d, x))
    ok = EvaluateCandidate(d, "reset", val)
    Call AssertTrue("SL failure not priced or selected", Not ok And val = NO_SOLUTION_COST And RunBestEvaluation = 1)
    mu = 0.6: qa = 1
    Call AssertTrue("BC rejection fixture violates bearing", Not CheckFS_BC(d, x, e, qt, qh))
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("BC failure not priced or selected", Not ok And val = NO_SOLUTION_COST And RunBestEvaluation = 1)
    qa = 20
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("cheaper candidate selected after all checks satisfied", ok And RunBestEvaluation = 5 And val = RunBestCost)
    Call FinishSearch

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

    ' One-evaluation regression: initial must be returned even with no neighbors.
    a = BisectionOptimization(1, 3, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, _
        True, 1, 23, 42, 61, 80, 104, 110, 104, 110, 104, 110, 42)
    Call AssertTrue("BA initial returned at budget=1", a.IsValid And EvaluationCount = 1 And RunBestEvaluation = 1)
    b = HillClimbingOptimization(1, 3, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, _
        True, 1, 23, 42, 61, 80, 104, 110, 104, 110, 104, 110, 42)
    Call AssertTrue("HCA initial returned at budget=1", b.IsValid And EvaluationCount = 1 And RunBestEvaluation = 1)
    Call Near("identical shared initial cost", a.TotalCost, b.TotalCost)
    a = BisectionOptimization(160, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=42)
    Call AssertTrue("BA evaluation budget", EvaluationCount = 160)
    pathA = RunFolder
    b = BisectionOptimization(160, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=42)
    pathB = RunFolder
    Call AssertTrue("trial paths never overwrite", pathA <> pathB)
    Call Near("BA same seed same best", a.TotalCost, b.TotalCost)
    Call AssertTrue("BA same seed entire trace", ReadText(pathA & "\evaluations.csv") = ReadText(pathB & "\evaluations.csv"))
    a = HillClimbingOptimization(160, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=42)
    Call AssertTrue("HCA equal evaluation budget", EvaluationCount = 160)
    pathA = RunFolder
    b = HillClimbingOptimization(160, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=42)
    Call AssertTrue("HCA same seed entire trace", ReadText(pathA & "\evaluations.csv") = ReadText(RunFolder & "\evaluations.csv"))
    WSDReviewed = False
    a = BisectionOptimization(80, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=7)
    Call AssertTrue("infeasible BA reopens original three ranges", RunRecoveryCount > 0)
    Call AssertTrue("infeasible BA keeps exact budget", EvaluationCount = 80)
    Call AssertTrue("BA no solution explicit", Not a.IsValid And RunStatus = "NO_SOLUTION" And RunBestEvaluation = 0)
    Call AssertTrue("no solution is not free", CalculateCost(a) = NO_SOLUTION_COST)
    Call AssertTrue("no solution report", InStr(FormatResults(a, mat), "NO_SOLUTION") > 0)
    b = HillClimbingOptimization(80, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=7)
    Call AssertTrue("HCA no solution and exact budget", Not b.IsValid And EvaluationCount = 80 And RunStatus = "NO_SOLUTION")
    Print #output, "TOTAL checks=" & checks & "; failures=" & failures
    Close #output
    Exit Sub
Fatal:
    If output > 0 Then
        Print #output, "FATAL " & Err.Number & ": " & Err.Description
        Close #output
    End If
End Sub

Private Function ReadText(path As String) As String
    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    ReadText = Space$(LOF(f))
    Get #f, , ReadText
    Close #f
End Function
