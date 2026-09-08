from pathlib import Path
P=Path(__file__).resolve().parent
def put(name,s): (P/name).write_bytes(s.replace('\n','\r\n').encode('ascii'))
put('Regression.vbp',r'''Type=Exe
Module=RegressionMain; RegressionMain.bas
Module=modDataStructures; ..\modDataStructures.bas
Module=modWSD; ..\modWSD.bas
Module=modShared; ..\modShared.bas
Module=modBA; ..\modBA.bas
Module=modHillClimbing; ..\modHillClimbing.bas
Startup="Sub Main"
Name="NativeRegression"
ExeName32="Regression.exe"
''')
start='''Attribute VB_Name = "RegressionMain"
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
    Open App.Path & "\\native-regression.txt" For Output As #output
    Print #output, "ACTUAL compiled VB6 regression. Synthetic criteria are TEST FIXTURES, NOT EIT 2562 compliance."
    InitializeArrays
    gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 20: cover = 0.075
    mat = GetSD40Material(320, 2601, 24): currentMaterial = mat
    currentWSD = CalculateWSDParameters(4000, 320)
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
    ' Synthetic thresholds exercise branches only; no normative claim.
    WSDReviewed = True: WSDSource = "SYNTHETIC TEST FIXTURE ONLY - not EIT"
    AllowableShear = 8: MinStemRatio = 0.0015: MinBaseRatio = 0.0015
'''
end='''
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
    Call AssertTrue("BA same seed entire trace", ReadText(pathA & "\\evaluations.csv") = ReadText(pathB & "\\evaluations.csv"))
    a = HillClimbingOptimization(160, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=42)
    Call AssertTrue("HCA equal evaluation budget", EvaluationCount = 160)
    pathA = RunFolder
    b = HillClimbingOptimization(160, 5, 1.2, 1.8, 2.4, 30, 0.6, 20, 0.075, mat, RandomSeed:=42)
    Call AssertTrue("HCA same seed entire trace", ReadText(pathA & "\\evaluations.csv") = ReadText(RunFolder & "\\evaluations.csv"))
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
'''
put('RegressionMain.bas',start+(P/'fixture_checks.inc').read_text()+end)
