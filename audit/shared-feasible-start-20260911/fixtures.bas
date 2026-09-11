Attribute VB_Name = "BestTrialProbe"
Option Explicit
Private f As Integer, checks As Long, failures As Long, calls As Long

Private Sub Verify(label As String, ok As Boolean)
    checks = checks + 1
    If Not ok Then failures = failures + 1
    Print #f, IIf(ok, "PASS ", "FAIL ") & label
End Sub

Private Function ReadAll(path As String) As String
    Dim n As Integer
    n = FreeFile
    Open path For Binary Access Read As #n
    ReadAll = Space$(LOF(n))
    If LOF(n) > 0 Then Get #n, , ReadAll
    Close #n
End Function

Private Function SameDesign(a As Design, b As Design) As Boolean
    SameDesign = a.tt = b.tt And a.tb = b.tb And a.TBase = b.TBase And a.Base = b.Base And a.LToe = b.LToe And a.LHeel = b.LHeel _
        And a.ASst_DB = b.ASst_DB And a.ASst_Sp = b.ASst_Sp And a.AStoe_DB = b.AStoe_DB And a.AStoe_Sp = b.AStoe_Sp _
        And a.ASheel_DB = b.ASheel_DB And a.ASheel_Sp = b.ASheel_Sp And a.TotalCost = b.TotalCost And a.IsValid And b.IsValid
End Function

Private Function RunOne(method As String, trial As Long) As Design
    Dim d As Design, r As ProjectDetail, hist() As Double, n As Long, starts As Boolean
    If method = "BA" Then
        d = BisectionOptimization(1, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, currentMaterial, TrialNumber:=trial)
    Else
        d = HillClimbingOptimization(1, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, currentMaterial, TrialNumber:=trial)
    End If
    calls = calls + 1
    Verify method & " H" & H & " actual initial passes unchanged project checks", CheckProjectDesign(d, r) And d.IsValid
    Verify method & " initial occupies exactly one evaluation", EvaluationCount = 1 And RunBestEvaluation = 1
    Verify method & " No.0 contains the real cost only in Better", ReadAll(LastAcceptCSVPath) = "No.,Rejected,Passed,Passed and Better value" & vbCrLf & "0,,," & CsvPrice(d.TotalCost) & vbCrLf
    Verify method & " original catalogue geometry", d.tt = WP_tt(9) And d.tb = WP_tb(28) And d.TBase = WP_TBase(44) And d.Base = WP_Base(63) And d.LToe = WP_LToe(88)
    Verify method & " original catalogue reinforcement", d.ASst_DB = 102 And d.ASst_Sp = 111 And d.AStoe_DB = 102 And d.AStoe_Sp = 111 And d.ASheel_DB = 101 And d.ASheel_Sp = 112
    ReadAcceptCostHistory LastAcceptCSVPath, hist, n, starts
    Verify method & " graph genuinely starts at CSV No.0", starts And n = 1 And hist(1) = CDbl(CsvPrice(d.TotalCost))
    If method = "BA" Then Verify "BA in-memory first point matches real cost", CostHistory_BA(1) = d.TotalCost
    If method = "HCA" Then Verify "HCA in-memory first point matches real cost", modDataStructures.CostHistory(1) = d.TotalCost
    Verify method & " no feasibility recovery", Not RecoveryActive
    Print #f, "INITIAL H=" & H & "; method=" & method & "; cost=" & CsvPrice(d.TotalCost) & "; FS=" & CsvPrice(r.OT) & "," & CsvPrice(r.SL) & "," & CsvPrice(r.BC)
    RunOne = d
End Function

Private Sub RejectChangedInputs(method As String)
    Dim path As String, previous As String, d As Design, errorCode As Long
    path = RESULT_CSV_ROOT & "\accept-" & method & "-H5-320.csv"
    previous = ReadAll(path)
    qa = 0.1
    On Error Resume Next
    If method = "BA" Then
        d = BisectionOptimization(1, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, currentMaterial)
    Else
        d = HillClimbingOptimization(1, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, currentMaterial)
    End If
    errorCode = Err.Number
    Err.Clear
    On Error GoTo 0
    Verify method & " invalid current inputs stop before No.0", errorCode = vbObjectError + 2101 And EvaluationCount = 0
    Verify method & " failed initializer preserves previous CSV", ReadAll(path) = previous
    qa = 30
End Sub

Public Sub Main()
    Dim height As Integer, ba As Design, hca As Design, again As Design, mat As MaterialProperties
    On Error GoTo Failed
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
    If RESULT_CSV_ROOT <> App.Path & "\output" Then Err.Raise 5, , "Unsafe output path"
    InitializeArrays
    EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4
    phi = 30: mu = 0.6: qa = 30: cover = 0.075
    For height = 3 To 5
        H = height
        mat = GetSD40Material(240 + (height - 3) * 40, GetConcretePrice(240 + (height - 3) * 40), STEEL_PRICE_SD40)
        currentMaterial = mat
        ba = RunOne("BA", 1)
        hca = RunOne("HCA", 1)
        Verify "H" & H & " BA/HCA same full design and raw price", SameDesign(ba, hca)
        again = RunOne("BA", 7)
        Verify "BA repeated trial has identical initializer", SameDesign(ba, again)
        again = RunOne("HCA", 7)
        Verify "HCA repeated trial has identical initializer", SameDesign(hca, again)
    Next height
    RejectChangedInputs "BA"
    RejectChangedInputs "HCA"
    Print #f, "CHECKS=" & checks & "; FAILURES=" & failures
    Print #f, "INITIAL_ONLY_CALLS=" & calls & "; NEIGHBOR_OR_MIDPOINT_EVALUATIONS=0"
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Number & ":" & Err.Description
    Close #f
End Sub
