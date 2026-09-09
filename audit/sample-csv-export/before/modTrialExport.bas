Attribute VB_Name = "modTrialExport"
Option Explicit

Public LastTrialAcceptCSVPath As String
Private RequestedTrials As Long
Private WrittenTrials As Long
Private BatchAlgorithm As String
Private BatchInputs As String
Private Const ACCEPT_HEADER As String = "No.,Rejected,Passed,Passed and Better value"

Public Sub BeginTrialAcceptExport(ByVal trials As Long)
    If trials < 1 Then Err.Raise 5, , "Trial count must be positive"
    RequestedTrials = trials: WrittenTrials = 0
    LastTrialAcceptCSVPath = "": BatchAlgorithm = "": BatchInputs = ""
End Sub

Private Function TrialInputValues() As String
    TrialInputValues = CsvNumber(H) & "," & CStr(currentMaterial.fc) & "," & CStr(EvaluationBudget) & "," & _
        CsvNumber(H1) & "," & CsvNumber(gamma_soil) & "," & CsvNumber(gamma_concrete) & "," & _
        CsvNumber(phi) & "," & CsvNumber(mu) & "," & CsvNumber(qa) & "," & CsvNumber(cover) & "," & _
        CsvNumber(currentMaterial.concretePrice) & "," & CsvNumber(currentMaterial.SteelPrice) & "," & CStr(currentMaterial.fy)
End Function

Public Sub AppendTrialAcceptExport()
    Dim f As Integer, raw As String, rows() As String, outputRows() As String
    Dim rowCount As Long, i As Long, fields() As String, suffix As String, inputs As String, bestPrice As String
    Dim number As Long, description As String, path As String
    On Error GoTo Failed
    If RequestedTrials < 1 Or RunTrial <> WrittenTrials + 1 Or RunTrial > RequestedTrials Then
        Err.Raise 5, , "Unexpected trial sequence in combined accept export"
    End If
    If RunAlgorithm <> "BA" And RunAlgorithm <> "HCA" Then Err.Raise 5, , "Unknown export algorithm"
    If EvaluationCount <> EvaluationBudget Then Err.Raise 5, , "Incomplete trial export"
    If Len(LastAcceptCSVPath) = 0 Then Err.Raise 5, , "Missing trial accept CSV"
    inputs = TrialInputValues()
    If WrittenTrials > 0 Then
        If RunAlgorithm <> BatchAlgorithm Or inputs <> BatchInputs Then Err.Raise 5, , "Trial inputs changed within export batch"
    End If
    If Len(Dir$(LastAcceptCSVPath)) = 0 Then Err.Raise 53, , "Trial accept CSV not found"
    f = FreeFile
    Open LastAcceptCSVPath For Binary Access Read As #f
    raw = Space$(LOF(f))
    If Len(raw) > 0 Then Get #f, , raw
    Close #f: f = 0
    rows = Split(raw, vbCrLf)
    If rows(0) <> ACCEPT_HEADER Then Err.Raise 5, , "Unexpected accept CSV header"
    rowCount = UBound(rows)
    If rowCount > 0 Then
        If Len(rows(rowCount)) = 0 Then rowCount = rowCount - 1
    End If
    If rowCount <> EvaluationCount Then Err.Raise 5, , "Accept CSV evaluation count mismatch"
    If RunBest.IsValid Then
        If RunBestEvaluation < 1 Or RunBestEvaluation > EvaluationCount Then Err.Raise 5, , "Invalid best evaluation"
        bestPrice = CsvNumber(RunBestCost)
    Else
        If RunBestEvaluation <> 0 Then Err.Raise 5, , "Invalid no-solution best evaluation"
    End If
    suffix = "," & CStr(RunTrial) & "," & RunAlgorithm & "," & inputs & "," & _
        CStr(RunBestEvaluation) & "," & bestPrice & "," & RunStatus & "," & CStr(RequestedTrials)
    ReDim outputRows(0 To rowCount - 1)
    For i = 0 To rowCount - 1
        fields = Split(rows(i + 1), ",")
        If UBound(fields) <> 3 Then Err.Raise 5, , "Invalid accept CSV columns"
        If fields(0) <> CStr(i) Then Err.Raise 5, , "Invalid accept CSV evaluation sequence"
        outputRows(i) = rows(i + 1) & suffix
    Next i
    If WrittenTrials = 0 Then
        path = UniqueExportPath("acceptRuns-" & RunAlgorithm & "-H" & Replace$(CStr(H), ",", ".") & _
            "-fc" & CStr(currentMaterial.fc) & "-" & Format$(Now, "yyyymmdd-hhnnss"), True)
        f = FreeFile
        Open path For Output As #f
        Print #f, ACCEPT_HEADER & ",Trial,Algorithm,H,fc,EvaluationBudget,H1,GammaSoil,GammaConcrete,Phi,Mu,Qa,Cover,ConcretePrice,SteelPrice,fy,BestEvaluation,BestPriceExact,Status,TrialsRequested"
    Else
        path = LastTrialAcceptCSVPath
        f = FreeFile
        Open path For Append As #f
    End If
    Print #f, Join(outputRows, vbCrLf)
    Close #f: f = 0
    LastTrialAcceptCSVPath = path
    BatchAlgorithm = RunAlgorithm: BatchInputs = inputs
    WrittenTrials = WrittenTrials + 1
    Exit Sub
Failed:
    number = Err.Number: description = Err.description
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise number, "AppendTrialAcceptExport", description
End Sub
