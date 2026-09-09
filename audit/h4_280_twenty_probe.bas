Attribute VB_Name = "TenTrialProbe"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private traceFile As Integer
Private failures As Long
Private trialCount As Long

Public Sub TraceCandidate(d As Design, entry As String, cost As Double)
    If traceFile = 0 Then Exit Sub
    Print #traceFile, RunAlgorithm & "," & RunTrial & "," & EvaluationCount & "," & entry & "," & CsvNumber(cost) & "," & CStr(d.IsValid) & "," & CsvNumber(RunBestCost) & "," & CsvNumber(d.tt) & "," & CsvNumber(d.tb) & "," & CsvNumber(d.TBase) & "," & CsvNumber(d.Base) & "," & CsvNumber(d.LToe) & "," & CsvNumber(CDbl(d.ASst_DB)) & "," & CsvNumber(CDbl(d.ASst_Sp)) & "," & CsvNumber(CDbl(d.AStoe_DB)) & "," & CsvNumber(CDbl(d.AStoe_Sp)) & "," & CsvNumber(CDbl(d.ASheel_DB)) & "," & CsvNumber(CDbl(d.ASheel_Sp)) & "," & RawDouble(cost) & "," & RawDouble(RunBestCost)
End Sub

Public Sub RecordTrial()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\trials.csv" For Append As #f
    Print #f, RunAlgorithm & "," & RunTrial & "," & EvaluationCount & "," & CsvNumber(RunBestCost) & "," & RunBestEvaluation & "," & CStr(RunBest.IsValid) & "," & CsvNumber(RunBest.FS_OT) & "," & CsvNumber(RunBest.FS_SL) & "," & CsvNumber(RunBest.FS_BC) & "," & RunRecoveryCount & "," & CsvNumber(RunBest.tb) & "," & CsvNumber(RunBest.TBase) & "," & CsvNumber(RunBest.Base) & "," & CsvNumber(H) & "," & CsvNumber(CDbl(currentMaterial.fc))
    Close #f
    trialCount = trialCount + 1
    If EvaluationCount <> 5000 Then failures = failures + 1
    If H <> 4 Or currentMaterial.fc <> 280 Then Err.Raise 5, , "Actual H/fc mismatch"
    If RunBest.IsValid Then
        If RunBest.FS_OT < FS_OT_MIN Or RunBest.FS_SL < FS_SL_MIN Or RunBest.FS_BC < FS_BC_MIN Then failures = failures + 1
    End If
End Sub

Private Sub RunMethod(method As String)
    Dim f As Integer
    traceFile = FreeFile
    Open App.Path & "\" & method & "-trace.csv" For Output As #traceFile
    Print #traceFile, "algorithm,trial,evaluation,entry,cost,valid,best_cost,tt,tb,TBase,Base,LToe,ASst_DB,ASst_Sp,AStoe_DB,AStoe_Sp,ASheel_DB,ASheel_Sp,exact_hex,best_hex"
    If method = "BA" Then
        Form1.cmdBA.Value = True
    Else
        Form1.cmdRun.Value = True
    End If
    Close #traceFile
    traceFile = 0
    If RunTrial <> 20 Or RunFolder <> "" Then failures = failures + 1
    f = FreeFile
    Open App.Path & "\" & method & "-screen.txt" For Output As #f
    Print #f, Form1.txtResults.Text
    Close #f
End Sub

Public Sub Main()
    Dim f As Integer
    On Error GoTo Failed
    If RESULT_CSV_ROOT <> App.Path & "\output" Then Err.Raise 5, , "Unsafe output root"
    f = FreeFile
    Open App.Path & "\trials.csv" For Output As #f
    Print #f, "algorithm,trial,evaluations,best_cost,best_evaluation,valid,FS_OT,FS_SL,FS_BC,recoveries,tb,TBase,Base,wall_height,fc"
    Close #f
    Load Form1
    Form1.txtH.Text = "4": Form1.txtH1.Text = "1.2"
    Form1.txtGammaSoil.Text = "1.8": Form1.txtGammaCon.Text = "2.4"
    Form1.txtPhi.Text = "30": Form1.txtMu.Text = "0.6"
    Form1.txtQa.Text = "30": Form1.txtCover.Text = "7.5"
    Form1.txtMaxIter.Text = "5000": Form1.txtTrials.Text = "20"
    Form1.cboConcreteStrength.ListIndex = 3
    If Val(Form1.cboConcreteStrength.Text) <> 280 Then Err.Raise 5, , "Concrete-strength selector mismatch"
    RunMethod "BA"
    RunMethod "HCA"
    If trialCount <> 40 Then failures = failures + 1
    Unload frmBestDesign
    Unload Form1
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
    Print #f, "TRIALS=" & trialCount
    Print #f, "FAILURES=" & failures
    Close #f
    Exit Sub
Failed:
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
    Print #f, "FATAL=" & Err.Number & ": " & Err.Description
    Close #f
    End
End Sub

Private Function RawDouble(value As Double) As String
    Dim bytes(0 To 7) As Byte, i As Integer, result As String
    CopyMemory bytes(0), value, 8
    For i = 0 To 7: result = result & Right$("0" & Hex$(bytes(i)), 2): Next i
    RawDouble = result
End Function
