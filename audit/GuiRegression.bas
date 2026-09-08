Attribute VB_Name = "GuiRegression"
Option Explicit

Public Sub Main()
    Dim f As Integer, i As Integer, failures As Long, s As String
    On Error GoTo Failed
    f = FreeFile
    Open App.Path & "\gui-regression.txt" For Output As #f
    Load Form1
    ProjectChecksEnabled = False
    Form1.Show
    Print #f, "Project passive after Form_Load=" & PassiveFactor
    If PassiveFactor <> 1# Then failures = failures + 1
    DoEvents
    Print #f, "Actual compiled VB6 Form1 loaded and shown."
    Form1.txtH.Text = "5": Form1.txtH1.Text = "1.2"
    Form1.txtMu.Text = "0.6": Form1.txtGammaSoil.Text = "1.8"
    Form1.txtGammaCon.Text = "2.4": Form1.txtPhi.Text = "30"
    Form1.txtQa.Text = "20": Form1.txtCover.Text = "7.5"
    Form1.txtMaxIter.Text = "4": Form1.txtTrials.Text = "2"
    Form1.txtSeed.Text = "12345"
    For i = 0 To Form1.cboConcreteStrength.ListCount - 1
        If Form1.cboConcreteStrength.List(i) = "320" Then Form1.cboConcreteStrength.ListIndex = i
    Next i
    WSDReviewed = False
    ' CommandButton.Value invokes the real Click handler, including validation,
    ' trial aggregation, reporting, no-solution graph handling, and control state.
    Form1.cmdBA.Value = True
    s = ResultText()
    Print #f, "Project passive after search=" & PassiveFactor
    If PassiveFactor <> 1# Then failures = failures + 1
    Print #f, "BA actual button: evaluations(last trial)=" & EvaluationCount & "; trial=" & RunTrial
    Print #f, s
    If InStr(s, "NO_SOLUTION") = 0 Or EvaluationCount <> 4 Or RunTrial <> 2 Then failures = failures + 1
    If Not Form1.cmdBA.Enabled Or Not Form1.cmdRun.Enabled Then failures = failures + 1
    Form1.cmdRun.Value = True
    s = ResultText()
    Print #f, "Project passive after search=" & PassiveFactor
    If PassiveFactor <> 1# Then failures = failures + 1
    Print #f, "HCA actual button: evaluations(last trial)=" & EvaluationCount & "; trial=" & RunTrial
    Print #f, s
    If InStr(s, "NO_SOLUTION") = 0 Or EvaluationCount <> 4 Or RunTrial <> 2 Then failures = failures + 1
    If Not Form1.cmdBA.Enabled Or Not Form1.cmdRun.Enabled Then failures = failures + 1
    Form1.cmdCompare.Value = True
    s = ResultText()
    Print #f, "Compare no-solution sessions: " & s
    If InStr(s, "NO_SOLUTION") = 0 Or InStr(s, "999,999,999") > 0 Then failures = failures + 1
    Call CheckCompare(f, failures, 100, 100, 200, 1, "HCA has the lower")
    Call CheckCompare(f, failures, 200, 1, 100, 100, "BA has the lower")
    Call CheckCompare(f, failures, 100, 1, 100, 2, "HCA first reached")
    Call CheckCompare(f, failures, 100, 2, 100, 1, "BA first reached")
    Call CheckCompare(f, failures, 100, 1, 100, 1, "equal cost and first-best")
    Call CheckCompare(f, failures, NO_SOLUTION_COST, 0, 100, 1, "NO_SOLUTION")
    Call CheckCompare(f, failures, 100, 1, NO_SOLUTION_COST, 0, "NO_SOLUTION")
    WSDReviewed = True: WSDSource = "SYNTHETIC GUI TEST FIXTURE ONLY - not EIT"
    AllowableShear = 8: MinStemRatio = 0.0015: MinBaseRatio = 0.0015
    Form1.txtMaxIter.Text = "24": Form1.txtTrials.Text = "1"
    Form1.cmdBA.Value = True
    s = ResultText()
    Print #f, "Project passive after search=" & PassiveFactor
    If PassiveFactor <> 1# Then failures = failures + 1
    Print #f, "BA actual button with SYNTHETIC criteria: " & RunStatus
    Print #f, s
    If InStr(s, "VERIFIED_CONFIGURED_CHECKS_ONLY") = 0 Then failures = failures + 1
    If InStr(s, "fc_actual=") = 0 Or InStr(s, "Stem height H-TBase") = 0 Then failures = failures + 1
    Form1.txtSeed.Text = "12346"
    Form1.cmdRun.Value = True
    s = ResultText()
    Print #f, "Project passive after search=" & PassiveFactor
    If PassiveFactor <> 1# Then failures = failures + 1
    Print #f, "HCA actual button with SYNTHETIC criteria: " & RunStatus
    Print #f, s
    If InStr(s, "VERIFIED_CONFIGURED_CHECKS_ONLY") = 0 Then failures = failures + 1
    If InStr(s, "fc_actual=") = 0 Then failures = failures + 1
    Form1.cmdCompare.Value = True
    s = ResultText()
    Print #f, "Compare mismatched seeds: " & s
    If InStr(s, "differ; no ranking") = 0 Then failures = failures + 1
    Print #f, "GUI failures=" & failures
    Unload Form1
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL " & Err.Number & ": " & Err.Description
    Close #f
    Unload Form1
End Sub

Private Function ResultText() As String
    Dim i As Integer
    For i = 0 To Form1.lstResults.ListCount - 1
        ResultText = ResultText & Form1.lstResults.List(i) & vbCrLf
    Next i
End Function

Private Sub CheckCompare(f As Integer, ByRef failures As Long, hc As Double, he As Long, bc As Double, be As Long, expected As String)
    Dim actual As String
    actual = Form1.CompareObservedResults(hc, he, bc, be)
    Print #f, "Compare fixture: " & actual
    If InStr(actual, expected) = 0 Then failures = failures + 1
End Sub
