Attribute VB_Name = "CsvPathRegression"
Option Explicit
Private logFile As Integer, checks As Long, failures As Long

Private Sub Verify(label As String, condition As Boolean)
    checks = checks + 1
    If Not condition Then failures = failures + 1
    Print #logFile, IIf(condition, "OK: ", "FAIL: ") & label
End Sub

Private Sub CheckMethod(method As String)
    Dim i As Long, displayed As Boolean, stem As String, firstPath As String
    If method = "BA" Then Form1.cmdBA.Value = True Else Form1.cmdRun.Value = True
    Verify method & " completed short trial", EvaluationCount = CLng(Form1.txtMaxIter.Text) And RunTrial = CLng(Form1.txtTrials.Text)
    Verify method & " uses requested root", Left$(RunFolder, Len(RESULT_CSV_ROOT) + 1) = RESULT_CSV_ROOT & "\"
    Verify method & " evaluations saved", Len(Dir$(RunFolder & "\evaluations.csv")) > 0
    Verify method & " summary saved below requested root", Left$(ProjectTrialSummary, Len(RESULT_CSV_ROOT) + 1) = RESULT_CSV_ROOT & "\" And Len(Dir$(ProjectTrialSummary)) > 0
    For i = 0 To Form1.lstResults.ListCount - 1
        If InStr(Form1.lstResults.List(i), "CSV and per-trial reports: " & RESULT_CSV_ROOT) > 0 Then displayed = True
    Next i
    Verify method & " UI displays actual root", displayed
    Verify method & " accept CSV saved automatically", Len(LastAcceptCSVPath) > 0 And Len(Dir$(LastAcceptCSVPath)) > 0
    Verify method & " accept is directly in root", Left$(LastAcceptCSVPath, Len(RESULT_CSV_ROOT & "\accept-" & method & "-H5-")) = RESULT_CSV_ROOT & "\accept-" & method & "-H5-"
    Verify method & " loop CSV saved automatically", Len(LastLoopCSVPath) > 0 And Len(Dir$(LastLoopCSVPath)) > 0
    Verify method & " loop is directly in root", Left$(LastLoopCSVPath, Len(RESULT_CSV_ROOT & "\loopPrice-" & method & "-H5")) = RESULT_CSV_ROOT & "\loopPrice-" & method & "-H5"
    Print #logFile, "ACCEPT=" & LastAcceptCSVPath
    Print #logFile, "LOOP=" & LastLoopCSVPath
    Print #logFile, "EVALUATIONS=" & RunFolder & "\evaluations.csv"
    firstPath = LastLoopCSVPath
    If method = "BA" Then SaveLoopPriceCSV_BA 5 Else SaveLoopPriceCSV 5
    Verify method & " repeated loop filename is unique", LastLoopCSVPath <> firstPath
    Verify method & " repeated loop CSV preserves first", Len(Dir$(firstPath)) > 0 And Len(Dir$(LastLoopCSVPath)) > 0
    Print #logFile, method & " run folder: " & RunFolder
    Print #logFile, method & " summary: " & ProjectTrialSummary
End Sub

Public Sub Main()
    Dim evidence As String, i As Long
    On Error GoTo Failed
    evidence = Trim$(Replace$(Command$, Chr$(34), ""))
    logFile = FreeFile
    Open evidence & "\native-csv-paths.txt" For Output As #logFile
    Load Form1: Form1.Show
    Verify "fixed output root independent of App.Path", ResultCsvRoot() = "C:\reserch 69\RC_Wall_Fix\result_csv" And App.Path <> "C:\reserch 69\RC_Wall_Fix"
    Verify "root exists or was created", Len(Dir$(ResultCsvRoot(), vbDirectory)) > 0
    RunFolder = ""
    Verify "export before first trial falls back to requested root", UniqueExportPath("pre-trial") = RESULT_CSV_ROOT & "\pre-trial.csv"
    Form1.txtH.Text = "5": Form1.txtH1.Text = "1.2"
    Form1.txtQa.Text = "30": Form1.txtCover.Text = "7.5"
    Form1.txtMaxIter.Text = "64": Form1.txtTrials.Text = "1": Form1.txtSeed.Text = "20260908"
    For i = 0 To Form1.cboConcreteStrength.ListCount - 1
        If Form1.cboConcreteStrength.List(i) = "320" Then Form1.cboConcreteStrength.ListIndex = i
    Next i
    CheckMethod "BA"
    CheckMethod "HCA"
    Form1.txtMaxIter.Text = "1": Form1.txtTrials.Text = "2": Form1.txtSeed.Text = "12345"
    CheckMethod "BA"
    CheckMethod "HCA"
    Unload Form1
    Print #logFile, "CSV_PATH checks=" & checks & "; failures=" & failures
    Close #logFile
    Exit Sub
Failed:
    Print #logFile, "FATAL " & Err.Number & ": " & Err.Description
    Close #logFile
    Unload Form1
End Sub
