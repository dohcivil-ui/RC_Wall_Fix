Attribute VB_Name = "CsvPathRegression"
Option Explicit
Private logFile As Integer, checks As Long, failures As Long
Private evidence As String, sessionNumber As Long

Private Sub Verify(label As String, condition As Boolean)
    checks = checks + 1
    If Not condition Then failures = failures + 1
    Print #logFile, IIf(condition, "OK: ", "FAIL: ") & label
End Sub

Private Sub CheckMethod(method As String)
    Dim i As Long, displayed As Boolean, firstPath As String, backupCount As Long
    Dim sample As Design, detail As ProjectDetail, reason As String, quantityOK As Boolean
    sessionNumber = sessionNumber + 1
    If method = "BA" Then Form1.cmdBA.Value = True Else Form1.cmdRun.Value = True
    If RunBest.IsValid Then
        sample = RunBest
        reason = LastValidationReason
        quantityOK = ProjectQuantityCost(sample, detail)
        Verify method & " quantity price equals accepted price", quantityOK And Abs(detail.Cost - RunBestCost) < 0.000001
        sample.LToe = cover - 0.01: sample.LHeel = sample.Base - sample.LToe - sample.tb
        Verify method & " negative main length has no invented price", Not ProjectQuantityCost(sample, detail)
        sample = RunBest: sample.ASst_DB = 0
        Verify method & " invalid bar index has no invented price", Not ProjectQuantityCost(sample, detail)
        Verify method & " quantity logging preserves validation reason", LastValidationReason = reason
    End If
    Verify method & " completed short trial", EvaluationCount = CLng(Form1.txtMaxIter.Text) And RunTrial = CLng(Form1.txtTrials.Text)
    Verify method & " uses requested root", Left$(RunFolder, Len(RESULT_CSV_ROOT) + 1) = RESULT_CSV_ROOT & "\"
    Verify method & " evaluations saved", Len(Dir$(RunFolder & "\evaluations.csv")) > 0
    Verify method & " summary saved below requested root", Left$(ProjectTrialSummary, Len(RESULT_CSV_ROOT) + 1) = RESULT_CSV_ROOT & "\" And Len(Dir$(ProjectTrialSummary)) > 0
    For i = 0 To Form1.lstResults.ListCount - 1
        If InStr(Form1.lstResults.List(i), "CSV and per-trial reports: " & RESULT_CSV_ROOT) > 0 Then displayed = True
    Next i
    Verify method & " UI displays actual root", displayed
    Verify method & " accept CSV saved automatically", Len(LastAcceptCSVPath) > 0 And Len(Dir$(LastAcceptCSVPath)) > 0
    Verify method & " exact accept filename", LastAcceptCSVPath = RESULT_CSV_ROOT & "\accept-" & method & "-H5.csv"
    Verify method & " loop CSV saved automatically", Len(LastLoopCSVPath) > 0 And Len(Dir$(LastLoopCSVPath)) > 0
    Verify method & " exact loop filename", LastLoopCSVPath = RESULT_CSV_ROOT & "\loopPrice-" & method & "-H5.csv"
    Print #logFile, "ACCEPT=" & LastAcceptCSVPath
    Print #logFile, "LOOP=" & LastLoopCSVPath
    Print #logFile, "EVALUATIONS=" & RunFolder & "\evaluations.csv"
    FileCopy LastAcceptCSVPath, evidence & "\session" & sessionNumber & "-accept.csv"
    FileCopy LastLoopCSVPath, evidence & "\session" & sessionNumber & "-loop.csv"
    Print #logFile, "ACCEPT_SNAPSHOT=" & evidence & "\session" & sessionNumber & "-accept.csv"
    Print #logFile, "LOOP_SNAPSHOT=" & evidence & "\session" & sessionNumber & "-loop.csv"
    backupCount = CountBackups(method)
    firstPath = LastLoopCSVPath
    If method = "BA" Then SaveLoopPriceCSV_BA 5 Else SaveLoopPriceCSV 5
    Verify method & " repeated loop keeps exact filename", LastLoopCSVPath = firstPath
    Verify method & " repeated loop archives previous file", CountBackups(method) = backupCount + 1
    Print #logFile, method & " run folder: " & RunFolder
    Print #logFile, method & " summary: " & ProjectTrialSummary
End Sub

Public Sub Main()
    Dim i As Long
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
    Form1.txtQa.Text = "0.01"
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

Private Function CountBackups(method As String) As Long
    Dim name As String
    If Dir$(RESULT_CSV_ROOT & "\archive", vbDirectory) = "" Then Exit Function
    name = Dir$(RESULT_CSV_ROOT & "\archive\loopPrice-" & method & "-H5-*.csv")
    Do While Len(name) > 0
        CountBackups = CountBackups + 1
        name = Dir$()
    Loop
End Function
