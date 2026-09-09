Attribute VB_Name = "ExportProbe"
Option Explicit
Private logFile As Integer, failures As Long

Private Sub Verify(label As String, ok As Boolean)
    Print #logFile, IIf(ok, "OK: ", "FAIL: ") & label
    If Not ok Then failures = failures + 1
End Sub

Private Function ReadText(path As String) As String
    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    ReadText = Space$(LOF(f))
    Get #f, , ReadText
    Close #f
End Function

Private Sub CheckMethod(method As String)
    Dim trial As Long, price As String, expected As String
    Dim candidate As Design, cost As Double, valid As Boolean
    If method = "BA" Then InitLoopCounter_BA Else InitLoopCounter
    ProjectTrialSummary = ""
    For trial = 1 To 2
        BeginSearch 3, 12344 + trial, method, trial
        Verify method & " no run folder", RunFolder = ""
        valid = EvaluateCandidate(candidate, "export-probe", cost)
        Verify method & " invalid candidate and evaluation count", Not valid And EvaluationCount = 1
        price = CStr(100 + trial) & ".00"
        If method = "BA" Then
            LogIteration_BA 1, price, True, True
            LogIteration_BA 2, "200.00", True, False
        Else
            LogIteration 1, price, True, True
            LogIteration 2, "200.00", True, False
        End If
        FinishSearch
        RecordProjectTrial
        Verify method & " no implicit summary", ProjectTrialSummary = ""
        Verify method & " fixed accept name", LastAcceptCSVPath = RESULT_CSV_ROOT & "\accept-" & method & "-H5.csv"
        expected = "No.,Rejected,Passed,Passed and Better value" & vbCrLf & "0,,," & vbCrLf & "1,,," & price & vbCrLf & "2,,200.00," & vbCrLf
        Verify method & " CSV categories and zero-based indices", ReadText(LastAcceptCSVPath) = expected
        If method = "BA" Then LogLoopResult_BA NO_SOLUTION_COST Else LogLoopResult NO_SOLUTION_COST
    Next trial
    If method = "BA" Then SaveLoopPriceCSV_BA 5 Else SaveLoopPriceCSV 5
    Verify method & " fixed loop name", LastLoopCSVPath = RESULT_CSV_ROOT & "\loopPrice-" & method & "-H5.csv"
    Verify method & " loop rows retained", ReadText(LastLoopCSVPath) = "No.,Loop,BestPrice" & vbCrLf & "1,1," & vbCrLf & "2,1," & vbCrLf
End Sub

Public Sub Main()
    On Error GoTo Failed
    logFile = FreeFile
    Open App.Path & "\runtime.txt" For Output As #logFile
    Verify "isolated output root", RESULT_CSV_ROOT = App.Path & "\output"
    If RESULT_CSV_ROOT <> App.Path & "\output" Then Err.Raise 5, , "Unsafe output root"
    H = 5
    ProjectChecksEnabled = True
    CheckMethod "BA"
    CheckMethod "HCA"
    Print #logFile, "FAILURES=" & failures
    Close #logFile
    Exit Sub
Failed:
    Print #logFile, "FATAL=" & Err.Number & ": " & Err.Description
    Close #logFile
End Sub
