Attribute VB_Name = "SampleCsvProbe"
Option Explicit
Private logFile As Integer
Private checks As Long, failures As Long
Private Const SAMPLE_ROOT As String = "D:\rc-rt-optimize-v2\vb6_samples"

Private Sub Verify(ByVal label As String, ByVal ok As Boolean)
    checks = checks + 1
    Print #logFile, IIf(ok, "PASS ", "FAIL ") & label
    If Not ok Then failures = failures + 1
End Sub

Private Function ReadBytes(ByVal path As String) As String
    Dim f As Integer
    f = FreeFile
    Open path For Binary Access Read As #f
    ReadBytes = Space$(LOF(f))
    If LOF(f) > 0 Then Get #f, , ReadBytes
    Close #f
End Function

Private Sub ReplayAccept(ByVal method As String)
    Dim f As Integer, line As String, fields() As String
    Dim rows As Long, price As String, valid As Boolean, better As Boolean
    Dim source As String, expected As String
    source = SAMPLE_ROOT & "\accept-" & method & "-H3-240.csv"
    If method = "BA" Then InitCSVExport_BA Else InitCSVExport
    f = FreeFile
    Open source For Input As #f
    Line Input #f, line
    Verify method & " source accept header", line = "No.,Rejected,Passed,Passed and Better value"
    Do While Not EOF(f)
        Line Input #f, line
        fields = Split(line, ",")
        If UBound(fields) <> 3 Then Err.Raise 5, , "Bad accept fixture"
        valid = Len(fields(1)) = 0
        better = Len(fields(3)) > 0
        If Not valid Then
            price = fields(1)
        ElseIf better Then
            price = fields(3)
        Else
            price = fields(2)
        End If
        If method = "BA" Then
            LogIteration_BA CLng(fields(0)), price, valid, better
        Else
            LogIteration CLng(fields(0)), price, valid, better
        End If
        rows = rows + 1
    Loop
    Close #f
    If method = "BA" Then SaveAcceptCSV_BA H Else SaveAcceptCSV H
    expected = RESULT_CSV_ROOT & "\accept-" & method & "-H3-240.csv"
    Verify method & " exact accept filename", LastAcceptCSVPath = expected
    Verify method & " complete accept sample bytes reproduced", ReadBytes(expected) = ReadBytes(source)
    Verify method & " replayed sample row count only", rows = 5001
End Sub

Private Sub ReplayLoopPrice(ByVal method As String)
    Dim f As Integer, line As String, fields() As String, rows As Long
    Dim source As String, expected As String, savedAccept As String
    source = SAMPLE_ROOT & "\loopPrice-" & method & "-H3-240.csv"
    savedAccept = ReadBytes(LastAcceptCSVPath)
    If method = "BA" Then InitLoopCounter_BA Else InitLoopCounter
    f = FreeFile
    Open source For Input As #f
    Line Input #f, line
    Verify method & " source loopPrice header", line = "No.,Loop,BestPrice"
    Do While Not EOF(f)
        Line Input #f, line
        fields = Split(line, ",")
        If UBound(fields) <> 2 Then Err.Raise 5, , "Bad loopPrice fixture"
        RunBest.IsValid = True
        If method = "BA" Then
            InitCSVExport_BA
            LogIteration_BA CLng(fields(1)), fields(2), True, True
            LogLoopResult_BA CDbl(fields(2))
        Else
            InitCSVExport
            LogIteration CLng(fields(1)), fields(2), True, True
            LogLoopResult CDbl(fields(2))
        End If
        rows = rows + 1
    Loop
    Close #f
    If method = "BA" Then SaveLoopPriceCSV_BA H Else SaveLoopPriceCSV H
    expected = RESULT_CSV_ROOT & "\loopPrice-" & method & "-H3-240.csv"
    Verify method & " exact loopPrice filename", LastLoopCSVPath = expected
    Verify method & " complete loopPrice sample bytes reproduced", ReadBytes(expected) = ReadBytes(source)
    Verify method & " replayed all sample summary rows", rows = 30
    Verify method & " summary save preserves last accept", ReadBytes(LastAcceptCSVPath) = savedAccept
End Sub

Private Sub CheckOtherScenario(ByVal method As String)
    Dim expected As String
    H = 4: currentMaterial.fc = 280
    RunBest.IsValid = False
    If method = "BA" Then
        InitCSVExport_BA
        InitLoopCounter_BA
        LogIteration_BA 0, "123.45", False, False
        LogLoopResult_BA NO_SOLUTION_COST
        SaveAcceptCSV_BA H
        SaveLoopPriceCSV_BA H
    Else
        InitCSVExport
        InitLoopCounter
        LogIteration 0, "123.45", False, False
        LogLoopResult NO_SOLUTION_COST
        SaveAcceptCSV H
        SaveLoopPriceCSV H
    End If
    Verify method & " names use actual H and fc", LastAcceptCSVPath = RESULT_CSV_ROOT & "\accept-" & method & "-H4-280.csv" And LastLoopCSVPath = RESULT_CSV_ROOT & "\loopPrice-" & method & "-H4-280.csv"
    expected = "No.,Loop,BestPrice" & vbCrLf & "1,0," & vbCrLf
    Verify method & " no-solution summary retained with blank price", ReadBytes(LastLoopCSVPath) = expected
    expected = "No.,Rejected,Passed,Passed and Better value" & vbCrLf & "0,123.45,," & vbCrLf
    Verify method & " no-solution rejection column retained", ReadBytes(LastAcceptCSVPath) = expected
End Sub

Public Sub Main()
    On Error GoTo Failed
    logFile = FreeFile
    Open App.Path & "\runtime.txt" For Output As #logFile
    If RESULT_CSV_ROOT <> App.Path & "\output" Then Err.Raise 5, , "Unsafe probe output path"
    Verify "isolated output path", True
    H = 3: currentMaterial.fc = 240
    ReplayAccept "BA"
    ReplayLoopPrice "BA"
    ReplayAccept "HCA"
    ReplayLoopPrice "HCA"
    Verify "no redundant last-trial accept archive", Len(Dir$(RESULT_CSV_ROOT & "\archive", vbDirectory)) = 0
    Verify "no combined acceptRuns files", Len(Dir$(RESULT_CSV_ROOT & "\acceptRuns-*.csv")) = 0
    CheckOtherScenario "BA"
    CheckOtherScenario "HCA"
    Print #logFile, "CHECKS=" & checks
    Print #logFile, "FAILURES=" & failures
    Print #logFile, "OPTIMIZER_RUNS=0"
    Print #logFile, "MODE=CSV_FIXTURE_REPLAY_ONLY"
    Close #logFile
    Exit Sub
Failed:
    Print #logFile, "FATAL=" & Err.Number & ": " & Err.Description
    Close #logFile
End Sub
