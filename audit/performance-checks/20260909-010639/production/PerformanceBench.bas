Attribute VB_Name = "PerformanceBench"
Option Explicit
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef count As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
Public ProfileCalls(0 To 8) As Long
Public ProfileSeconds(0 To 8) As Double
Public Function ClockSeconds() As Double
    Dim counter As Currency, frequency As Currency
    QueryPerformanceCounter counter: QueryPerformanceFrequency frequency
    ClockSeconds = counter / frequency
End Function
Public Sub Main()
    Dim budget As Long
    Dim d As Design, mat As MaterialProperties, f As Integer, i As Integer, method As Integer, started As Double
    On Error GoTo Failed
    f = FreeFile: Open App.Path & "\timing.txt" For Output As #f
    InitializeArrays: EnableProjectChecks
    mat = GetSD40Material(320, GetConcretePrice(320), 24)
    For method = 0 To 3
        For i = 0 To 8: ProfileCalls(i) = 0: ProfileSeconds(i) = 0: Next i
        budget = 1024: If method >= 2 Then budget = 5000
        started = ClockSeconds()
        If method Mod 2 = 0 Then
            d = BisectionOptimization(budget, 5, 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12345, TrialNumber:=1)
        Else
            d = HillClimbingOptimization(budget, 5, 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12345, TrialNumber:=1)
        End If
        Print #f, "method=" & RunAlgorithm & "; evaluations=" & EvaluationCount & "; elapsed=" & (ClockSeconds() - started)
        For i = 0 To 8: Print #f, "slot=" & i & "; calls=" & ProfileCalls(i) & "; seconds=" & ProfileSeconds(i): Next i
        Print #f, "run=" & RunFolder
        If ProfileCalls(0) > EvaluationCount + 10 Then Print #f, "DUPLICATE_VALIDATION_DETECTED"
    Next method
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL " & Err.Number & ": " & Err.Description
    Close #f
End Sub

Public Sub ProfileEvents()
    Dim started As Double
    started = ClockSeconds()
    DoEvents
    ProfileSeconds(6) = ProfileSeconds(6) + ClockSeconds() - started
End Sub
