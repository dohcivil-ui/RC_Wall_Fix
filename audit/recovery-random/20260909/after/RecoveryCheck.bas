Attribute VB_Name = "RecoveryCheck"
Option Explicit
Private logFile As Integer
Public Sub RecordRecovery(tb As Integer, thick As Integer, base As Integer, nextTb As Integer, nextThick As Integer, nextBase As Integer)
    Print #logFile, "RECOVERY=" & tb & "," & thick & "," & base & "," & nextTb & "," & nextThick & "," & nextBase
End Sub
Public Sub Main()
    Dim height As Integer, mat As MaterialProperties, d As Design
    On Error GoTo Failed
    Call InitializeArrays
    Call EnableProjectChecks
    logFile = FreeFile
    Open App.Path & "\recovery.txt" For Output As #logFile
    For height = 3 To 5
        mat = GetSD40Material(120 + 40 * height, GetConcretePrice(120 + 40 * height), 24)
        d = BisectionOptimization(125, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 0.01, 0.075, mat, RandomSeed:=12345)
        Print #logFile, "NO_SOLUTION=" & RunBest.IsValid
        Print #logFile, "RUN=" & RunFolder
        d = HillClimbingOptimization(125, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 0.01, 0.075, mat, RandomSeed:=12345)
        Print #logFile, "RUN=" & RunFolder
        d = BisectionOptimization(500, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12345)
        Print #logFile, "RUN=" & RunFolder
    Next height
    Close #logFile
    Exit Sub
Failed:
    Print #logFile, "FATAL=" & Err.Description
    Close #logFile
End Sub
