Attribute VB_Name = "Holdout"
Option Explicit
Public Sub Main()
    Dim height As Integer, sample As Integer, seed As Long, method As Integer
    Dim mat As MaterialProperties, d As Design, f As Integer
    On Error GoTo Failed
    Call InitializeArrays
    Call EnableProjectChecks
    f = FreeFile
    Open App.Path & "\holdout.txt" For Output As #f
    For height = 3 To 5
        mat = GetSD40Material(120 + 40 * height, GetConcretePrice(120 + 40 * height), 24)
        For sample = 0 To 0
            If height < 5 And sample > 0 Then Exit For
            Select Case sample
                Case 0: seed = 24681
                Case 1: seed = 12360
                Case 2: seed = 12374
                Case 3: seed = 24680
            End Select
            For method = 0 To 1
                If method = 0 Then
                    d = HillClimbingOptimization(5000, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=seed)
                Else
                    d = BisectionOptimization(5000, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=seed)
                End If
                Print #f, "H=" & height & "; method=" & RunAlgorithm & "; seed=" & seed & "; cost=" & RunBestCost & "; evaluation=" & RunBestEvaluation
                Print #f, "RUN=" & RunFolder
            Next method
        Next sample
    Next height
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Description
    Close #f
End Sub
