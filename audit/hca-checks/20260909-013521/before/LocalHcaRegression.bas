Attribute VB_Name = "LocalHcaRegression"
Option Explicit
Public Sub Main()
    Dim d As Design, mat As MaterialProperties, f As Integer
    On Error GoTo Failed
    Call InitializeArrays
    Call EnableProjectChecks
    mat = GetSD40Material(320, GetConcretePrice(320), 24)
    d = HillClimbingOptimization(125, 5, 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, True, 1, 24, 40, 63, 87, 101, 110, 100, 110, 101, 110, 12365)
    f = FreeFile: Open App.Path & "\local-regression.txt" For Output As #f
    Print #f, "StartCost=8573.88408; EndCost=" & RunBestCost & "; Evaluations=" & EvaluationCount
    If RunBestCost < 8573.88408 Then
        Print #f, "PASS: cheaper admissible main-bar neighbor found inside the 125-evaluation budget"
    Else
        Print #f, "FAIL: cheaper admissible main-bar neighbor missed inside the 125-evaluation budget"
    End If
    Close #f
    Exit Sub
Failed:
    f = FreeFile: Open App.Path & "\local-regression.txt" For Output As #f
    Print #f, "FATAL=" & Err.Description
    Close #f
End Sub
