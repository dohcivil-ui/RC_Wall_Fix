Attribute VB_Name = "NeighborParity"
Option Explicit
Public Sub Main()
    Dim f As Integer, sample As Integer, height As Integer, k As Integer, checks As Long, failures As Long
    Dim state(1 To 11) As Integer, ba(1 To 11) As Integer, hca(1 To 11) As Integer
    Dim baNext As Single, hcaNext As Single, mat As MaterialProperties, d As Design, fc As Integer
    On Error GoTo Failed
    Call InitializeArrays
    Call EnableProjectChecks
    f = FreeFile: Open App.Path & "\parity.txt" For Output As #f
    For height = 3 To 5
        H = height
        For sample = 0 To 149
            state(1) = TT_MIN + sample Mod (TT_MAX - TT_MIN + 1)
            state(2) = TB_MIN + (sample * 3) Mod (tb_max - TB_MIN + 1)
            state(3) = TBASE_MIN + (sample * 7) Mod (TBase_max - TBASE_MIN + 1)
            state(4) = BASE_MIN + sample Mod (BASE_MAX - BASE_MIN + 1)
            state(5) = LTOE_MIN + sample Mod (LTOE_MAX - LTOE_MIN + 1)
            state(6) = DB_MIN + sample Mod 5: state(7) = SP_MIN + sample Mod 4
            state(8) = DB_MIN + (sample + 1) Mod 5: state(9) = SP_MIN + (sample + 2) Mod 4
            state(10) = DB_MIN + (sample + 2) Mod 5: state(11) = SP_MIN + (sample + 3) Mod 4
            Call modBA.BAParity(state, 12345 + CLng(sample), ba, baNext)
            Call modHillClimbing.HCAParity(state, 12345 + CLng(sample), hca, hcaNext)
            For k = 1 To 11
                checks = checks + 1
                If ba(k) <> hca(k) Then
                    failures = failures + 1
                    If failures <= 8 Then Print #f, "DIFF H=" & height & "; sample=" & sample & "; variable=" & k & "; BA=" & ba(k) & "; HCA=" & hca(k)
                End If
            Next k
            checks = checks + 1
            If baNext <> hcaNext Then failures = failures + 1
        Next sample
    Next height
    Print #f, "GENERATOR checks=" & checks & "; failures=" & failures
    If Command$ = "final" Then
        For height = 3 To 5
            fc = 120 + 40 * height
            mat = GetSD40Material(fc, GetConcretePrice(fc), 24)
            d = BisectionOptimization(64, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12345)
            Print #f, "RUN=" & RunFolder
            d = HillClimbingOptimization(64, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12345)
            Print #f, "RUN=" & RunFolder
        Next height
        d = HillClimbingOptimization(5000, 5, 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12365)
        Print #f, "HCA5000 cost=" & RunBestCost & "; bestEval=" & RunBestEvaluation
        Print #f, "RUN=" & RunFolder
        d = HillClimbingOptimization(125, 5, 1.2, 1.8, 2.4, 30, 0.6, 0.01, 0.075, mat, RandomSeed:=12345)
        Print #f, "NO_SOLUTION valid=" & RunBest.IsValid
        Print #f, "RUN=" & RunFolder
        If RunBest.IsValid Then Print #f, "FAIL: expected no solution"
    End If
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Description
    Close #f
End Sub
