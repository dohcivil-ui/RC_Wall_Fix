Attribute VB_Name = "BestTrialProbe"
Option Explicit
Private f As Integer, checks As Long, failures As Long

Private Sub Verify(label As String, ok As Boolean)
    checks = checks + 1
    If Not ok Then failures = failures + 1
    Print #f, IIf(ok, "PASS ", "FAIL ") & label
End Sub

Private Function ReadAll(path As String) As String
    Dim n As Integer
    n = FreeFile
    Open path For Binary Access Read As #n
    ReadAll = Space$(LOF(n))
    If LOF(n) > 0 Then Get #n, , ReadAll
    Close #n
End Function

Private Function ReferenceDesign(Optional moreSteel As Boolean = False) As Design
    Dim d As Design, ot As Double, sl As Double, bc As Double
    d.tt = WP_tt(9): d.tb = WP_tb(28): d.TBase = WP_TBase(44): d.Base = WP_Base(63): d.LToe = WP_LToe(88)
    d.LHeel = d.Base - d.tb - d.LToe
    d.ASst_DB = 102: d.ASst_Sp = 111: d.AStoe_DB = 102: d.AStoe_Sp = 111
    d.ASheel_DB = 101: d.ASheel_Sp = 112
    If moreSteel Then d.ASheel_Sp = 111
    d.IsValid = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, ot, sl, bc)
    If Not d.IsValid Then Err.Raise 5, , "Invalid drawing fixture: " & LastValidationReason
    d.TotalCost = CalculateCost(d)
    ReferenceDesign = d
End Function

Private Sub Replay(method As String, trial As Long, d As Design, bestLoop As Long)
    Dim i As Long, better As Boolean, price As String
    If method = "BA" Then InitCSVExport_BA Else InitCSVExport
    RunAlgorithm = method: RunTrial = trial: RunBest = d: RunBestCost = d.TotalCost: RunBestEvaluation = bestLoop + 1
    For i = 0 To 5
        better = i = bestLoop
        price = CsvPrice(d.TotalCost)
        If method = "BA" Then
            LogIteration_BA i, price, better, better
        Else
            LogIteration i, price, better, better
        End If
    Next i
    FinishSearch
    If method = "BA" Then LogLoopResult_BA d.TotalCost Else LogLoopResult d.TotalCost
End Sub

Private Sub CheckMethod(method As String)
    Dim selected As Design, higher As Design, trace As String, loops As String, picPath As String, beforePic As String
    selected = ReferenceDesign(): higher = ReferenceDesign(True)
    If method = "BA" Then InitLoopCounter_BA Else InitLoopCounter
    Replay method, 1, higher, 0
    Replay method, 2, selected, 4
    Replay method, 3, selected, 2
    trace = ReadAll(LastAcceptCSVPath)
    Replay method, 4, selected, 2
    Replay method, 5, higher, 1
    If method = "BA" Then SaveLoopPriceCSV_BA H Else SaveLoopPriceCSV H
    Verify method & " H" & H & " whole accept is best trial, not last", ReadAll(LastAcceptCSVPath) = trace
    loops = "No.,Loop,BestPrice" & vbCrLf & "1,0," & CsvPrice(higher.TotalCost) & vbCrLf & "2,4," & CsvPrice(selected.TotalCost) & vbCrLf
    loops = loops & "3,2," & CsvPrice(selected.TotalCost) & vbCrLf & "4,2," & CsvPrice(selected.TotalCost) & vbCrLf & "5,1," & CsvPrice(higher.TotalCost) & vbCrLf
    Verify method & " loopPrice keeps all trial rows", ReadAll(LastLoopCSVPath) = loops
    frmBestDesign.ShowBest selected, H, H1, method, 3, selected.TotalCost, cover, 5, 3
    picPath = frmBestDesign.SaveResultImage(H, method)
    Verify method & " BMP reports selected trial and zero-based Loop", InStr(frmBestDesign.Tag, "; trial=3;") > 0 And InStr(frmBestDesign.Tag, "; Loop=2;") > 0
    Verify method & " image saved with fixed method/H/fc filename", picPath = RESULT_CSV_ROOT & "\design-" & method & "-H" & H & "-" & currentMaterial.fc & ".bmp"
    beforePic = ReadAll(picPath)
    Verify method & " valid BMP", Left$(beforePic, 2) = "BM"
    If H = 3 Then FileCopy picPath, App.Path & "\review-" & method & "-selected-trial.bmp"
    If method = "BA" Then InitLoopCounter_BA Else InitLoopCounter
    Replay method, 1, selected, 0
    trace = ReadAll(LastAcceptCSVPath)
    Replay method, 2, higher, 0
    If method = "BA" Then SaveLoopPriceCSV_BA H Else SaveLoopPriceCSV H
    Verify method & " repeated run overwrites accept at same name", ReadAll(LastAcceptCSVPath) = trace
    loops = "No.,Loop,BestPrice" & vbCrLf & "1,0," & CsvPrice(selected.TotalCost) & vbCrLf & "2,0," & CsvPrice(higher.TotalCost) & vbCrLf
    Verify method & " repeated run replaces loopPrice without stale rows", ReadAll(LastLoopCSVPath) = loops
    frmBestDesign.ShowBest selected, H, H1, method, 1, selected.TotalCost, cover, 2, 1
    Verify method & " BMP overwrite uses same path", frmBestDesign.SaveResultImage(H, method) = picPath
    Verify method & " BMP updates best trial and Loop 0", InStr(frmBestDesign.Tag, "; trial=1;") > 0 And InStr(frmBestDesign.Tag, "; Loop=0;") > 0 And ReadAll(picPath) <> beforePic
    Unload frmBestDesign
End Sub

Public Sub Main()
    Dim height As Integer
    On Error GoTo Failed
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
    If RESULT_CSV_ROOT <> App.Path & "\output" Then Err.Raise 5, , "Unsafe result path"
    Load Form1
    InitializeArrays
    EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 30: cover = 0.075
    For height = 3 To 5
        H = height
        currentMaterial = GetSD40Material(240 + (height - 3) * 40, GetConcretePrice(240 + (height - 3) * 40), STEEL_PRICE_SD40)
        currentWSD = CalculateWSDParameters(currentMaterial.fy, currentMaterial.fc)
        CheckMethod "BA"
        CheckMethod "HCA"
    Next height
    Print #f, "CHECKS=" & checks & "; FAILURES=" & failures & "; OPTIMIZER_RUNS=0"
    Close #f
    Unload Form1
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Number & ":" & Err.Description
    Close #f
    Unload frmBestDesign
    Unload Form1
End Sub
