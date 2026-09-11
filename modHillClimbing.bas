Attribute VB_Name = "modHillClimbing"
'================================================================================
' Module: modHillClimbing.bas
' Project: RC_RT_HCA v2.8 - Cantilever Retaining Wall Optimization
' Purpose: Hill Climbing Algorithm (HCA) - ???????? functions ??? modShared
' Version: 5.1 - Fixed CSV Export (Rejected ????????????)
' Date: 2567
'
' ??????????? v5.1:
' - ????? CSV Export: ????????????? Check Valid
' - Rejected ???????????? ?????? 999999999
'================================================================================
Option Explicit

'================================================================================
' SECTION 1: Module-Level Variables (?????? HCA)
'================================================================================

' Current Design Indices
Private Currenttt As Integer
Private Currenttb As Integer
Private CurrentTBase As Integer
Private CurrentBase As Integer
Private CurrentLToe As Integer
Private CurrentStemDB As Integer
Private CurrentStemSP As Integer
Private CurrentToeDB As Integer
Private CurrentToeSP As Integer
Private CurrentHeelDB As Integer
Private CurrentHeelSP As Integer

' Full design-domain bounds. HCA never contracts these ranges.
Private FixedTbMax As Integer, FixedTBaseMax As Integer
Private FixedBaseMin As Integer, FixedBaseMax As Integer

' Cost History ?????? modDataStructures.CostHistory

' CSV Export
Private csvAcceptData As String
Private csvLoopData As String
Private loopCount As Long
Private bestIterationInRun As Long
' Retain a whole trial trace for the graph: lowest valid cost, then earliest loop.
Private bestAcceptData As String
Private bestAcceptTrial As Long
Private bestAcceptIteration As Long
Private bestAcceptCost As Double
Public SelectedTrialCSVPath As String

'================================================================================
' SECTION 2: Initialize Design (Common feasible reference)
'================================================================================

Private Sub InitializeCurrentDesign()
    Call SetCommonInitialIndices(Currenttt, Currenttb, CurrentTBase, CurrentBase, CurrentLToe, _
        CurrentStemDB, CurrentStemSP, CurrentToeDB, CurrentToeSP, CurrentHeelDB, CurrentHeelSP)
    FixedTbMax = tb_max: FixedTBaseMax = TBase_max
    FixedBaseMin = BASE_MIN: FixedBaseMax = BASE_MAX
End Sub

'================================================================================
' SECTION 3: Get Design from Current Indices
'================================================================================

Private Function GetDesignFromCurrent() As Design
    Dim d As Design
    
    d.tt = WP_tt(Currenttt)
    d.tb = WP_tb(Currenttb)
    d.TBase = WP_TBase(CurrentTBase)
    d.Base = WP_Base(CurrentBase)
    d.LToe = WP_LToe(CurrentLToe)
    d.LHeel = d.Base - d.LToe - d.tb
    
    ' Steel indices (for reference)
    d.ASst_DB = CurrentStemDB
    d.ASst_Sp = CurrentStemSP
    d.AStoe_DB = CurrentToeDB
    d.AStoe_Sp = CurrentToeSP
    d.ASheel_DB = CurrentHeelDB
    d.ASheel_Sp = CurrentHeelSP
    
    GetDesignFromCurrent = d
End Function

'================================================================================
' SECTION 4: Generate Neighbor (???????????????????)
'================================================================================

Private Sub GenerateNeighbor(ByRef Newtt As Integer, ByRef Newtb As Integer, _
                                ByRef NewTBase As Integer, ByRef NewBase As Integer, _
                                ByRef NewLToe As Integer, _
                                ByRef NewStemDB As Integer, ByRef NewStemSP As Integer, _
                                ByRef NewToeDB As Integer, ByRef NewToeSP As Integer, _
                                ByRef NewHeelDB As Integer, ByRef NewHeelSP As Integer)

    ' Same draw order, step sizes and repairs as BA, using the full domain.
    Dim Step As Integer, move(0 To 11) As Integer
    Dim LToe_min_idx As Integer, LToe_max_idx As Integer
    Dim i As Integer

    Call DrawSearchMove(move)

    LToe_min_idx = LTOE_MIN: LToe_max_idx = LTOE_MAX

    Step = move(2)
    Newtb = Currenttb + Step
    If Newtb < TB_MIN Then Newtb = TB_MIN
    If Newtb > FixedTbMax Then Newtb = FixedTbMax

    Step = move(1)
    Newtt = Currenttt + Step
    If Newtt < TT_MIN Then Newtt = TT_MIN
    If Newtt > TT_MAX Then Newtt = TT_MAX

    If WP_tt(Newtt) > WP_tb(Newtb) Then
        For i = TT_MAX To TT_MIN Step -1
            If WP_tt(i) <= WP_tb(Newtb) Then
                Newtt = i
                Exit For
            End If
        Next i
    End If

    Step = move(3)
    NewTBase = CurrentTBase + Step
    If NewTBase < TBASE_MIN Then NewTBase = TBASE_MIN
    If NewTBase > FixedTBaseMax Then NewTBase = FixedTBaseMax

    Step = move(5)
    NewLToe = CurrentLToe + Step
    If NewLToe < LToe_min_idx Then NewLToe = LToe_min_idx
    If NewLToe > LToe_max_idx Then NewLToe = LToe_max_idx

    Step = move(4)
    NewBase = CurrentBase + Step
    If NewBase < FixedBaseMin Then NewBase = FixedBaseMin
    If NewBase > FixedBaseMax Then NewBase = FixedBaseMax

    NewStemDB = CurrentStemDB + move(6)
    If NewStemDB < DB_MIN Then NewStemDB = DB_MIN
    If NewStemDB > DB_MAX Then NewStemDB = DB_MAX

    NewStemSP = CurrentStemSP + move(7)
    If NewStemSP < SP_MIN Then NewStemSP = SP_MIN
    If NewStemSP > SP_MAX Then NewStemSP = SP_MAX

    NewToeDB = CurrentToeDB + move(8)
    If NewToeDB < DB_MIN Then NewToeDB = DB_MIN
    If NewToeDB > DB_MAX Then NewToeDB = DB_MAX

    NewToeSP = CurrentToeSP + move(9)
    If NewToeSP < SP_MIN Then NewToeSP = SP_MIN
    If NewToeSP > SP_MAX Then NewToeSP = SP_MAX

    NewHeelDB = CurrentHeelDB + move(10)
    If NewHeelDB < DB_MIN Then NewHeelDB = DB_MIN
    If NewHeelDB > DB_MAX Then NewHeelDB = DB_MAX

    NewHeelSP = CurrentHeelSP + move(11)
    If NewHeelSP < SP_MIN Then NewHeelSP = SP_MIN
    If NewHeelSP > SP_MAX Then NewHeelSP = SP_MAX


    ' Keep components outside the selected move at the current state.
    If Not SearchMoveIncludes(move(0), 1) Then
        Newtt = Currenttt: Newtb = Currenttb
    End If
    If Not SearchMoveIncludes(move(0), 2) Then
        NewTBase = CurrentTBase: NewBase = CurrentBase: NewLToe = CurrentLToe
    End If
    If Not SearchMoveIncludes(move(0), 3) Then
        NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
    End If
    If Not SearchMoveIncludes(move(0), 4) Then
        NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
    End If
    If Not SearchMoveIncludes(move(0), 5) Then
        NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
    End If
    ' Keep dependent toe/heel layout compatible with the proposed width and stem.
    ' Apply the SAME geometry repair to BA and HCA; the validator remains decisive.
    If WP_Base(NewBase) - WP_tb(Newtb) - WP_LToe(NewLToe) <= WP_LToe(NewLToe) + 0.000000001 Then
        For i = NewLToe To LTOE_MIN Step -1
            If WP_Base(NewBase) - WP_tb(Newtb) - WP_LToe(i) > WP_LToe(i) + 0.000000001 Then
                NewLToe = i
                Exit For
            End If
        Next i
    End If

End Sub

'================================================================================
' SECTION 5: Hill Climbing Algorithm (Main Optimization Function)
' v5.1: Fixed CSV Export - ????????????? Check Valid
'================================================================================

Public Function HillClimbingOptimization(MaxIterations As Long, _
                       wall_height As Double, _
                       backfill_height As Double, _
                       soil_gamma As Double, _
                       concrete_gamma As Double, _
                       friction_angle As Double, _
                       friction_coef As Double, _
                       allowable_bearing As Double, _
                       concrete_cover As Double, _
                       material As MaterialProperties, _
                       Optional useSharedInit As Boolean = False, _
                       Optional sharedTt As Integer = 0, _
                       Optional sharedTb As Integer = 0, _
                       Optional sharedTBase As Integer = 0, _
                       Optional sharedBase As Integer = 0, _
                       Optional sharedLToe As Integer = 0, _
                       Optional sharedStemDB As Integer = 0, _
                       Optional sharedStemSP As Integer = 0, _
                       Optional sharedToeDB As Integer = 0, _
                       Optional sharedToeSP As Integer = 0, _
                       Optional sharedHeelDB As Integer = 0, _
                       Optional sharedHeelSP As Integer = 0, _
                       Optional TrialNumber As Long = 1, Optional ReportAsBA As Boolean = False) As Design

    Dim current As Design, neighbor As Design
    Dim currentCost As Double, neighborCost As Double, currentValid As Boolean, ok As Boolean
    Dim saved(1 To 11) As Integer
    Dim Newtt As Integer, Newtb As Integer, NewTBase As Integer, NewBase As Integer, NewLToe As Integer
    Dim NewStemDB As Integer, NewStemSP As Integer, NewToeDB As Integer, NewToeSP As Integer
    Dim NewHeelDB As Integer, NewHeelSP As Integer, i As Long, j As Integer
    modShared.H = wall_height: modShared.H1 = backfill_height
    modShared.gamma_soil = soil_gamma: modShared.gamma_concrete = concrete_gamma
    modShared.phi = friction_angle: modShared.mu = friction_coef
    modShared.qa = allowable_bearing: modShared.cover = concrete_cover
    modShared.currentMaterial = material
    modShared.currentWSD = CalculateWSDParameters(material.fy, material.fc)
    Call InitializeArrays
    ReDim modDataStructures.CostHistory(1 To MaxIterations)
    ' Reporting only: both buttons use the unchanged HCA search below.
    If ReportAsBA Then
        Call BeginSearch(MaxIterations, "BA", TrialNumber)
    Else
        Call BeginSearch(MaxIterations, "HCA", TrialNumber)
    End If
    Call InitializeCurrentDesign
    If useSharedInit Then
        If sharedtt < TT_MIN Or sharedtt > TT_MAX Then Err.Raise 5, , "Invalid shared tt index"
        Currenttt = sharedtt
        If sharedtb < TB_MIN Or sharedtb > tb_max Then Err.Raise 5, , "Invalid shared tb index"
        Currenttb = sharedtb
        If sharedTBase < TBASE_MIN Or sharedTBase > TBase_max Then Err.Raise 5, , "Invalid shared TBase index"
        CurrentTBase = sharedTBase
        If sharedBase < BASE_MIN Or sharedBase > BASE_MAX Then Err.Raise 5, , "Invalid shared Base index"
        CurrentBase = sharedBase
        If sharedLToe < LTOE_MIN Or sharedLToe > LTOE_MAX Then Err.Raise 5, , "Invalid shared LToe index"
        CurrentLToe = sharedLToe
        If sharedStemDB < DB_MIN Or sharedStemDB > DB_MAX Then Err.Raise 5, , "Invalid shared StemDB index"
        CurrentStemDB = sharedStemDB
        If sharedStemSP < SP_MIN Or sharedStemSP > SP_MAX Then Err.Raise 5, , "Invalid shared StemSP index"
        CurrentStemSP = sharedStemSP
        If sharedToeDB < DB_MIN Or sharedToeDB > DB_MAX Then Err.Raise 5, , "Invalid shared ToeDB index"
        CurrentToeDB = sharedToeDB
        If sharedToeSP < SP_MIN Or sharedToeSP > SP_MAX Then Err.Raise 5, , "Invalid shared ToeSP index"
        CurrentToeSP = sharedToeSP
        If sharedHeelDB < DB_MIN Or sharedHeelDB > DB_MAX Then Err.Raise 5, , "Invalid shared HeelDB index"
        CurrentHeelDB = sharedHeelDB
        If sharedHeelSP < SP_MIN Or sharedHeelSP > SP_MAX Then Err.Raise 5, , "Invalid shared HeelSP index"
        CurrentHeelSP = sharedHeelSP
    End If
    current = GetDesignFromCurrent()
    Call RequireFeasibleInitial(current)
    currentValid = EvaluateCandidate(current, "initial", currentCost)
    If Not currentValid Then currentCost = NO_SOLUTION_COST
    Call modFeasibilityRecovery.InitializeRecovery(current, currentValid)
    modDataStructures.CostHistory(EvaluationCount) = RunBestCost
    Do While EvaluationCount < MaxIterations
        saved(1) = Currenttt
        saved(2) = Currenttb
        saved(3) = CurrentTBase
        saved(4) = CurrentBase
        saved(5) = CurrentLToe
        saved(6) = CurrentStemDB
        saved(7) = CurrentStemSP
        saved(8) = CurrentToeDB
        saved(9) = CurrentToeSP
        saved(10) = CurrentHeelDB
        saved(11) = CurrentHeelSP
        Call GenerateNeighbor(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP)
        Currenttt = Newtt
        Currenttb = Newtb
        CurrentTBase = NewTBase
        CurrentBase = NewBase
        CurrentLToe = NewLToe
        CurrentStemDB = NewStemDB
        CurrentStemSP = NewStemSP
        CurrentToeDB = NewToeDB
        CurrentToeSP = NewToeSP
        CurrentHeelDB = NewHeelDB
        CurrentHeelSP = NewHeelSP
        neighbor = GetDesignFromCurrent()
        ok = EvaluateCandidate(neighbor, "neighbor", neighborCost)
        If ok And (Not currentValid Or neighborCost < currentCost) Then
            current = neighbor: currentCost = neighborCost: currentValid = True
            modFeasibilityRecovery.RecoveryActive = False
        ElseIf Not currentValid And modFeasibilityRecovery.AcceptRecovery(neighbor) Then
            current = neighbor: currentCost = NO_SOLUTION_COST
        Else
            Currenttt = saved(1)
            Currenttb = saved(2)
            CurrentTBase = saved(3)
            CurrentBase = saved(4)
            CurrentLToe = saved(5)
            CurrentStemDB = saved(6)
            CurrentStemSP = saved(7)
            CurrentToeDB = saved(8)
            CurrentToeSP = saved(9)
            CurrentHeelDB = saved(10)
            CurrentHeelSP = saved(11)
        End If
        modDataStructures.CostHistory(EvaluationCount) = RunBestCost
        DoEvents
    Loop
    Call FinishSearch
    HillClimbingOptimization = RunBest
End Function

'================================================================================
' SECTION 6: CSV Export Functions
'================================================================================

Public Sub InitCSVExport()
    csvAcceptData = "No.,Rejected,Passed,Passed and Better value" & vbCrLf
    bestIterationInRun = 0
End Sub

Public Sub InitLoopCounter()
    csvLoopData = "No.,Loop,BestPrice" & vbCrLf
    LastLoopCSVPath = ""
    loopCount = 0
    bestAcceptData = ""
    bestAcceptTrial = 0
    bestAcceptIteration = 0
    bestAcceptCost = NO_SOLUTION_COST
    SelectedTrialCSVPath = ""
End Sub

'--------------------------------------------------------------------------------
' LogIteration - ????????????? iteration
' - Rejected: ??????? column 2 (Invalid)
' - Passed: ??????? column 3 (Valid ???????????? best)
' - Passed and Better value: ??????? column 4 (Valid ????????? best)
'--------------------------------------------------------------------------------
Public Sub LogIteration(iteration As Long, cost As String, IsValid As Boolean, isBetter As Boolean)
    If Not IsValid Then
        ' Rejected - ??????? column 2
        csvAcceptData = csvAcceptData & iteration & "," & cost & ",," & vbCrLf
    ElseIf isBetter Then
        ' Passed and Better value - ??????? column 4
        csvAcceptData = csvAcceptData & iteration & ",,," & cost & vbCrLf
        bestIterationInRun = iteration
    Else
        ' Passed - ??????? column 3
        csvAcceptData = csvAcceptData & iteration & ",," & cost & "," & vbCrLf
    End If
End Sub

'--------------------------------------------------------------------------------
' LogLoopResult - ???????? Trial
' No., Loop (iteration ????? best), BestPrice
'--------------------------------------------------------------------------------
Public Sub LogLoopResult(bestPrice As Double)
    Dim price As String
    loopCount = loopCount + 1
    If RunBest.IsValid Then price = CsvPrice(bestPrice)
    csvLoopData = csvLoopData & loopCount & "," & bestIterationInRun & "," & price & vbCrLf
    ' Use unrounded cost, matching Form1's best-trial selection. Keep the first
    ' trial when both cost and loop tie. Invalid trials remain in loopPrice only.
    If RunBest.IsValid Then
        If bestAcceptTrial = 0 Or bestPrice < bestAcceptCost Or _
           (bestPrice = bestAcceptCost And bestIterationInRun < bestAcceptIteration) Then
            bestAcceptData = csvAcceptData
            bestAcceptTrial = loopCount
            bestAcceptIteration = bestIterationInRun
            bestAcceptCost = bestPrice
        End If
    End If
End Sub

Public Sub SaveAcceptCSV(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-HCA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csvAcceptData, True)
End Sub

Public Sub SaveLoopPriceCSV(wallHeight As Double)
    LastLoopCSVPath = WriteExportCSV("loopPrice-HCA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csvLoopData, True)
    Call SaveSelectedTrialCSV(wallHeight)
End Sub

Private Sub SaveSelectedTrialCSV(wallHeight As Double)
    Dim caseSuffix As String, selectedData As String
    caseSuffix = "HCA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc)
    selectedData = "No.,Loop,BestPrice,Trials,Status" & vbCrLf
    If bestAcceptTrial > 0 Then
        ' FinishSearch already saved the last trial. Replace it only if another
        ' trial won; WriteExportCSV preserves that last trace in the archive.
        If bestAcceptTrial <> loopCount Then
            LastAcceptCSVPath = WriteExportCSV("accept-" & caseSuffix, bestAcceptData, True)
        End If
        selectedData = selectedData & bestAcceptTrial & "," & bestAcceptIteration & "," & _
                       CsvPrice(bestAcceptCost) & "," & loopCount & ",SELECTED" & vbCrLf
    Else
        ' Retain the last rejected trace for diagnosis; do not invent a winner.
        selectedData = selectedData & ",,," & loopCount & ",NO_SOLUTION" & vbCrLf
    End If
    SelectedTrialCSVPath = WriteExportCSV("selectedTrial-" & caseSuffix, selectedData, True)
End Sub

'================================================================================
' END OF MODULE: modHillClimbing.bas v5.1
'================================================================================

Public Sub AssertSteelNeighborhood(ByVal dbIndex As Integer, ByVal spIndex As Integer)
    Dim nt As Integer, nb As Integer, nf As Integer, nw As Integer, nl As Integer
    Dim sd As Integer, ss As Integer, td As Integer, ts As Integer, hd As Integer, hs As Integer
    Dim k As Integer, seenLower As Boolean, seenUpper As Boolean, seenSPDown As Boolean, seenSPUp As Boolean
    Call InitializeCurrentDesign
    CurrentStemDB = dbIndex: CurrentToeDB = dbIndex: CurrentHeelDB = dbIndex
    CurrentStemSP = spIndex: CurrentToeSP = spIndex: CurrentHeelSP = spIndex

    For k = 1 To 100
        Call GenerateNeighbor(nt, nb, nf, nw, nl, sd, ss, td, ts, hd, hs)
        If sd < DB_MIN Or sd > DB_MAX Or td < DB_MIN Or td > DB_MAX Or hd < DB_MIN Or hd > DB_MAX Then Err.Raise 5, , "Steel outside catalogue"
        If ss < SP_MIN Or ss > SP_MAX Or ts < SP_MIN Or ts > SP_MAX Or hs < SP_MIN Or hs > SP_MAX Then Err.Raise 5, , "Spacing outside catalogue"
        If Abs(sd - dbIndex) > 1 Or Abs(td - dbIndex) > 1 Or Abs(hd - dbIndex) > 1 Then Err.Raise 5, , "Steel jumped more than one index"
        If Abs(ss - spIndex) > 1 Or Abs(ts - spIndex) > 1 Or Abs(hs - spIndex) > 1 Then Err.Raise 5, , "Spacing jumped more than one index"
        If sd = dbIndex - 1 Or td = dbIndex - 1 Or hd = dbIndex - 1 Then seenLower = True
        If sd = dbIndex + 1 Or td = dbIndex + 1 Or hd = dbIndex + 1 Then seenUpper = True
        If ss = spIndex - 1 Or ts = spIndex - 1 Or hs = spIndex - 1 Then seenSPDown = True
        If ss = spIndex + 1 Or ts = spIndex + 1 Or hs = spIndex + 1 Then seenSPUp = True
        ' Deliberately keep the incumbent: rejected candidates must not accumulate.
    Next k
    If dbIndex > DB_MIN And Not seenLower Then Err.Raise 5, , "Lower adjacent steel not reached"
    If dbIndex < DB_MAX And Not seenUpper Then Err.Raise 5, , "Upper adjacent steel not reached"
    If spIndex > SP_MIN And Not seenSPDown Then Err.Raise 5, , "Lower adjacent spacing not reached"
    If spIndex < SP_MAX And Not seenSPUp Then Err.Raise 5, , "Upper adjacent spacing not reached"
End Sub
