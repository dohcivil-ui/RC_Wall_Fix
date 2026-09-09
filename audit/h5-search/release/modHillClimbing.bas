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

'================================================================================
' SECTION 2: Initialize Design (Conservative - Max values)
'================================================================================

Private Sub InitializeCurrentDesign()
    Dim tb_max_val As Double, TBase_max_val As Double
    Dim Base_target As Double, LToe_target As Double
    Dim LBase_Max_Ratio As Double
    Dim i As Integer
    
    
    ' === tb: Max = 0.12 × H ===
    tb_max_val = 0.12 * modShared.H
    Currenttb = TB_MIN
    For i = modShared.tb_max To modShared.TB_MIN Step -1
        If WP_tb(i) <= tb_max_val Then
            Currenttb = i
            Exit For
        End If
    Next i
    
    ' === tt: ???????? Max ???? <= tb ===
Currenttt = TT_MAX
For i = modShared.TT_MAX To modShared.TT_MIN Step -1
    If WP_tt(i) <= WP_tb(Currenttb) Then
        Currenttt = i
        Exit For
    End If
Next i

    ' ???? >= tt
    If WP_tb(Currenttb) < WP_tt(Currenttt) Then
        For i = modShared.TB_MIN To modShared.tb_max
            If WP_tb(i) >= WP_tt(Currenttt) Then
                Currenttb = i
                Exit For
            End If
        Next i
    End If
    
    ' === TBase: Max = 0.15 × H ===
    TBase_max_val = 0.15 * modShared.H
    CurrentTBase = TBASE_MIN
    For i = modShared.TBase_max To modShared.TBASE_MIN Step -1
        If WP_TBase(i) <= TBase_max_val Then
            CurrentTBase = i
            Exit For
        End If
    Next i
    
    ' === Base: 0.5H - 0.7H ===
    Base_target = 0.7 * modShared.H
    CurrentBase = BASE_MIN
    For i = modShared.BASE_MAX To modShared.BASE_MIN Step -1
        If WP_Base(i) <= Base_target Then
            CurrentBase = i
            Exit For
        End If
    Next i
    
    ' === LToe: 0.2 × H ===
    LToe_target = 0.2 * modShared.H
    CurrentLToe = LTOE_MIN
    For i = modShared.LTOE_MAX To modShared.LTOE_MIN Step -1
        If WP_LToe(i) <= LToe_target Then
            CurrentLToe = i
            Exit For
        End If
    Next i
    
    ' === ?????: Max (DB28 @ 0.10m) ===
    CurrentStemDB = DB_MAX:  CurrentStemSP = SP_MIN
    CurrentToeDB = DB_MAX:   CurrentToeSP = SP_MIN
    CurrentHeelDB = DB_MAX:  CurrentHeelSP = SP_MIN

    FixedTbMax = Currenttb: FixedTBaseMax = CurrentTBase
    FixedBaseMax = CurrentBase: FixedBaseMin = BASE_MIN
    For i = BASE_MIN To BASE_MAX
        If WP_Base(i) >= 0.5 * modShared.H Then FixedBaseMin = i: Exit For
    Next i
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

    LToe_min_idx = LTOE_MIN
    For i = LTOE_MIN To LTOE_MAX
        If WP_LToe(i) >= 0.1 * modShared.H Then
            LToe_min_idx = i
            Exit For
        End If
    Next i

    LToe_max_idx = LTOE_MAX
    For i = LTOE_MAX To LTOE_MIN Step -1
        If WP_LToe(i) <= 0.2 * modShared.H Then
            LToe_max_idx = i
            Exit For
        End If
    Next i

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

    NewStemDB = move(6)
    If NewStemDB < DB_MIN Then NewStemDB = DB_MIN
    If NewStemDB > DB_MAX Then NewStemDB = DB_MAX

    NewStemSP = move(7)
    If NewStemSP < SP_MIN Then NewStemSP = SP_MIN
    If NewStemSP > SP_MAX Then NewStemSP = SP_MAX

    NewToeDB = move(8)
    If NewToeDB < DB_MIN Then NewToeDB = DB_MIN
    If NewToeDB > DB_MAX Then NewToeDB = DB_MAX

    NewToeSP = move(9)
    If NewToeSP < SP_MIN Then NewToeSP = SP_MIN
    If NewToeSP > SP_MAX Then NewToeSP = SP_MAX

    NewHeelDB = move(10)
    If NewHeelDB < DB_MIN Then NewHeelDB = DB_MIN
    If NewHeelDB > DB_MAX Then NewHeelDB = DB_MAX

    NewHeelSP = move(11)
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
                       Optional RandomSeed As Long = 12345, _
                       Optional TrialNumber As Long = 1) As Design

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
    Call BeginSearch(MaxIterations, RandomSeed, "HCA", TrialNumber)
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
    currentValid = EvaluateCandidate(current, "initial", currentCost)
    If Not currentValid Then currentCost = NO_SOLUTION_COST
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
End Sub

Public Sub SaveAcceptCSV(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-HCA-H" & Replace$(CStr(wallHeight), ",", "."), csvAcceptData)
End Sub

Public Sub SaveLoopPriceCSV(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-HCA-H" & Replace$(CStr(wallHeight), ",", "."), csvAcceptData, True)
    LastLoopCSVPath = WriteExportCSV("loopPrice-HCA-H" & Replace$(CStr(wallHeight), ",", "."), csvLoopData, True)
End Sub

'================================================================================
' END OF MODULE: modHillClimbing.bas v5.1
'================================================================================

Public Sub HCAParity(state() As Integer, sampleSeed As Long, result() As Integer, ByRef nextRandom As Single)
    Dim dummy As Single
    Call InitializeCurrentDesign
    Currenttt = state(1): Currenttb = state(2): CurrentTBase = state(3)
    CurrentBase = state(4): CurrentLToe = state(5)
    CurrentStemDB = state(6): CurrentStemSP = state(7)
    CurrentToeDB = state(8): CurrentToeSP = state(9)
    CurrentHeelDB = state(10): CurrentHeelSP = state(11)
    dummy = Rnd(-1): Randomize sampleSeed
    Call GenerateNeighbor(result(1), result(2), result(3), result(4), result(5), result(6), result(7), result(8), result(9), result(10), result(11))
    nextRandom = Rnd
End Sub
