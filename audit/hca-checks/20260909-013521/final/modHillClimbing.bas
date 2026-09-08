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

Public Const HCA_SEARCH_POLICY As String = "HCA_FINAL_MAIN_STEEL_SWEEP_V1"
Public HCARefinementCount As Long

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
    
    Dim Step As Integer
    Dim LBase_Max_Ratio As Double, Base_Max_calc As Double
    Dim LHeel_Max As Double
    
    ' === ????????????? (?????? HCA ???) ===
    
    ' tt: step = Rand(-2, 2)
    Step = Rand(-2, 2)
    Newtt = Currenttt + Step
    If Newtt < TT_MIN Then Newtt = TT_MIN
    If Newtt > TT_MAX Then Newtt = TT_MAX
    
    ' tb: step = Rand(-1, 1)
    Step = Rand(-2, 2)
    Newtb = Currenttb + Step
    If Newtb < TB_MIN Then Newtb = TB_MIN
    If Newtb > tb_max Then Newtb = tb_max
    ' Constraint: tb >= tt
    If WP_tb(Newtb) < WP_tt(Newtt) Then
        Newtb = TB_MIN
        Do While Newtb <= tb_max
            If WP_tb(Newtb) >= WP_tt(Newtt) Then Exit Do
            Newtb = Newtb + 1
        Loop
    End If
    ' Constraint: tb <= 0.12H
    Do While Newtb > TB_MIN And WP_tb(Newtb) > 0.12 * modShared.H
        Newtb = Newtb - 1
    Loop
    
    ' TBase: step = Rand(-1, 1)
    Step = Rand(-5, 5)
    NewTBase = CurrentTBase + Step
    If NewTBase < TBASE_MIN Then NewTBase = TBASE_MIN
    If NewTBase > TBase_max Then NewTBase = TBase_max
    ' Constraint: TBase <= 0.15H
    Do While NewTBase > TBASE_MIN And WP_TBase(NewTBase) > 0.15 * modShared.H
        NewTBase = NewTBase - 1
    Loop
    
    ' LToe: step = Rand(-2, 2)
    Step = Rand(-2, 2)
    NewLToe = CurrentLToe + Step
    If NewLToe < LTOE_MIN Then NewLToe = LTOE_MIN
    If NewLToe > LTOE_MAX Then NewLToe = LTOE_MAX
    ' Constraint: 0.1H <= LToe <= 0.2H
    Do While NewLToe > LTOE_MIN And WP_LToe(NewLToe) > 0.2 * modShared.H
        NewLToe = NewLToe - 1
    Loop
    Do While NewLToe < LTOE_MAX And WP_LToe(NewLToe) < 0.1 * modShared.H
        NewLToe = NewLToe + 1
    Loop
    
    ' Base: step = Rand(-1, 1)
    ' Constraint: 0.5H <= Base <= 0.7H
    Base_Max_calc = 0.7 * modShared.H
    Step = Rand(-1, 1)
    NewBase = CurrentBase + Step
    If NewBase < BASE_MIN Then NewBase = BASE_MIN
    If NewBase > BASE_MAX Then NewBase = BASE_MAX
    ' Constraint: Base >= 0.5H
    Do While NewBase < BASE_MAX And WP_Base(NewBase) < 0.5 * modShared.H
        NewBase = NewBase + 1
    Loop
    ' Constraint: Base <= Base_Max_calc
    Do While NewBase > BASE_MIN And WP_Base(NewBase) > Base_Max_calc
        NewBase = NewBase - 1
    Loop
    
    ' === Steel: step = Rand(-2, 2) ===
    Step = Rand(-2, 2)
    NewStemDB = CurrentStemDB + Step
    If NewStemDB < DB_MIN Then NewStemDB = DB_MIN
    If NewStemDB > DB_MAX Then NewStemDB = DB_MAX
    
    Step = Rand(-2, 2)
    NewStemSP = CurrentStemSP + Step
    If NewStemSP < SP_MIN Then NewStemSP = SP_MIN
    If NewStemSP > SP_MAX Then NewStemSP = SP_MAX
    
    Step = Rand(-2, 2)
    NewToeDB = CurrentToeDB + Step
    If NewToeDB < DB_MIN Then NewToeDB = DB_MIN
    If NewToeDB > DB_MAX Then NewToeDB = DB_MAX
    
    Step = Rand(-2, 2)
    NewToeSP = CurrentToeSP + Step
    If NewToeSP < SP_MIN Then NewToeSP = SP_MIN
    If NewToeSP > SP_MAX Then NewToeSP = SP_MAX
    
    Step = Rand(-2, 2)
    NewHeelDB = CurrentHeelDB + Step
    If NewHeelDB < DB_MIN Then NewHeelDB = DB_MIN
    If NewHeelDB > DB_MAX Then NewHeelDB = DB_MAX
    
    Step = Rand(-2, 2)
    NewHeelSP = CurrentHeelSP + Step
    If NewHeelSP < SP_MIN Then NewHeelSP = SP_MIN
    If NewHeelSP > SP_MAX Then NewHeelSP = SP_MAX
    
End Sub

Private Function MainSteelNeighbor(d As Design, visit As Long) As Design
    ' Visit each existing DB/spacing pair for one member; keep geometry fixed.
    Dim pairs As Long, pair As Long, DB As Integer, SP As Integer, candidate As Design
    candidate = d
    pairs = (DB_MAX - DB_MIN + 1) * (SP_MAX - SP_MIN + 1)
    pair = (visit - 1) Mod pairs
    DB = DB_MIN + pair \ (SP_MAX - SP_MIN + 1)
    SP = SP_MIN + pair Mod (SP_MAX - SP_MIN + 1)
    Select Case (visit - 1) \ pairs
        Case 0: candidate.ASst_DB = DB: candidate.ASst_Sp = SP
        Case 1: candidate.AStoe_DB = DB: candidate.AStoe_Sp = SP
        Case 2: candidate.ASheel_DB = DB: candidate.ASheel_Sp = SP
    End Select
    MainSteelNeighbor = candidate
End Function

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
    Dim refinementVisit As Long, refinementVisits As Long, entry As String
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
    HCARefinementCount = 0
    refinementVisits = 3 * (DB_MAX - DB_MIN + 1) * (SP_MAX - SP_MIN + 1)
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
        ' Reserve one final 60-candidate steel sweep, counted inside the same budget.
        ' Short runs keep the original random neighborhood throughout.
        If MaxIterations > 2 * refinementVisits And MaxIterations - EvaluationCount = refinementVisits Then
            refinementVisit = 1
            HCARefinementCount = 1
        End If
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
        If refinementVisit > 0 Then
            entry = "steel_refine"
            neighbor = MainSteelNeighbor(current, refinementVisit)
            CurrentStemDB = neighbor.ASst_DB: CurrentStemSP = neighbor.ASst_Sp
            CurrentToeDB = neighbor.AStoe_DB: CurrentToeSP = neighbor.AStoe_Sp
            CurrentHeelDB = neighbor.ASheel_DB: CurrentHeelSP = neighbor.ASheel_Sp
        Else
            entry = "neighbor"
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
        End If
        ok = EvaluateCandidate(neighbor, entry, neighborCost)
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
        If refinementVisit > 0 Then
            refinementVisit = refinementVisit + 1
            If refinementVisit > refinementVisits Then refinementVisit = 0
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
