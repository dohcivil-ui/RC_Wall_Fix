Attribute VB_Name = "modBA"
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
Private PairDB() As Integer, PairSP() As Integer, PairWeight() As Double
Private PairCount As Integer
Public CostHistory_BA() As Double

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

' Cost History ?????? CostHistory_BA

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

    ' Five one-level geometry moves plus three simultaneous coupled-steel moves.
    Dim Step As Integer, move(0 To 11) As Integer
    Dim LToe_min_idx As Integer, LToe_max_idx As Integer
    Dim i As Integer

    Call DrawGeometryMove(move)

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

    ' Do not move tt by extra levels: unchanged GeometryOK rejects tt > tb.

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

    Call PairNeighbor(CurrentStemDB, CurrentStemSP, NewStemDB, NewStemSP)
    Call PairNeighbor(CurrentToeDB, CurrentToeSP, NewToeDB, NewToeSP)
    Call PairNeighbor(CurrentHeelDB, CurrentHeelSP, NewHeelDB, NewHeelSP)

End Sub

'================================================================================
' SECTION 5: Hill Climbing Algorithm (Main Optimization Function)
' v5.1: Fixed CSV Export - ????????????? Check Valid
'================================================================================

Public Function BisectionOptimization(MaxIterations As Long, _
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
                       Optional TrialNumber As Long = 1) As Design

    Dim lo(1 To 3) As Integer, hi(1 To 3) As Integer
    Dim phase As Long, phaseMoves As Long, needMidpoint As Boolean
    Dim referenceCost As Double
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
    Call InitializePairOptions
    ReDim CostHistory_BA(1 To MaxIterations)
    Call BeginSearch(MaxIterations, "BA", TrialNumber)
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
    CostHistory_BA(EvaluationCount) = RunBestCost
    lo(1) = TB_MIN: hi(1) = FixedTbMax
    lo(2) = TBASE_MIN: hi(2) = FixedTBaseMax
    lo(3) = FixedBaseMin: hi(3) = FixedBaseMax
    phase = 1: phaseMoves = 0: needMidpoint = True
    referenceCost = NO_SOLUTION_COST
    Do While EvaluationCount < MaxIterations
        If needMidpoint Then
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
            Call MoveToCenter(lo, hi)
            neighbor = GetDesignFromCurrent()
            ok = EvaluateCandidate(neighbor, "midpoint", neighborCost)
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
            CostHistory_BA(EvaluationCount) = RunBestCost
            phaseMoves = 0: needMidpoint = False
        Else
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
        CostHistory_BA(EvaluationCount) = RunBestCost
        phaseMoves = phaseMoves + 1
        If phaseMoves = 20 * phase And EvaluationCount < MaxIterations Then
            Call UpdateCenterBounds(lo, hi, currentCost, currentValid, referenceCost)
            phase = phase + 1: needMidpoint = True
        End If
        End If
        DoEvents
    Loop
    Call FinishSearch
    BisectionOptimization = RunBest
End Function


' EXPERIMENT: center relocation + full-domain HCA, adapted from the supplied column.
' All three centers change together. Center bounds do NOT clamp random proposals.
' Every relocated state is evaluated and counted; global best never resets.

Private Function NearestCenter(values() As Double, lo As Integer, hi As Integer) As Integer
    Dim center As Double, distance As Double, bestDistance As Double, i As Integer
    center = (values(lo) + values(hi)) / 2
    bestDistance = 1E+30
    For i = lo To hi
        distance = Abs(values(i) - center)
        ' At a numerical tie use the lower existing size; no new design sizes.
        If distance < bestDistance - 0.000000000001 Then
            NearestCenter = i: bestDistance = distance
        End If
    Next i
End Function

Private Sub MoveToCenter(lo() As Integer, hi() As Integer)
    Dim i As Integer
    Currenttb = NearestCenter(WP_tb, lo(1), hi(1))
    CurrentTBase = NearestCenter(WP_TBase, lo(2), hi(2))
    CurrentBase = NearestCenter(WP_Base, lo(3), hi(3))
    ' Preserve the wall's existing tt <= tb compatibility repair.
    If WP_tt(Currenttt) > WP_tb(Currenttb) Then
        For i = TT_MAX To TT_MIN Step -1
            If WP_tt(i) <= WP_tb(Currenttb) Then Currenttt = i: Exit For
        Next i
    End If
End Sub

Private Sub UpdateCenterBounds(lo() As Integer, hi() As Integer, cost As Double, valid As Boolean, referenceCost As Double)
    Dim axis As Integer, swap As Integer
    If Not valid Then Exit Sub
    If cost < referenceCost Then
        referenceCost = cost
        hi(1) = Currenttb: hi(2) = CurrentTBase: hi(3) = CurrentBase
    Else
        lo(1) = Currenttb: lo(2) = CurrentTBase: lo(3) = CurrentBase
    End If
    ' A global-domain walk may cross the old center bounds. Keep them ordered.
    For axis = 1 To 3
        If lo(axis) > hi(axis) Then
            swap = lo(axis): lo(axis) = hi(axis): hi(axis) = swap
        End If
    Next axis
End Sub

Public Function GetConcretePrice_BA(fc As Integer) As Double
    GetConcretePrice_BA = modShared.GetConcretePrice(fc)
End Function

'--------------------------------------------------------------------------------
' Get SD40 Material Properties
'--------------------------------------------------------------------------------
Public Function GetSD40Material_BA(fc As Integer, _
                                   concPrice As Double, _
                                   steelPrice As Double) As MaterialProperties
    GetSD40Material_BA = modShared.GetSD40Material(fc, concPrice, steelPrice)
End Function

Public Sub InitCSVExport_BA()
    csvAcceptData = "No.,Rejected,Passed,Passed and Better value" & vbCrLf
    bestIterationInRun = 0
End Sub

Public Sub InitLoopCounter_BA()
    csvLoopData = "No.,Loop,BestPrice" & vbCrLf
    LastLoopCSVPath = ""
    loopCount = 0
End Sub

'--------------------------------------------------------------------------------
' v3.1: LogIteration_BA - ???????????? HCA
' - Rejected: ??????? column 2 (Invalid)
' - Passed: ??????? column 3 (Valid ???????????? best)
' - Passed and Better value: ??????? column 4 (Valid ????????? best)
'--------------------------------------------------------------------------------
Public Sub LogIteration_BA(iteration As Long, cost As String, IsValid As Boolean, isBetter As Boolean)
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
' v3.1: LogLoopResult_BA - ???????????? HCA
' ??????: No., Loop (iteration ????? best), BestPrice
'--------------------------------------------------------------------------------
Public Sub LogLoopResult_BA(bestPrice As Double)
    Dim price As String
    loopCount = loopCount + 1
    If RunBest.IsValid Then price = CsvPrice(bestPrice)
    csvLoopData = csvLoopData & loopCount & "," & bestIterationInRun & "," & price & vbCrLf
End Sub

Public Sub SaveAcceptCSV_BA(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-BA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csvAcceptData, True)
End Sub

Public Sub SaveLoopPriceCSV_BA(wallHeight As Double)
    LastLoopCSVPath = WriteExportCSV("loopPrice-BA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csvLoopData, True)
End Sub


' Only sampler representation changes; all twenty original DB/SP pairs remain.
Private Sub InitializePairOptions()
    Dim db As Integer, sp As Integer, i As Integer, j As Integer
    Dim tempIndex As Integer, tempWeight As Double
    PairCount = (DB_MAX - DB_MIN + 1) * (SP_MAX - SP_MIN + 1)
    ReDim PairDB(1 To PairCount): ReDim PairSP(1 To PairCount): ReDim PairWeight(1 To PairCount)
    i = 0
    For db = DB_MIN To DB_MAX
        For sp = SP_MIN To SP_MAX
            i = i + 1: PairDB(i) = db: PairSP(i) = sp
            PairWeight(i) = ProvidedSteelWeight(db, sp, 1#)
        Next sp
    Next db
    For i = 1 To PairCount - 1
        For j = i + 1 To PairCount
            If PairWeight(j) < PairWeight(i) Or _
               (PairWeight(j) = PairWeight(i) And PairDB(j) < PairDB(i)) Or _
               (PairWeight(j) = PairWeight(i) And PairDB(j) = PairDB(i) And PairSP(j) < PairSP(i)) Then
                tempWeight = PairWeight(i): PairWeight(i) = PairWeight(j): PairWeight(j) = tempWeight
                tempIndex = PairDB(i): PairDB(i) = PairDB(j): PairDB(j) = tempIndex
                tempIndex = PairSP(i): PairSP(i) = PairSP(j): PairSP(j) = tempIndex
            End If
        Next j
    Next i
End Sub

Private Function BuildPairChoices(db As Integer, sp As Integer, choices() As Integer) As Integer
    Dim rank As Integer, i As Integer, count As Integer
    For i = 1 To PairCount
        If PairDB(i) = db And PairSP(i) = sp Then rank = i: Exit For
    Next i
    If rank = 0 Then Err.Raise 5, , "Unknown current steel pair"
    ReDim choices(1 To PairCount)
    For i = 1 To PairCount
        If (Abs(PairDB(i) - db) <= 1 And Abs(PairSP(i) - sp) <= 1) Or Abs(i - rank) <= 1 Then
            count = count + 1: choices(count) = i
        End If
    Next i
    BuildPairChoices = count
End Function

Private Sub PairNeighbor(db As Integer, sp As Integer, newDB As Integer, newSP As Integer)
    Dim choices() As Integer, count As Integer, ticket As Integer, i As Integer
    count = BuildPairChoices(db, sp, choices)
    newDB = db: newSP = sp
    If count <= 1 Then Exit Sub
    ' Independent draw for every member: half stay, half select another pair.
    ' All previous non-self choices remain equally likely and reachable.
    ticket = Rand(1, 2 * (count - 1))
    If ticket <= count - 1 Then Exit Sub
    ticket = ticket - (count - 1)
    For i = 1 To count
        If PairDB(choices(i)) <> db Or PairSP(choices(i)) <> sp Then
            ticket = ticket - 1
            If ticket = 0 Then
                newDB = PairDB(choices(i)): newSP = PairSP(choices(i))
                Exit Sub
            End If
        End If
    Next i
    Err.Raise 5, , "Invalid coupled-pair draw"
End Sub

Private Sub DrawGeometryMove(move() As Integer)
    move(2) = GeometryStep(): move(1) = GeometryStep()
    move(3) = GeometryStep(): move(5) = GeometryStep(): move(4) = GeometryStep()
End Sub

Private Function GeometryStep() As Integer
    ' Six equiprobable tickets: one step down, four stay, one step up.
    ' Every geometry variable draws independently before the full evaluation.
    Select Case Rand(1, 6)
        Case 1: GeometryStep = -1
        Case 6: GeometryStep = 1
        Case Else: GeometryStep = 0
    End Select
End Function
