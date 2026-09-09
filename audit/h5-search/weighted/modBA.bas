Attribute VB_Name = "modBA"
'================================================================================
' Module: modBA.bas
' Project: RC_RT_HCA v2.8 - Cantilever Retaining Wall Optimization
' ?????????: ?????????????????? (Bisection Algorithm - BA) - ???¡?? functions ?? modShared
'??????: 3.1 - Triple Bisection + ???????????????? CSV
'?????: 2567

'?????÷?????????? 3#:
'Triple Bisection: ???????? Base, TBase, ??? tb ?¡????????? (Bisection) ??????????þ??????

'HCA-style: ??????????????????? (tt, LToe) ??????????????????????? 6 ?????

'Constraint (??????): ??????? tt <= tb (???? Clamp ??????????????????)

'Inner Loop: ??? 20 ? Countloop ?? (????????????????? ??????÷? Bisection)

'????????????? (Mid): ????????????? HCA ???????????????????? (Max)

'??û???????????? 3.1:
'??? CSV Export: ???????????????? HCA ?????ä?????????????? ????????µ????????????? (Check Valid)

'?????????????? (Rejected): ???????????????????????? ??????????????? (999999999) ??????
'================================================================================
Option Explicit

'================================================================================
' SECTION 1: Module-Level Variables
'================================================================================

' Current Design Indices (????? HCA)
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

' Bisection Variables (????? Base)
Private MinBase As Integer      ' Index ??????? Base
Private MaxBase As Integer      ' Index **??????** Base
Private MidBase As Integer      ' Index **?????** Base
Private MidPrice As Double      ' ???? MidBase

' Bisection Variables (?????? TBase) - v2.0
Private MinTBase As Integer     ' Index ????????? TBase
Private MaxTBase As Integer     ' Index ????????? TBase
Private MidTBase As Integer     ' Index ??????? TBase
Private MidPriceTBase As Double ' ??????? MidTBase

' Bisection Variables (?????? tb) - v3.0
Private Mintb As Integer        ' Index ????????? tb
Private Maxtb As Integer        ' Index ????????? tb
Private Midtb As Integer        ' Index ??????? tb
Private MidPricetb As Double    ' ??????? Midtb

' Loop Counters
Private Countloop As Long
Private totalcount As Long

' CSV Export
Private csvAcceptData As String
Private csvLoopData As String
Private loopCount As Long
Private bestIterationInRun As Long

' Tracking (?????????? modDataStructures)
Public CostHistory_BA() As Double

'================================================================================
' SECTION 1.5: Helper Functions (**???¡?? Form1**)
'================================================================================

'--------------------------------------------------------------------------------
' Get Concrete Price by f'c (Maha Sarakham Province)
'--------------------------------------------------------------------------------
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

'--------------------------------------------------------------------------------
' Steel Price Constant
'--------------------------------------------------------------------------------
' Public Const STEEL_PRICE_SD40_BA As Double = 24  ' Baht/kg    ' DEAD CODE: orphan Const, no refs outside L98 (commented 20260512-121028)

'================================================================================
' SECTION 2: Initialize Design (Mid values - **????? HCA**)
' v3.0: **????** tb Bisection bounds
'================================================================================

Private Sub InitializeCurrentDesign_BA()
    Dim tb_max_idx As Integer, TBase_max_idx As Integer
    Dim Base_min_idx As Integer, Base_max_idx As Integer
    Dim LToe_min_idx As Integer, LToe_max_idx As Integer
    Dim i As Integer
    
    ' === tb: ?? Min/Max Index ??? constraint tb <= 0.12H ===
    ' v3.0: ??????? Bisection bounds ?????? tb
    tb_max_idx = tb_max
    For i = tb_max To TB_MIN Step -1
        If WP_tb(i) <= 0.12 * modShared.H Then
            tb_max_idx = i
            Exit For
        End If
    Next i
    
    ' ??????? tb Bisection bounds
    Mintb = TB_MIN
    Maxtb = tb_max_idx
    Midtb = (Mintb + Maxtb) / 2
    Currenttb = Maxtb
    
   ' === tt: ???????????? ???? <= tb ===
'Currenttt = (TT_MIN + TT_MAX) / 2

' === tt: ???????? Max ???? <= tb ===
Currenttt = TT_MAX
For i = TT_MAX To TT_MIN Step -1
    If WP_tt(i) <= WP_tb(Currenttb) Then
        Currenttt = i
        Exit For
    End If
Next i

    If Currenttt < TT_MIN Then Currenttt = TT_MIN
    If Currenttt > TT_MAX Then Currenttt = TT_MAX
    ' Constraint: tt <= tb
    If WP_tt(Currenttt) > WP_tb(Currenttb) Then
        For i = TT_MAX To TT_MIN Step -1
            If WP_tt(i) <= WP_tb(Currenttb) Then
                Currenttt = i
                Exit For
            End If
        Next i
    End If
    
    ' === TBase: ?? Max Index ??? constraint TBase <= 0.15H ===
    ' v2.0: ??????? Bisection bounds ?????? TBase
    TBase_max_idx = TBase_max
    For i = TBase_max To TBASE_MIN Step -1
        If WP_TBase(i) <= 0.15 * modShared.H Then
            TBase_max_idx = i
            Exit For
        End If
    Next i
    
    ' ??????? TBase Bisection bounds
    MinTBase = TBASE_MIN
    MaxTBase = TBase_max_idx
    MidTBase = (MinTBase + MaxTBase) / 2
    CurrentTBase = MaxTBase
    
    ' === Base: ?? Min/Max Index ??? constraint 0.5H <= Base <= 0.7H ===
    Base_min_idx = BASE_MIN
    For i = BASE_MIN To BASE_MAX
        If WP_Base(i) >= 0.5 * modShared.H Then
            Base_min_idx = i
            Exit For
        End If
    Next i
    
    Base_max_idx = BASE_MAX
    For i = BASE_MAX To BASE_MIN Step -1
        If WP_Base(i) <= 0.7 * modShared.H Then
            Base_max_idx = i
            Exit For
        End If
    Next i
    
    ' ??????? Base Bisection bounds
    MinBase = Base_min_idx
    MaxBase = Base_max_idx
    MidBase = (MinBase + MaxBase) / 2
    CurrentBase = MaxBase
    
    ' === LToe: ?? Min/Max Index ??? constraint 0.1H <= LToe <= 0.2H ===
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
    
    ' ???????????????
    CurrentLToe = LToe_max_idx
    
    ' === ?????: ??????????????? ===
    CurrentStemDB = DB_MAX
    CurrentStemSP = SP_MIN
    CurrentToeDB = DB_MAX
    CurrentToeSP = SP_MIN
    CurrentHeelDB = DB_MAX
    CurrentHeelSP = SP_MIN
    
End Sub

'================================================================================
' SECTION 3: Get Design from Current Indices
'================================================================================

Private Function GetDesignFromCurrent_BA() As Design
    Dim d As Design
    
    d.tt = WP_tt(Currenttt)
    d.tb = WP_tb(Currenttb)
    d.TBase = WP_TBase(CurrentTBase)
    d.Base = WP_Base(CurrentBase)
    d.LToe = WP_LToe(CurrentLToe)
    d.LHeel = d.Base - d.LToe - d.tb
    
    ' Steel indices
    d.ASst_DB = CurrentStemDB
    d.ASst_Sp = CurrentStemSP
    d.AStoe_DB = CurrentToeDB
    d.AStoe_Sp = CurrentToeSP
    d.ASheel_DB = CurrentHeelDB
    d.ASheel_Sp = CurrentHeelSP
    
    GetDesignFromCurrent_BA = d
End Function

'================================================================================
' SECTION 4: Generate Neighbor (????????????? - HCA Style)
' v3.0: tb ??????? [Mintb, Maxtb] ??? tt <= tb
'================================================================================

Private Sub GenerateNeighbor_BA(ByRef Newtt As Integer, ByRef Newtb As Integer, _
                                ByRef NewTBase As Integer, ByRef NewBase As Integer, _
                                ByRef NewLToe As Integer, _
                                ByRef NewStemDB As Integer, ByRef NewStemSP As Integer, _
                                ByRef NewToeDB As Integer, ByRef NewToeSP As Integer, _
                                ByRef NewHeelDB As Integer, ByRef NewHeelSP As Integer)

    Dim Step As Integer, move(0 To 11) As Integer
    Dim LToe_min_idx As Integer, LToe_max_idx As Integer
    Dim i As Integer

    Call DrawSearchMove(move)

    ' === ????? LToe constraint indices ===
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

    ' === tb: step = Rand(-2, 2) ?????????? [Mintb, Maxtb] ===
    ' v3.0: ??? Bisection bounds
    Step = move(2)
    Newtb = Currenttb + Step
    If Newtb < Mintb Then Newtb = Mintb
    If Newtb > Maxtb Then Newtb = Maxtb

    ' === tt: step = Rand(-2, 2) ??????? <= tb ===
    ' v3.0: Constraint tt <= tb
    Step = move(1)
    Newtt = Currenttt + Step
    If Newtt < TT_MIN Then Newtt = TT_MIN
    If Newtt > TT_MAX Then Newtt = TT_MAX

    ' Clamp tt ??? <= tb (????????!)
    If WP_tt(Newtt) > WP_tb(Newtb) Then
        ' ?? tt index ??????????????? <= tb
        For i = TT_MAX To TT_MIN Step -1
            If WP_tt(i) <= WP_tb(Newtb) Then
                Newtt = i
                Exit For
            End If
        Next i
    End If

    ' === TBase: step = Rand(-5, 5) ?????????? [MinTBase, MaxTBase] ===
    ' v2.0: ??? Bisection bounds
    Step = move(3)
    NewTBase = CurrentTBase + Step
    If NewTBase < MinTBase Then NewTBase = MinTBase
    If NewTBase > MaxTBase Then NewTBase = MaxTBase

    ' === LToe: step = Rand(-2, 2) ===
    Step = move(5)
    NewLToe = CurrentLToe + Step
    If NewLToe < LToe_min_idx Then NewLToe = LToe_min_idx
    If NewLToe > LToe_max_idx Then NewLToe = LToe_max_idx

    ' === Base: step = Rand(-1, 1) ?????????? [MinBase, MaxBase] ===
    Step = move(4)
    NewBase = CurrentBase + Step
    If NewBase < MinBase Then NewBase = MinBase
    If NewBase > MaxBase Then NewBase = MaxBase

    ' === ?????: uniform DB/SP pairs ===

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
' SECTION 5: Main Bisection Optimization Function
' v3.1: Triple Bisection + Fixed CSV Export
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
                       Optional RandomSeed As Long = 12345, _
                       Optional TrialNumber As Long = 1) As Design

    Dim current As Design, neighbor As Design
    Dim currentCost As Double, neighborCost As Double, currentValid As Boolean, ok As Boolean
    Dim saved(1 To 11) As Integer
    Dim Newtt As Integer, Newtb As Integer, NewTBase As Integer, NewBase As Integer, NewLToe As Integer
    Dim NewStemDB As Integer, NewStemSP As Integer, NewToeDB As Integer, NewToeSP As Integer
    Dim NewHeelDB As Integer, NewHeelSP As Integer, i As Long, j As Integer
    Dim rootMintb As Integer, rootMaxtb As Integer, rootMinTBase As Integer, rootMaxTBase As Integer
    Dim rootMinBase As Integer, rootMaxBase As Integer, innerIterations As Long, roundValid As Boolean
    modShared.H = wall_height: modShared.H1 = backfill_height
    modShared.gamma_soil = soil_gamma: modShared.gamma_concrete = concrete_gamma
    modShared.phi = friction_angle: modShared.mu = friction_coef
    modShared.qa = allowable_bearing: modShared.cover = concrete_cover
    modShared.currentMaterial = material
    modShared.currentWSD = CalculateWSDParameters(material.fy, material.fc)
    Call InitializeArrays
    ReDim CostHistory_BA(1 To MaxIterations)
    Call BeginSearch(MaxIterations, RandomSeed, "BA", TrialNumber)
    Call InitializeCurrentDesign_BA
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
    rootMintb = Mintb: rootMaxtb = Maxtb
    rootMinTBase = MinTBase: rootMaxTBase = MaxTBase
    rootMinBase = MinBase: rootMaxBase = MaxBase
    MidPricetb = NO_SOLUTION_COST: MidPriceTBase = NO_SOLUTION_COST: MidPrice = NO_SOLUTION_COST
    Countloop = 0
    current = GetDesignFromCurrent_BA()
    currentValid = EvaluateCandidate(current, "initial", currentCost)
    If Not currentValid Then currentCost = NO_SOLUTION_COST
    CostHistory_BA(EvaluationCount) = RunBestCost
    Do While EvaluationCount < MaxIterations
        Countloop = Countloop + 1
        If Countloop > 1 Then
            Currenttb = Midtb: CurrentTBase = MidTBase: CurrentBase = MidBase
            If WP_tt(Currenttt) > WP_tb(Currenttb) Then
                For j = TT_MAX To TT_MIN Step -1
                    If WP_tt(j) <= WP_tb(Currenttb) Then Currenttt = j: Exit For
                Next j
            End If
            current = GetDesignFromCurrent_BA()
            currentValid = EvaluateCandidate(current, "reset", currentCost)
            If Not currentValid Then currentCost = NO_SOLUTION_COST
            CostHistory_BA(EvaluationCount) = RunBestCost
        End If
        roundValid = currentValid
        innerIterations = 20 * Countloop
        For i = 1 To innerIterations
            If EvaluationCount >= MaxIterations Then Exit For
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
        Call GenerateNeighbor_BA(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP)
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
        neighbor = GetDesignFromCurrent_BA()
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
        DoEvents
        If ok Then roundValid = True
        Next i
        ' Preserve the original THREE-variable heuristic, not a monotone proof.
        ' Invalid rounds supply no evidence for removing any part of the domain.
        If Not roundValid Or (Mintb = Maxtb And MinTBase = MaxTBase And MinBase = MaxBase) Then
            Mintb = rootMintb: Maxtb = rootMaxtb
            MinTBase = rootMinTBase: MaxTBase = rootMaxTBase
            MinBase = rootMinBase: MaxBase = rootMaxBase
            MidPricetb = NO_SOLUTION_COST: MidPriceTBase = NO_SOLUTION_COST: MidPrice = NO_SOLUTION_COST
            ' Reopen bounds without a full-domain random jump.
            ' Resume from current; all random moves use GenerateNeighbor_BA.
            Midtb = Currenttb: MidTBase = CurrentTBase: MidBase = CurrentBase
            RunRecoveryCount = RunRecoveryCount + 1
        Else
            If currentCost < MidPricetb Then
                Maxtb = Currenttb: MidPricetb = currentCost
            Else
                Mintb = Currenttb
            End If
            If currentCost < MidPriceTBase Then
                MaxTBase = CurrentTBase: MidPriceTBase = currentCost
            Else
                MinTBase = CurrentTBase
            End If
            If currentCost < MidPrice Then
                MaxBase = CurrentBase: MidPrice = currentCost
            Else
                MinBase = CurrentBase
            End If
            Midtb = (Mintb + Maxtb) \ 2
            MidTBase = (MinTBase + MaxTBase) \ 2
            MidBase = (MinBase + MaxBase) \ 2
        End If
    Loop
    Call FinishSearch
    BisectionOptimization = RunBest
End Function

'================================================================================
' SECTION 6: CSV Export Functions (?????? HCA)
' v3.1: ???????????????????? HCA
'================================================================================

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
    LastAcceptCSVPath = WriteExportCSV("accept-BA-H" & Replace$(CStr(wallHeight), ",", "."), csvAcceptData)
End Sub

Public Sub SaveLoopPriceCSV_BA(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-BA-H" & Replace$(CStr(wallHeight), ",", "."), csvAcceptData, True)
    LastLoopCSVPath = WriteExportCSV("loopPrice-BA-H" & Replace$(CStr(wallHeight), ",", "."), csvLoopData, True)
End Sub

'================================================================================
' END OF MODULE: modBA.bas v3.1 - Triple Bisection + Fixed CSV Export
'================================================================================
