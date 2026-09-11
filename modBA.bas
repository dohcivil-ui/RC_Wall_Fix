Attribute VB_Name = "modBA"
' BA: initial three-axis midpoint neighbors, then local feasibility recovery.
' After feasibility, alternate midpoint/local proposals while ranges are wide.
' A bounded feasible archive and adjacent steel transitions preserve RunBest.
' Every proposed design uses the shared validator and evaluation budget.
Option Explicit
Public ConceptUpper(1 To 3) As Integer
Public ConceptProposed(1 To 3) As Integer
Public ProposalMidpoint As Boolean
Private ProposalGroup As Integer
Private CoupledLocalCount As Long
Private SeenProposalIndices As Object
Private ProposalRedraw As Boolean, SamplingForEvaluation As Boolean
Public ProposalDraws As Integer, TotalProposalDraws As Long
Public ProposalActiveBridge As Boolean
Public ProposalFrontierEntry As Boolean
Public ProposalFixedBaseCoupled As Boolean
Private BridgeDesign As Design, BridgeOrigin As Design
Private BridgePending As Boolean, BridgeInUse As Boolean, BridgeGroup As Integer
Private BridgeNext As Boolean, BridgeRemaining As Integer
Private BridgeHoldDB As Boolean
Private Const BRIDGE_STEEL As Integer = 1
Private Const BRIDGE_EXPANSION As Integer = 3
Private BridgeKind As Integer
Public ProposalExpansionGroup As Integer
Private ProposalDBStep As Integer, ProposalSPStep As Integer
Private PairDB() As Integer, PairSP() As Integer, PairWeight() As Double
Private PairCount As Integer
Private Const BRANCH_CAPACITY As Integer = 4
Private Const BRANCH_PERIOD As Integer = 8
Private BranchDesign(1 To BRANCH_CAPACITY) As Design
Private BranchIndices(1 To 11, 1 To BRANCH_CAPACITY) As Integer
Private BranchCount As Integer, BranchProposals As Long

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

' Original catalogue bounds remain available to local BA proposals.
Private FixedTbMax As Integer, FixedTBaseMax As Integer
Private FixedBaseMin As Integer, FixedBaseMax As Integer

' Cost History ?????? CostHistory_BA

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
Public SelectedTrialCSVPath_BA As String

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

Private Sub DrawNeighbor(ByRef Newtt As Integer, ByRef Newtb As Integer, _
                                ByRef NewTBase As Integer, ByRef NewBase As Integer, _
                                ByRef NewLToe As Integer, _
                                ByRef NewStemDB As Integer, ByRef NewStemSP As Integer, _
                                ByRef NewToeDB As Integer, ByRef NewToeSP As Integer, _
                                ByRef NewHeelDB As Integer, ByRef NewHeelSP As Integer)

    ' Begin with shared HCA-sized moves, then apply the BA phase policy.
    Dim Step As Integer, move(0 To 11) As Integer
    Dim coupledPeriod As Integer
    Dim LToe_min_idx As Integer, LToe_max_idx As Integer
    Dim i As Integer

    coupledPeriod = 8
    If ConceptUpper(1) - TB_MIN > 4 Or ConceptUpper(2) - TBASE_MIN > 10 Or ConceptUpper(3) - BASE_MIN > 2 Then coupledPeriod = 4
    Call DrawSearchMove(move)
    ' After feasibility, refine one related member group per proposal.
    If Not modFeasibilityRecovery.RecoveryActive Then move(0) = Rand(1, 5)
    If Not modFeasibilityRecovery.RecoveryActive And Not BridgeInUse Then
        If Not ProposalRedraw Then CoupledLocalCount = CoupledLocalCount + 1
        If CoupledLocalCount Mod coupledPeriod = 0 Then move(0) = 0
    End If
    If BridgeInUse Then move(0) = BridgeGroup
    ProposalGroup = move(0)
    ProposalExpansionGroup = 0
    If BridgeInUse And BridgeKind = BRIDGE_EXPANSION Then ProposalExpansionGroup = BridgeGroup
    If BridgeInUse Then
        If ProposalExpansionGroup > 0 Then
            move(2 * BridgeGroup) = 0: move(2 * BridgeGroup + 1) = 1
        Else
            move(2 * BridgeGroup) = -1
            If BridgeHoldDB Then move(2 * BridgeGroup) = 0
            move(2 * BridgeGroup + 1) = -1
        End If
    End If
    If ProposalGroup >= 3 And ProposalGroup <= 5 Then
        ProposalDBStep = move(2 * ProposalGroup)
        ProposalSPStep = move(2 * ProposalGroup + 1)
    End If

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

    ' Three-axis midpoint proposals start from the first neighbor.
    ' Override all three axes AFTER the HCA move-group selection.
    ProposalFixedBaseCoupled = False
    If Not modFeasibilityRecovery.RecoveryActive And Not BridgeInUse Then
        If CoupledLocalCount Mod (2 * coupledPeriod) = 0 Then
            NewTBase = CurrentTBase: NewBase = CurrentBase
            ProposalFixedBaseCoupled = True
        End If
    End If
    ' Bisection neighbors first; alternate with incumbent HCA neighbors while wide.
    ' Once all ranges are narrow, keep both local directions open.
    ProposalMidpoint = modFeasibilityRecovery.RecoveryActive And (EvaluationCount = 1)
    If Not modFeasibilityRecovery.RecoveryActive Then
        If ConceptUpper(1) - TB_MIN > 4 Or ConceptUpper(2) - TBASE_MIN > 10 Or ConceptUpper(3) - BASE_MIN > 2 Then
            If EvaluationCount - RunBestEvaluation < 32 Then
                ProposalMidpoint = (EvaluationCount Mod 2 = 0)
            Else
                ProposalMidpoint = (EvaluationCount Mod 8 = 0)
            End If
        End If
    End If
    If BridgeInUse Then ProposalMidpoint = False
    If ProposalMidpoint Then
        ProposalFixedBaseCoupled = False
        Newtb = ConceptNeighborIndex(TB_MIN, ConceptUpper(1), 2)
        NewTBase = ConceptNeighborIndex(TBASE_MIN, ConceptUpper(2), 5)
        NewBase = ConceptNeighborIndex(BASE_MIN, ConceptUpper(3), 1)
    End If
    ConceptProposed(1) = Newtb: ConceptProposed(2) = NewTBase: ConceptProposed(3) = NewBase
    If WP_tt(Newtt) > WP_tb(Newtb) Then
        For i = TT_MAX To TT_MIN Step -1
            If WP_tt(i) <= WP_tb(Newtb) Then Newtt = i: Exit For
        Next i
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
    ReDim CostHistory_BA(1 To MaxIterations)
    Call BeginSearch(MaxIterations, "BA", TrialNumber)
    CoupledLocalCount = 0
    Set SeenProposalIndices = CreateObject("Scripting.Dictionary")
    ProposalRedraw = False: SamplingForEvaluation = False
    ProposalDraws = 0: TotalProposalDraws = 0: ProposalActiveBridge = False
    ProposalFrontierEntry = False
    ProposalFixedBaseCoupled = False
    BridgePending = False: BridgeInUse = False: BridgeGroup = 0
    BridgeNext = False: BridgeRemaining = 0: BridgeHoldDB = False
    BridgeKind = 0: ProposalExpansionGroup = 0
    Call InitializeCurrentDesign
    Call ResetBranches
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
    SeenProposalIndices.Item(Join(Array(Currenttt, Currenttb, CurrentTBase, CurrentBase, CurrentLToe, CurrentStemDB, CurrentStemSP, CurrentToeDB, CurrentToeSP, CurrentHeelDB, CurrentHeelSP), ":")) = currentValid
    If Not currentValid Then currentCost = NO_SOLUTION_COST
    Call modFeasibilityRecovery.InitializeRecovery(current, currentValid)
    If currentValid Then Call RememberBranch(current, currentCost)
    CostHistory_BA(EvaluationCount) = RunBestCost
    ConceptUpper(1) = Currenttb: ConceptUpper(2) = CurrentTBase: ConceptUpper(3) = CurrentBase
    Do While EvaluationCount < MaxIterations
        If currentValid Then Call SelectBridgeParent(current, currentCost)
        If currentValid And Not BridgeInUse Then Call SelectBranch(current, currentCost)
        If currentValid Then
            ConceptUpper(1) = Currenttb: ConceptUpper(2) = CurrentTBase: ConceptUpper(3) = CurrentBase
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
        SamplingForEvaluation = True
        Call GenerateNeighbor(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP)
        SamplingForEvaluation = False
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
        If ok Then Call RememberBranch(neighbor, neighborCost)
        TotalProposalDraws = TotalProposalDraws + ProposalDraws
        SeenProposalIndices.Item(Join(Array(Currenttt, Currenttb, CurrentTBase, CurrentBase, CurrentLToe, CurrentStemDB, CurrentStemSP, CurrentToeDB, CurrentToeSP, CurrentHeelDB, CurrentHeelSP), ":")) = ok
        If currentValid Then Call RememberBridge(neighbor, neighborCost, currentCost, ok, current)
        If ok And (Not currentValid Or neighborCost < currentCost) Then
            current = neighbor: currentCost = neighborCost: currentValid = True
            modFeasibilityRecovery.RecoveryActive = False
            ConceptUpper(1) = Currenttb: ConceptUpper(2) = CurrentTBase: ConceptUpper(3) = CurrentBase
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
        CostHistory_BA(EvaluationCount) = RunBestCost
        DoEvents
    Loop
    Call FinishSearch
    BisectionOptimization = RunBest
End Function


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
    bestAcceptData = ""
    bestAcceptTrial = 0
    bestAcceptIteration = 0
    bestAcceptCost = NO_SOLUTION_COST
    SelectedTrialCSVPath_BA = ""
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

Public Sub SaveAcceptCSV_BA(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-BA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csvAcceptData, True)
End Sub

Public Sub SaveLoopPriceCSV_BA(wallHeight As Double)
    LastLoopCSVPath = WriteExportCSV("loopPrice-BA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csvLoopData, True)
    Call SaveSelectedTrialCSV_BA(wallHeight)
End Sub

Private Sub SaveSelectedTrialCSV_BA(wallHeight As Double)
    Dim caseSuffix As String, selectedData As String
    caseSuffix = "BA-H" & Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc)
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
    SelectedTrialCSVPath_BA = WriteExportCSV("selectedTrial-" & caseSuffix, selectedData, True)
End Sub


' Original DB/SP catalogues and adjacent index bounds are unchanged.

' Experiment only: midpoint is a sampling center, never a separate evaluation.
Public Function ConceptNeighborIndex(ByVal lower As Integer, ByVal upper As Integer, ByVal radius As Integer, Optional ByVal ticket As Integer = 0) As Integer
    Dim center As Integer, offset As Integer, candidate As Integer
    Dim choices(1 To 20) As Integer, count As Integer
    If upper < lower Then Err.Raise 5, , "Reversed concept interval"
    If upper = lower Then ConceptNeighborIndex = lower: Exit Function
    ' All three original catalogues are evenly spaced. A midpoint tie rounds up.
    center = (lower + upper + 1) \ 2
    For offset = -radius To radius
        If offset <> 0 Then
            candidate = center + offset
            If candidate >= lower And candidate <= upper Then
                count = count + 1: choices(count) = candidate
            End If
        End If
    Next offset
    If count = 0 Then Err.Raise 5, , "No concept neighbor inside interval"
    If ticket = 0 Then ticket = Rand(1, count)
    If ticket < 1 Or ticket > count Then Err.Raise 5, , "Invalid concept ticket"
    ConceptNeighborIndex = choices(ticket)
End Function

Public Sub ConceptFixtureChecks()
    Dim expected(1 To 4) As Integer, i As Integer
    expected(1) = 26: expected(2) = 27: expected(3) = 29: expected(4) = 30
    For i = 1 To 4
        If ConceptNeighborIndex(TB_MIN, tb_max, 2, i) <> expected(i) Then Err.Raise 5, , "First-center tb fixture"
    Next i
    expected(1) = 21: expected(2) = 22: expected(3) = 24: expected(4) = 25
    For i = 1 To 4
        If ConceptNeighborIndex(TB_MIN, 26, 2, i) <> expected(i) Then Err.Raise 5, , "Accepted .50 recenter fixture"
    Next i
    If ConceptNeighborIndex(TB_MIN, TB_MIN + 1, 2, 1) <> TB_MIN Then Err.Raise 5, , "Minimum unreachable"
    If ConceptNeighborIndex(TB_MIN, TB_MIN, 2, 1) <> TB_MIN Then Err.Raise 5, , "Collapsed interval"
End Sub

Public Sub AssertSteelNeighborhood(ByVal dbIndex As Integer, ByVal spIndex As Integer)
    Dim nt As Integer, nb As Integer, nf As Integer, nw As Integer, nl As Integer
    Dim sd As Integer, ss As Integer, td As Integer, ts As Integer, hd As Integer, hs As Integer
    Dim k As Integer, seenLower As Boolean, seenUpper As Boolean, seenSPDown As Boolean, seenSPUp As Boolean
    Call InitializeCurrentDesign
    CurrentStemDB = dbIndex: CurrentToeDB = dbIndex: CurrentHeelDB = dbIndex
    CurrentStemSP = spIndex: CurrentToeSP = spIndex: CurrentHeelSP = spIndex
    ConceptUpper(1) = tb_max: ConceptUpper(2) = TBase_max: ConceptUpper(3) = BASE_MAX
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

Private Sub RestoreFeasibleState(d As Design, ByRef current As Design, ByRef cost As Double)
    Dim i As Integer
    For i = TT_MIN To TT_MAX
        If WP_tt(i) = d.tt Then Currenttt = i: Exit For
    Next i
    For i = TB_MIN To tb_max
        If WP_tb(i) = d.tb Then Currenttb = i: Exit For
    Next i
    For i = TBASE_MIN To TBase_max
        If WP_TBase(i) = d.TBase Then CurrentTBase = i: Exit For
    Next i
    For i = BASE_MIN To BASE_MAX
        If WP_Base(i) = d.Base Then CurrentBase = i: Exit For
    Next i
    For i = LTOE_MIN To LTOE_MAX
        If WP_LToe(i) = d.LToe Then CurrentLToe = i: Exit For
    Next i
    CurrentStemDB = d.ASst_DB: CurrentStemSP = d.ASst_Sp
    CurrentToeDB = d.AStoe_DB: CurrentToeSP = d.AStoe_Sp
    CurrentHeelDB = d.ASheel_DB: CurrentHeelSP = d.ASheel_Sp
    current = d: cost = d.TotalCost
End Sub

Private Sub SelectBridgeParent(ByRef current As Design, ByRef cost As Double)
    If BridgeInUse Then
        If BridgeNext Then
            Call RestoreFeasibleState(BridgeDesign, current, cost)
            BridgeNext = False
            Exit Sub
        End If
        If cost >= BridgeOrigin.TotalCost Then Call RestoreFeasibleState(BridgeOrigin, current, cost)
        BridgeInUse = False: BridgeKind = 0
    End If
    If BridgePending Then
        If BridgeKind <> BRIDGE_EXPANSION Then BridgeOrigin = current
        Call RestoreFeasibleState(BridgeDesign, current, cost)
        BridgeInUse = True: BridgePending = False: BridgeRemaining = 3
    End If
End Sub

Private Sub RememberBridge(d As Design, ByVal cost As Double, ByVal parentCost As Double, ByVal valid As Boolean, parent As Design)
    Dim db As Integer, spacing As Integer
    If BridgeInUse And BridgeKind = BRIDGE_EXPANSION Then
        BridgeRemaining = BridgeRemaining - 1
        If Not valid Or cost < BridgeOrigin.TotalCost Or BridgeRemaining <= 0 Then Exit Sub
        Select Case BridgeGroup
            Case 3: spacing = d.ASst_Sp
            Case 4: spacing = d.AStoe_Sp
            Case 5: spacing = d.ASheel_Sp
        End Select
        If spacing >= SP_MAX Then Exit Sub
        BridgeDesign = d: BridgeNext = True
        Exit Sub
    End If
    If Not BridgeInUse Then
        If StartExpansionBridge(d, cost, parentCost, valid, parent) Then Exit Sub
    End If
    If BridgeInUse Then
        BridgeRemaining = BridgeRemaining - 1
        If valid And cost < BridgeOrigin.TotalCost Then Exit Sub
        If BridgeRemaining <= 0 Then Exit Sub
        If valid Then
            If cost > BridgeOrigin.TotalCost * 1.1 Then Exit Sub
            Select Case BridgeGroup
                Case 3: db = d.ASst_DB
                Case 4: db = d.AStoe_DB
                Case 5: db = d.ASheel_DB
            End Select
            If db <= DB_MIN Then Exit Sub
            BridgeDesign = d: BridgeNext = True: BridgeHoldDB = False
        ElseIf ProposalDBStep = -1 Then
            ' Keep the last feasible parent; densify it once before retrying DB-1.
            BridgeDesign = parent: BridgeNext = True: BridgeHoldDB = True
        End If
        Exit Sub
    End If
    If Not valid Or ProposalMidpoint Or ProposalGroup < 3 Or ProposalGroup > 5 Then Exit Sub
    If ProposalDBStep <> 0 Or ProposalSPStep <> -1 Then Exit Sub
    Select Case ProposalGroup
        Case 3: db = d.ASst_DB
        Case 4: db = d.AStoe_DB
        Case 5: db = d.ASheel_DB
    End Select
    ' Further densification cannot reduce cost when the DB catalogue is at minimum.
    If db <= DB_MIN Then Exit Sub
    If cost <= parentCost Or cost > parentCost * 1.1 Then Exit Sub
    BridgeDesign = d: BridgeGroup = ProposalGroup: BridgePending = True: BridgeHoldDB = False
    BridgeKind = BRIDGE_STEEL
End Sub

Private Function StartExpansionBridge(d As Design, ByVal cost As Double, ByVal parentCost As Double, ByVal valid As Boolean, parent As Design) As Boolean
    Dim expected As Design, spacing As Integer
    If Not valid Or ProposalMidpoint Or ProposalGroup < 3 Or ProposalGroup > 5 Then Exit Function
    If d.tt <> parent.tt Or d.tb <> parent.tb Or d.TBase <> parent.TBase Then Exit Function
    If d.Base <> parent.Base Or d.LToe <> parent.LToe Or d.LHeel <> parent.LHeel Then Exit Function
    expected = parent
    Select Case ProposalGroup
        Case 3
            expected.ASst_DB = expected.ASst_DB + 1: expected.ASst_Sp = expected.ASst_Sp + 1
            spacing = d.ASst_Sp
        Case 4
            expected.AStoe_DB = expected.AStoe_DB + 1: expected.AStoe_Sp = expected.AStoe_Sp + 1
            spacing = d.AStoe_Sp
        Case 5
            expected.ASheel_DB = expected.ASheel_DB + 1: expected.ASheel_Sp = expected.ASheel_Sp + 1
            spacing = d.ASheel_Sp
    End Select
    If d.ASst_DB <> expected.ASst_DB Or d.ASst_Sp <> expected.ASst_Sp Then Exit Function
    If d.AStoe_DB <> expected.AStoe_DB Or d.AStoe_Sp <> expected.AStoe_Sp Then Exit Function
    If d.ASheel_DB <> expected.ASheel_DB Or d.ASheel_Sp <> expected.ASheel_Sp Then Exit Function
    If spacing >= SP_MAX Then Exit Function
    If cost <= parentCost Or cost > parentCost * 1.1 Then Exit Function
    ' A larger bar may permit later spacing increases; evaluate every intermediate.
    BridgeOrigin = parent: BridgeDesign = d: BridgeKind = BRIDGE_EXPANSION
    BridgeGroup = ProposalGroup: BridgePending = True: BridgeNext = False: BridgeHoldDB = False
    StartExpansionBridge = True
End Function

Private Sub GenerateNeighbor(ByRef Newtt As Integer, ByRef Newtb As Integer, _
                                ByRef NewTBase As Integer, ByRef NewBase As Integer, _
                                ByRef NewLToe As Integer, _
                                ByRef NewStemDB As Integer, ByRef NewStemSP As Integer, _
                                ByRef NewToeDB As Integer, ByRef NewToeSP As Integer, _
                                ByRef NewHeelDB As Integer, ByRef NewHeelSP As Integer)
    Dim drawIndex As Integer, drawLimit As Integer, signature As String
    ' Standalone sampler fixtures neither publish nor accumulate actual-run draws.
    If Not SamplingForEvaluation Then
        ProposalRedraw = False
        Call DrawNeighbor(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP)
        Exit Sub
    End If
    ProposalActiveBridge = BridgeInUse
    ProposalFrontierEntry = False
    drawLimit = 1
    If Not modFeasibilityRecovery.RecoveryActive And Not ProposalActiveBridge Then drawLimit = 16
    For drawIndex = 1 To drawLimit
        ProposalRedraw = (drawIndex > 1)
        Call DrawNeighbor(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP)
        ProposalDraws = drawIndex
        If drawLimit = 1 Then Exit For
        signature = Join(Array(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP), ":")
        If Not SeenProposalIndices.Exists(signature) Then Exit For
        If HasUnseenExpansionContinuation(signature) Then
            ProposalFrontierEntry = True
            Exit For
        End If
    Next drawIndex
    ProposalRedraw = False
End Sub

Private Function HasUnseenExpansionContinuation(ByVal signature As String) As Boolean
    Dim expected As Variant, chain As Variant, dbPosition As Integer, spPosition As Integer
    Dim stepIndex As Integer, nextSpacing As Integer, nextSignature As String
    If modFeasibilityRecovery.RecoveryActive Or ProposalActiveBridge Or ProposalMidpoint Then Exit Function
    If ProposalGroup < 3 Or ProposalGroup > 5 Then Exit Function
    expected = Array(Currenttt, Currenttb, CurrentTBase, CurrentBase, CurrentLToe, CurrentStemDB, CurrentStemSP, CurrentToeDB, CurrentToeSP, CurrentHeelDB, CurrentHeelSP)
    dbPosition = 2 * ProposalGroup - 1: spPosition = 2 * ProposalGroup
    expected(dbPosition) = expected(dbPosition) + 1
    expected(spPosition) = expected(spPosition) + 1
    If Join(expected, ":") <> signature Then Exit Function
    If Not SeenProposalIndices.Exists(signature) Then Exit Function
    If Not CBool(SeenProposalIndices.Item(signature)) Then Exit Function
    ' Historical booleans guide sampling only; the retained entry is evaluated again.
    chain = Split(signature, ":")
    For stepIndex = 1 To 3
        nextSpacing = CInt(chain(spPosition)) + 1
        If nextSpacing > SP_MAX Then Exit Function
        chain(spPosition) = CStr(nextSpacing)
        nextSignature = Join(chain, ":")
        If Not SeenProposalIndices.Exists(nextSignature) Then
            HasUnseenExpansionContinuation = True
            Exit Function
        End If
        If Not CBool(SeenProposalIndices.Item(nextSignature)) Then Exit Function
    Next stepIndex
End Function

Private Sub ResetBranches()
    BranchCount = 0: BranchProposals = 0
End Sub

Private Sub RememberBranch(d As Design, ByVal cost As Double)
    Dim slot As Integer, i As Integer, worst As Integer
    For i = 1 To BranchCount
        If BranchIndices(2, i) = Currenttb And BranchIndices(3, i) = CurrentTBase And BranchIndices(4, i) = CurrentBase And BranchIndices(6, i) = CurrentStemDB Then
            If cost >= BranchDesign(i).TotalCost Then Exit Sub
            slot = i: Exit For
        End If
    Next i
    If slot = 0 Then
        If BranchCount < BRANCH_CAPACITY Then
            BranchCount = BranchCount + 1: slot = BranchCount
        Else
            worst = 1
            For i = 2 To BranchCount
                If BranchDesign(i).TotalCost > BranchDesign(worst).TotalCost Then worst = i
            Next i
            If cost >= BranchDesign(worst).TotalCost Then Exit Sub
            slot = worst
        End If
    End If
    BranchDesign(slot) = d
    BranchIndices(1, slot) = Currenttt: BranchIndices(2, slot) = Currenttb
    BranchIndices(3, slot) = CurrentTBase: BranchIndices(4, slot) = CurrentBase
    BranchIndices(5, slot) = CurrentLToe
    BranchIndices(6, slot) = CurrentStemDB: BranchIndices(7, slot) = CurrentStemSP
    BranchIndices(8, slot) = CurrentToeDB: BranchIndices(9, slot) = CurrentToeSP
    BranchIndices(10, slot) = CurrentHeelDB: BranchIndices(11, slot) = CurrentHeelSP
End Sub

Private Sub SelectBranch(ByRef d As Design, ByRef cost As Double)
    Dim slot As Integer, i As Integer
    BranchProposals = BranchProposals + 1
    If BranchCount = 0 Or (BranchProposals Mod BRANCH_PERIOD) <> 0 Then Exit Sub
    slot = Rand(1, BranchCount)
    If Rand(1, 4) > 1 Then
        For i = 1 To BranchCount
            If BranchDesign(i).TotalCost < BranchDesign(slot).TotalCost Then slot = i
        Next i
    End If
    d = BranchDesign(slot): cost = d.TotalCost
    Currenttt = BranchIndices(1, slot): Currenttb = BranchIndices(2, slot)
    CurrentTBase = BranchIndices(3, slot): CurrentBase = BranchIndices(4, slot)
    CurrentLToe = BranchIndices(5, slot)
    CurrentStemDB = BranchIndices(6, slot): CurrentStemSP = BranchIndices(7, slot)
    CurrentToeDB = BranchIndices(8, slot): CurrentToeSP = BranchIndices(9, slot)
    CurrentHeelDB = BranchIndices(10, slot): CurrentHeelSP = BranchIndices(11, slot)
End Sub
