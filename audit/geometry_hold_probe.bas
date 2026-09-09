Attribute VB_Name = "ColumnProbe"
Option Explicit
Private beforeCenter As Design, beforeCenterCost As Double, beforeCenterValid As Boolean
Private lastCenter As Design, lastCenterCost As Double, centerPending As Boolean
Private pairDB(1 To 20) As Integer, pairSP(1 To 20) As Integer
Private pairsReady As Boolean, midpointRejected As Long, midpointAccepted As Long
Private checks As Long, mids As Long, proposals As Long, phases As Long, crosses As Long
Private previousEvaluation As Long, expectedMid As Long, currentPhase As Long
Private best As Double, first As Long, hist(1 To 1000) As Double
Private move(0 To 11) As Integer, origin As Design
Private tbHi As Integer, tfHi As Integer, bLo As Integer, bHi As Integer
Private midLo(1 To 3) As Integer, midHi(1 To 3) As Integer
Private runMids As Long
Public ForcingTicket As Boolean
Public ForcingGeometry As Boolean
Private geometryTickets(1 To 5) As Integer, geometryConsumed As Integer
Private geometryVectors As Long, geometryForcedDraws As Long
Private geometryTicketCounts(1 To 5, 1 To 6) As Long
Private rawGeometryTickets(1 To 5) As Long
Private forcedValue As Long, forcedHigh As Long, consumed As Boolean
Private ticketsChecked As Long, originsChecked As Long
Private randCalls As Long, pairCalls As Long, totalPairCalls As Long, totalGeometryCalls As Long
Private proposalActive As Boolean, pairHigh As Long
Private pendingNeighbor As Boolean, lastNeighbor As Design, lastNeighborCost As Double

Public Sub Check(ok As Boolean, message As String)
    checks = checks + 1
    If Not ok Then Err.Raise 5, , message
End Sub

Public Sub RememberMove(values() As Integer)
    Dim i As Integer
    Check proposalActive And randCalls = 5 And pairCalls = 0, "Five geometry draws must precede all steel pairs"
    For i = 0 To 11: move(i) = values(i): Next i
    For i = 1 To 5
        Check move(GeometrySlot(i)) = OffsetFromTicket(rawGeometryTickets(i)), "Real geometry output differs from its own raw ticket"
    Next i
    For i = 1 To 11: Check Abs(move(i)) <= 1, "Sample is not a local one-level offset": Next i
End Sub

Private Function Idx(values() As Double, value As Double) As Integer
    Dim i As Integer
    For i = LBound(values) To UBound(values)
        If values(i) = value Then Idx = i: Exit Function
    Next i
    Err.Raise 5, , "Design value not in unchanged table"
End Function

Private Function Clamp(value As Integer, low As Integer, high As Integer) As Integer
    Clamp = value
    If value < low Then Clamp = low
    If value > high Then Clamp = high
End Function

Public Sub RememberOrigin(d As Design)
    origin = d
    randCalls = 0: pairCalls = 0: proposalActive = True
End Sub

Public Sub RememberBounds(lo() As Integer, hi() As Integer, th As Integer, tf As Integer, bl As Integer, bh As Integer)
    Dim i As Integer
    For i = 1 To 3: midLo(i) = lo(i): midHi(i) = hi(i): Next i
    tbHi = th: tfHi = tf: bLo = bl: bHi = bh
End Sub

Private Sub CheckCenterValue(values() As Double, value As Double, low As Integer, high As Integer)
    Dim center As Double, i As Integer, distance As Double
    center = (values(low) + values(high)) / 2
    distance = Abs(value - center)
    Check Idx(values, value) >= low And Idx(values, value) <= high, "Center outside endpoint window"
    For i = low To high
        Check distance <= Abs(values(i) - center) + 0.000000000001, "Not nearest physical midpoint"
        If Abs(Abs(values(i) - center) - distance) < 0.000000000001 Then Check value <= values(i), "Midpoint tie rule differs"
    Next i
End Sub

Public Sub CheckCenter(d As Design)
    CheckCenterValue WP_tb, d.tb, midLo(1), midHi(1)
    CheckCenterValue WP_TBase, d.TBase, midLo(2), midHi(2)
    CheckCenterValue WP_Base, d.Base, midLo(3), midHi(3)
    Check d.tt <= d.tb, "Center tt repair missing"
End Sub

Public Sub CheckProposal(d As Design)
    Dim li As Integer, lh As Integer, i As Integer, nti As Integer
    Check proposalActive And randCalls = 8 And pairCalls = 3, "Full proposal requires five geometry and three independent pair draws"
    proposalActive = False
    proposals = proposals + 1
    Check d.tb = WP_tb(Clamp(Idx(WP_tb, origin.tb) + move(2), TB_MIN, tbHi)), "tb not full-domain HCA draw"
    Check d.TBase = WP_TBase(Clamp(Idx(WP_TBase, origin.TBase) + move(3), TBASE_MIN, tfHi)), "TBase not full-domain HCA draw"
    Check d.Base = WP_Base(Clamp(Idx(WP_Base, origin.Base) + move(4), bLo, bHi)), "Base not full-domain HCA draw"
    nti = Clamp(Idx(WP_tt, origin.tt) + move(1), TT_MIN, TT_MAX)
    Check d.tt = WP_tt(nti), "tt draw/repair differs"
    li = LTOE_MIN: lh = LTOE_MAX
    For i = LTOE_MIN To LTOE_MAX
        If WP_LToe(i) >= 0.1 * H Then li = i: Exit For
    Next i
    For i = LTOE_MAX To LTOE_MIN Step -1
        If WP_LToe(i) <= 0.2 * H Then lh = i: Exit For
    Next i
    Check d.LToe = WP_LToe(Clamp(Idx(WP_LToe, origin.LToe) + move(5), li, lh)), "Toe not sampled with remaining vector"
    If RunAlgorithm = "BA" Then
        CheckPair origin.ASst_DB, origin.ASst_Sp, d.ASst_DB, d.ASst_Sp
        CheckPair origin.AStoe_DB, origin.AStoe_Sp, d.AStoe_DB, d.AStoe_Sp
        CheckPair origin.ASheel_DB, origin.ASheel_Sp, d.ASheel_DB, d.ASheel_Sp
    Else
    Check d.ASst_DB = Clamp(origin.ASst_DB + move(6), DB_MIN, DB_MAX) And d.ASst_Sp = Clamp(origin.ASst_Sp + move(7), SP_MIN, SP_MAX), "Stem bar sample ignored"
    Check d.AStoe_DB = Clamp(origin.AStoe_DB + move(8), DB_MIN, DB_MAX) And d.AStoe_Sp = Clamp(origin.AStoe_Sp + move(9), SP_MIN, SP_MAX), "Toe bar sample ignored"
    Check d.ASheel_DB = Clamp(origin.ASheel_DB + move(10), DB_MIN, DB_MAX) And d.ASheel_Sp = Clamp(origin.ASheel_Sp + move(11), SP_MIN, SP_MAX), "Heel bar sample ignored"
    End If
    If d.tb < WP_tb(midLo(1)) Or d.tb > WP_tb(midHi(1)) Or d.TBase < WP_TBase(midLo(2)) Or d.TBase > WP_TBase(midHi(2)) Or d.Base < WP_Base(midLo(3)) Or d.Base > WP_Base(midHi(3)) Then crosses = crosses + 1
End Sub

Public Sub CheckState(d As Design, cost As Double, valid As Boolean, indexed As Design)
    If pendingNeighbor Then
        If lastNeighbor.IsValid And (Not origin.IsValid Or lastNeighborCost < origin.TotalCost) Then
            Check SameVector(d, lastNeighbor) And valid And cost = lastNeighborCost, "Admissible neighbor not adopted"
        Else
            Check SameVector(d, origin) And valid = origin.IsValid And cost = origin.TotalCost, "Rejected neighbor changed incumbent"
        End If
        pendingNeighbor = False
    End If
    If centerPending Then
        If lastCenter.IsValid And (Not beforeCenterValid Or lastCenterCost < beforeCenterCost) Then
            midpointAccepted = midpointAccepted + 1
            Check SameVector(d, lastCenter) And valid And cost = lastCenterCost, "Admissible midpoint not adopted"
        Else
            midpointRejected = midpointRejected + 1
            Check SameVector(d, beforeCenter) And valid = beforeCenterValid And cost = beforeCenterCost, "Rejected midpoint changed incumbent"
        End If
        centerPending = False
    End If
    Check cost = d.TotalCost And valid = d.IsValid, "Unchecked or stale local price/validity"
    Check d.tb = indexed.tb And d.tt = indexed.tt And d.TBase = indexed.TBase And d.Base = indexed.Base And d.LToe = indexed.LToe, "Rejected geometry was not restored"
    Check d.ASst_DB = indexed.ASst_DB And d.ASst_Sp = indexed.ASst_Sp And d.AStoe_DB = indexed.AStoe_DB And d.AStoe_Sp = indexed.AStoe_Sp And d.ASheel_DB = indexed.ASheel_DB And d.ASheel_Sp = indexed.ASheel_Sp, "Rejected steel state was not restored"
End Sub

Public Sub Trace(d As Design, entry As String, cost As Double)
    Check EvaluationCount = previousEvaluation + 1, "Evaluation skipped or hidden"
    previousEvaluation = EvaluationCount
    If entry = "neighbor" Then
        lastNeighbor = d: lastNeighborCost = cost: pendingNeighbor = True
    End If
    If entry = "midpoint" Then
        lastCenter = d: lastCenterCost = cost: centerPending = True
    End If
    If RunAlgorithm = "BA" And EvaluationCount = expectedMid Then
        Check entry = "midpoint", "Expected simultaneous midpoint relocation"
        currentPhase = currentPhase + 1: runMids = runMids + 1: mids = mids + 1
        expectedMid = expectedMid + 20 * currentPhase + 1
    ElseIf EvaluationCount > 1 Then
        Check entry = "neighbor", "Phase quota differs"
    Else
        Check entry = "initial", "Initial evaluation missing"
    End If
    If d.IsValid Then
        Check d.FS_OT >= FS_OT_MIN And d.FS_SL >= FS_SL_MIN And d.FS_BC >= FS_BC_MIN, "Invalid safety accepted"
        If cost < best Then best = cost: first = EvaluationCount
    End If
    Check RunBestCost = best And RunBestEvaluation = first, "Global best lost or wrong first index"
    hist(EvaluationCount) = best
End Sub

Public Sub Main()
    Dim methodIndex As Integer
    Dim scenario As Integer, budget As Long, bearing As Double, d As Design, mat As MaterialProperties, f As Integer, i As Long
    On Error GoTo Failed
    If RESULT_CSV_ROOT <> App.Path & "\output" Then Err.Raise 5, , "Unsafe output root"
    Randomize
    InitializeArrays
    modBA.CheckGeometryRules
    modBA.CheckPairRules
    EnableProjectChecks
    mat = GetSD40Material(240, GetConcretePrice(240), 24)
    For methodIndex = 1 To 1
    For scenario = 1 To 9
        bearing = 30
        Select Case scenario
            Case 1: budget = 1
            Case 2: budget = 2
            Case 3: budget = 21
            Case 4: budget = 22
            Case 5: budget = 23
            Case 6: budget = 63
            Case 7: budget = 64
            Case 8: budget = 128: bearing = 0.000001
            Case 9: budget = 1000
        End Select
        best = NO_SOLUTION_COST: first = 0: previousEvaluation = 0
        expectedMid = 2: currentPhase = 0: runMids = 0
        If methodIndex = 1 Then
            d = BisectionOptimization(budget, 3, 1.2, 1.8, 2.4, 30, 0.6, bearing, 0.075, mat)
        Else
            d = HillClimbingOptimization(budget, 3, 1.2, 1.8, 2.4, 30, 0.6, bearing, 0.075, mat)
        End If
        Check EvaluationCount = budget, "Budget not exhausted exactly"
        Check BestCostIteration = first, "Final first-best mismatch"
        Check d.IsValid = RunBest.IsValid And d.TotalCost = RunBest.TotalCost, "Wrong returned global best"
        For i = 1 To budget
            If methodIndex = 1 Then
                Check CostHistory_BA(i) = hist(i), "BA history differs from evaluated feasible best"
            Else
                Check modDataStructures.CostHistory(i) = hist(i), "HCA history differs from evaluated feasible best"
            End If
        Next i
        If scenario = 8 Then Check Not d.IsValid, "Impossible bearing passed"
    Next scenario
    Next methodIndex
    modBA.CheckCenterRules
    Check mids > 0 And proposals > 0 And crosses > 0, "Center relocation/full-domain movement not exercised"
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
    Print #f, "CHECKS=" & checks & "; MIDPOINTS=" & mids & "; PROPOSALS=" & proposals & "; CROSSING_PROPOSALS=" & crosses
    Print #f, "MIDPOINT_ACCEPTED=" & midpointAccepted & "; MIDPOINT_REJECTED=" & midpointRejected
    Print #f, "PAIR_ORIGINS=" & originsChecked & "; TICKETS=" & ticketsChecked & "; INDEPENDENT_PAIR_DRAWS=" & totalPairCalls & "; GEOMETRY_DRAWS=" & totalGeometryCalls
    Print #f, "GEOMETRY_TICKET_VECTORS=" & geometryVectors & "; GEOMETRY_FORCED_DRAWS=" & geometryForcedDraws
    Print #f, "FAILURES=0"
    Close #f
    Exit Sub
Failed:
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
    Print #f, "FATAL=" & Err.Number & ": " & Err.Description
    Close #f
End Sub

Public Sub SetGlobalBounds(th As Integer, tf As Integer, bl As Integer, bh As Integer)
    tbHi = th: tfHi = tf: bLo = bl: bHi = bh
End Sub

Public Sub RememberBeforeCenter(d As Design, cost As Double, valid As Boolean)
    beforeCenter = d: beforeCenterCost = cost: beforeCenterValid = valid
End Sub

Private Function SameVector(a As Design, b As Design) As Boolean
    SameVector = a.tt = b.tt And a.tb = b.tb And a.TBase = b.TBase And a.Base = b.Base And a.LToe = b.LToe And _
        a.ASst_DB = b.ASst_DB And a.ASst_Sp = b.ASst_Sp And a.AStoe_DB = b.AStoe_DB And a.AStoe_Sp = b.AStoe_Sp And a.ASheel_DB = b.ASheel_DB And a.ASheel_Sp = b.ASheel_Sp
End Function

Public Sub CheckPair(db As Integer, sp As Integer, newDB As Integer, newSP As Integer)
    Dim oldRank As Integer, newRank As Integer, i As Integer, j As Integer, n As Integer, temp As Integer
    If Not pairsReady Then
        For i = DB_MIN To DB_MAX
            For j = SP_MIN To SP_MAX
                n = n + 1: pairDB(n) = i: pairSP(n) = j
            Next j
        Next i
        For i = 1 To 19
            For j = i + 1 To 20
                If ProvidedSteelWeight(pairDB(j), pairSP(j), 1#) < ProvidedSteelWeight(pairDB(i), pairSP(i), 1#) Then
                    temp = pairDB(i): pairDB(i) = pairDB(j): pairDB(j) = temp
                    temp = pairSP(i): pairSP(i) = pairSP(j): pairSP(j) = temp
                End If
            Next j
        Next i
        pairsReady = True
    End If
    For i = 1 To 20
        If pairDB(i) = db And pairSP(i) = sp Then oldRank = i
        If pairDB(i) = newDB And pairSP(i) = newSP Then newRank = i
    Next i
    Check oldRank > 0 And newRank > 0, "Unknown steel pair"
    Check (Abs(newDB - db) <= 1 And Abs(newSP - sp) <= 1) Or Abs(newRank - oldRank) <= 1, "Pair proposal outside declared union"
End Sub


Public Sub ForceTicket(ticket As Integer, high As Integer)
    ForcingTicket = True: forcedValue = ticket: forcedHigh = high: consumed = False
End Sub

Public Function TakeTicket(low As Long, high As Long) As Long
    Check ForcingTicket And Not consumed, "Ticket must use exactly one native Rand call"
    Check low = 1 And high = forcedHigh, "PairNeighbor ticket bounds violate half-stay contract"
    Check forcedValue >= low And forcedValue <= high, "Forced ticket outside actual Rand bounds"
    consumed = True
    TakeTicket = forcedValue
End Function

Public Sub FinishTicket()
    Check ForcingTicket And consumed, "PairNeighbor did not consume native Rand ticket"
    ForcingTicket = False: ticketsChecked = ticketsChecked + 1
End Sub

Public Sub PairOriginDone()
    originsChecked = originsChecked + 1
End Sub

Public Sub BeginPair(db As Integer, sp As Integer, count As Integer)
    If ForcingTicket Then Exit Sub
    Check proposalActive And randCalls = 5 + pairCalls, "Pair draws were skipped, shared, or interleaved incorrectly"
    pairCalls = pairCalls + 1: totalPairCalls = totalPairCalls + 1
    Check pairCalls <= 3, "More than three member pair draws"
    Select Case pairCalls
        Case 1: Check db = origin.ASst_DB And sp = origin.ASst_Sp, "Stem pair origin changed before proposal"
        Case 2: Check db = origin.AStoe_DB And sp = origin.AStoe_Sp, "Toe pair origin changed before proposal"
        Case 3: Check db = origin.ASheel_DB And sp = origin.ASheel_Sp, "Heel pair origin changed before proposal"
    End Select
    pairHigh = 2 * (count - 1)
End Sub

Public Sub ObserveRand(low As Long, high As Long)
    If Not proposalActive Then Exit Sub
    randCalls = randCalls + 1
    If randCalls <= 5 Then
        Check low = 1 And high = 6 And pairCalls = 0, "Geometry draw ticket bounds violate two-thirds hold contract"
        totalGeometryCalls = totalGeometryCalls + 1
    Else
        Check randCalls = 5 + pairCalls And low = 1 And high = pairHigh, "Pair requires its own half-stay Rand draw"
    End If
End Sub


Private Function GeometrySlot(drawPosition As Integer) As Integer
    Select Case drawPosition
        Case 1: GeometrySlot = 2
        Case 2: GeometrySlot = 1
        Case 3: GeometrySlot = 3
        Case 4: GeometrySlot = 5
        Case 5: GeometrySlot = 4
        Case Else: Err.Raise 5, , "Unknown geometry draw position"
    End Select
End Function

Private Function OffsetFromTicket(ByVal ticket As Long) As Integer
    Check ticket >= 1 And ticket <= 6, "Geometry ticket out of declared range"
    Select Case ticket
        Case 1: OffsetFromTicket = -1
        Case 6: OffsetFromTicket = 1
        Case Else: OffsetFromTicket = 0
    End Select
End Function

Public Sub ForceGeometry(vectorCode As Long)
    Dim i As Integer, remaining As Long
    Check Not ForcingTicket And Not ForcingGeometry, "Overlapping forced streams"
    Check vectorCode >= 0 And vectorCode < 7776, "Geometry vector code outside enumeration"
    remaining = vectorCode
    For i = 1 To 5
        geometryTickets(i) = (remaining Mod 6) + 1
        remaining = remaining \ 6
    Next i
    ForcingGeometry = True: geometryConsumed = 0
End Sub

Public Function TakeGeometryTicket(low As Long, high As Long) As Long
    Check ForcingGeometry And geometryConsumed < 5, "Expected exactly five independent geometry calls"
    Check low = 1 And high = 6, "Geometry draw ticket bounds violate two-thirds hold contract"
    geometryConsumed = geometryConsumed + 1
    geometryForcedDraws = geometryForcedDraws + 1
    TakeGeometryTicket = geometryTickets(geometryConsumed)
End Function

Public Sub FinishGeometry(values() As Integer)
    Dim i As Integer, slot As Integer
    Check ForcingGeometry And geometryConsumed = 5, "Geometry skipped or reused one or more draws"
    For i = 1 To 5
        slot = GeometrySlot(i)
        Check values(slot) = OffsetFromTicket(geometryTickets(i)), "Geometry output does not use its independent ticket"
        geometryTicketCounts(i, geometryTickets(i)) = geometryTicketCounts(i, geometryTickets(i)) + 1
    Next i
    For i = 6 To 11: Check values(i) = 0, "Geometry sampler overwrote a steel move": Next i
    Check values(0) = 0, "Geometry sampler added member grouping"
    geometryVectors = geometryVectors + 1: ForcingGeometry = False
End Sub

Public Sub CheckGeometryCoverage()
    Dim i As Integer, ticket As Integer
    Check geometryVectors = 7776 And geometryForcedDraws = 38880, "Incomplete geometry ticket Cartesian product"
    For i = 1 To 5
        For ticket = 1 To 6
            Check geometryTicketCounts(i, ticket) = 1296, "Not every geometry ticket was covered in every slot"
        Next ticket
    Next i
End Sub

Public Sub ObserveRandValue(value As Long)
    If Not proposalActive Then Exit Sub
    If randCalls <= 5 Then
        Check value >= 1 And value <= 6, "Real geometry ticket out of bounds"
        rawGeometryTickets(randCalls) = value
    End If
End Sub
