Attribute VB_Name = "modFeasibilityMerit"
Option Explicit
' Search heuristic ONLY. Zero merit is NEVER a validity or export decision.
' CheckProjectDesign remains the authoritative, unchanged acceptance gate.
' Demand math mirrors PROJECT_BOOK_WSD_ACI99_V5_MEMBER_LOADS for ranking only.
' Stem stresses use 17 stations, not the validator's certified interval bounds.
Private Const FM_PSI_KSC As Double = 0.0703069579640175
Private Const FM_IN_M As Double = 0.0254
Private Const FM_TOL As Double = 0.000000001

Private Function FMMax(ByVal a As Double, ByVal b As Double) As Double
    FMMax = a: If b > a Then FMMax = b
End Function

Private Sub FMAdd(ByRef total As Double, ByVal residual As Double)
    If residual <= 0 Then Exit Sub
    If residual > 1000000# Then residual = 1000000#
    total = total + residual * residual
End Sub

Private Sub FMSpacing(ByRef total As Double, ByVal bar As Double, ByVal spacing As Double, ByVal thickness As Double, ByVal cc As Double, ByVal stress As Double)
    Dim requiredGap As Double, fksi As Double
    requiredGap = FMMax(bar, 0.0266666666666667)
    Call FMAdd(total, (requiredGap - (spacing - bar)) / requiredGap)
    Call FMAdd(total, spacing / (3# * thickness) - 1#)
    Call FMAdd(total, spacing / (18# * FM_IN_M) - 1#)
    Call FMAdd(total, -cc / 0.075)
    ' Positive-product forms remain meaningful when a computed spacing limit is negative.
    fksi = stress / FM_PSI_KSC / 1000#
    Call FMAdd(total, fksi * (spacing + 2.5 * cc) / FM_IN_M / 540# - 1#)
    Call FMAdd(total, fksi * spacing / FM_IN_M / 432# - 1#)
End Sub

Private Function FMGeometryDistance(d As Design) As Double
    Dim value As Double
    Call FMAdd(value, (0.000001 - d.tt) / 0.1)
    Call FMAdd(value, (d.tt - d.tb) / 0.1)
    Call FMAdd(value, (0.000001 - d.TBase) / 0.1)
    Call FMAdd(value, (d.TBase - H + 0.000001) / FMMax(H, 0.1))
    Call FMAdd(value, (d.TBase - H1) / FMMax(H1, 0.1))
    Call FMAdd(value, (H1 - H) / FMMax(H, 0.1))
    Call FMAdd(value, (0.000001 - d.Base) / 1#)
    Call FMAdd(value, (0.000001 - d.LToe) / 0.3)
    Call FMAdd(value, (0.3 - d.LHeel) / 0.3)
    Call FMAdd(value, (d.LToe + 0.000001 - d.LHeel) / 0.3)
    Call FMAdd(value, Abs(d.Base - d.LToe - d.tb - d.LHeel) / FMMax(d.Base, 0.1))
    Call FMAdd(value, (cover + 0.000001 - d.tb) / 0.1)
    Call FMAdd(value, (cover + 0.000001 - d.TBase) / 0.1)
    FMGeometryDistance = value
End Function

Private Function FMNetStem(ByVal hs As Double, ByVal hp As Double, ByVal y As Double, ByVal ka As Double, ByVal kp As Double) As Double
    FMNetStem = gamma_soil * (ka * FMMax(hs - y, 0) ^ 3 - PassiveFactor * kp * FMMax(hp - y, 0) ^ 3) / 6#
End Function

Public Function FeasibilityMerit(d As Design) As Double
    Dim savedReason As String, merit As Double, geometryDistance As Double
    Dim db(0 To 2) As Integer, sp(0 To 2) As Integer, bar(0 To 2) As Double
    Dim steel(0 To 2) As Double, dep(0 To 2) As Double, revDep(0 To 2) As Double
    Dim moments(0 To 2) As Double, shears(0 To 2) As Double
    Dim reverseM(0 To 2) As Double, reverseV(0 To 2) As Double
    Dim w(0 To 3) As Double, wx(0 To 3) As Double, factor(0 To 3) As Double
    Dim i As Integer, mask As Integer, station As Integer, hs As Double, hp As Double
    Dim ka As Double, kp As Double, pa As Double, pp As Double, net As Double
    Dim weight As Double, mr As Double, qt As Double, qh As Double, slope As Double
    Dim length As Double, wd As Double, qfree As Double, nm As Double, rm As Double, nv As Double, rv As Double
    Dim beta As Double, rhoMax As Double, vLimit As Double, thickness As Double, topDepth As Double
    Dim ca As Double, st As Double, j As Double, peakCa As Double, peakSt As Double
    Dim y As Double, localDepth As Double, m As Double, netMin As Double, rootRatio As Double, cc As Double
    savedReason = LastValidationReason
    On Error GoTo Failed
    geometryDistance = FMGeometryDistance(d)
    FeasibilityMerit = 1000000# + geometryDistance
    If Not GeometryOK(d) Or Not CheckHeelLayout(d) Or d.LToe <= 0 Then GoTo Finish
    If currentWSD.fcPrime <= 0 Or currentWSD.fy <= 0 Or currentWSD.fc <= 0 Or currentWSD.fs <= 0 Or currentWSD.n <= 0 Then GoTo Finish
    If currentWSD.fcPrime <> currentMaterial.fc Or currentWSD.fy <> currentMaterial.fy Then GoTo Finish
    If PassiveFactor <> PROJECT_PASSIVE_FACTOR Then GoTo Finish
    If d.UseDoubleStem Or d.UseDoubleToe Or d.UseDoubleHeel Then GoTo Finish
    db(0) = d.ASst_DB: sp(0) = d.ASst_Sp
    db(1) = d.AStoe_DB: sp(1) = d.AStoe_Sp
    db(2) = d.ASheel_DB: sp(2) = d.ASheel_Sp
    For i = 0 To 2
        If db(i) < DB_MIN Or db(i) > DB_MAX Or sp(i) < SP_MIN Or sp(i) > SP_MAX Then GoTo Finish
        bar(i) = WP_DB(db(i)) / 1000#
        steel(i) = CalculateAsProv(db(i), sp(i))
        thickness = d.TBase: If i = 0 Then thickness = d.tb
        dep(i) = thickness - cover - bar(i) / 2#
        revDep(i) = cover + bar(i) / 2#
        If dep(i) <= 0 Or revDep(i) <= 0 Or steel(i) <= 0 Then
            FeasibilityMerit = FeasibilityMerit + (FMMax(0.000001 - dep(i), 0) / 0.1) ^ 2
            GoTo Finish
        End If
    Next i
    topDepth = d.tt - cover - bar(0) / 2#
    If topDepth <= 0 Then
        FeasibilityMerit = FeasibilityMerit + ((0.000001 - topDepth) / 0.1) ^ 2
        GoTo Finish
    End If
    hs = H - d.TBase: hp = H1 - d.TBase
    ka = CalculateKa(): kp = CalculateKp(): pa = CalculatePa(): pp = CalculatePp()
    net = CalculateMO(d): vLimit = ProjectShearLimit()
    beta = FMMax(0.65, 0.85 - 0.05 * FMMax(currentWSD.fcPrime / FM_PSI_KSC - 4000#, 0) / 1000#)
    rhoMax = 0.75 * 0.85 * beta * currentWSD.fcPrime / currentWSD.fy * 0.003 / (0.003 + currentWSD.fy / 2040000#)
    Call FMAdd(merit, (0.075 - cover) / 0.075)
    w(0) = CalculateW1(d, wx(0)): w(1) = CalculateW2(d, wx(1))
    w(2) = CalculateW3(d, wx(2)): w(3) = CalculateW4(d, wx(3))
    wd = gamma_concrete * d.TBase + gamma_soil * hs
    moments(2) = wd * d.LHeel ^ 2 / 2#: shears(2) = wd * d.LHeel
    For mask = 0 To 15
        weight = 0: mr = 0
        For i = 0 To 3
            factor(i) = 1#: If (mask And (2 ^ i)) <> 0 Then factor(i) = 0.85
            weight = weight + factor(i) * w(i): mr = mr + factor(i) * w(i) * wx(i)
        Next i
        qt = 4# * weight / d.Base - 6# * (mr - net) / d.Base ^ 2
        qh = -2# * weight / d.Base + 6# * (mr - net) / d.Base ^ 2
        Call FMAdd(merit, -qt / qa): Call FMAdd(merit, -qh / qa)
        Call FMAdd(merit, FMMax(qt, qh) * FS_BC_MIN / qa - 1#)
        Call FMAdd(merit, FS_OT_MIN * pa * H / 3# / (mr + pp * H1 / 3#) - 1#)
        Call FMAdd(merit, FS_SL_MIN * pa / (mu * weight + pp) - 1#)
        slope = (qh - qt) / d.Base
        Call ProjectCantileverEnvelope(qt, slope, d.LToe, nm, rm, nv, rv)
        moments(1) = FMMax(moments(1), nm): shears(1) = FMMax(shears(1), nv)
        For i = 1 To 2
            If i = 1 Then
                length = d.LToe: wd = factor(3) * gamma_concrete * d.TBase + factor(0) * gamma_soil * hp
                qfree = qt - wd
            Else
                length = d.LHeel: wd = factor(3) * gamma_concrete * d.TBase + factor(1) * gamma_soil * hs
                qfree = wd - qh
            End If
            Call ProjectCantileverEnvelope(qfree, slope, length, nm, rm, nv, rv)
            reverseM(i) = FMMax(reverseM(i), rm)
            If rm > FM_TOL Then reverseV(i) = FMMax(reverseV(i), rv)
        Next i
    Next mask
    moments(0) = gamma_soil * ka * hs ^ 3 / 6#
    shears(0) = gamma_soil * ka * hs ^ 2 / 2#
    For station = 0 To 16
        y = hs * station / 16#
        localDepth = dep(0) - (d.tb - d.tt) * y / hs
        m = gamma_soil * ka * (hs - y) ^ 3 / 6#
        If Not SectionStresses(m, localDepth, steel(0), currentWSD.n, ca, st, j) Then GoTo Failed
        peakCa = FMMax(peakCa, ca): peakSt = FMMax(peakSt, st)
        netMin = FMMax(netMin, -FMNetStem(hs, hp, y, ka, kp))
    Next station
    rootRatio = Sqr(PassiveFactor * kp / ka)
    If Abs(rootRatio - 1#) > 0.000000000001 Then
        y = (rootRatio * hp - hs) / (rootRatio - 1#)
        If y >= 0 And y <= hp Then netMin = FMMax(netMin, -FMNetStem(hs, hp, y, ka, kp))
    End If
    If netMin > FM_TOL Then Call FMAdd(merit, netMin / FMMax(moments(0), 0.01))
    For i = 0 To 2
        thickness = d.TBase: If i = 0 Then thickness = d.tb
        Call FMAdd(merit, ProjectMinimum(thickness, dep(i), i = 0) / steel(i) - 1#)
        topDepth = dep(i): If i = 0 Then topDepth = d.tt - cover - bar(i) / 2#
        Call FMAdd(merit, steel(i) / (rhoMax * 10000# * topDepth) - 1#)
        If i = 0 Then
            ca = peakCa: st = peakSt: thickness = d.tt
        Else
            If Not SectionStresses(moments(i), dep(i), steel(i), currentWSD.n, ca, st, j) Then GoTo Failed
        End If
        Call FMAdd(merit, ca / currentWSD.fc - 1#): Call FMAdd(merit, st / currentWSD.fs - 1#)
        Call FMAdd(merit, shears(i) / (10# * dep(i) * vLimit) - 1#)
        If i = 0 And d.tb <> d.tt Then Call FMAdd(merit, shears(i) / (5# * dep(i) * vLimit) - 1#)
        Call FMAdd(merit, (2# * cover + bar(i) - thickness) / FMMax(2# * cover + bar(i), 0.001))
        Call FMSpacing(merit, bar(i), WP_SP(sp(i)), thickness, cover, st)
        If i > 0 And reverseM(i) > FM_TOL Then
            cc = d.TBase - cover - bar(i)
            Call FMAdd(merit, steel(i) / (rhoMax * 10000# * revDep(i)) - 1#)
            If Not SectionStresses(reverseM(i), revDep(i), steel(i), currentWSD.n, ca, st, j) Then GoTo Failed
            Call FMAdd(merit, ca / currentWSD.fc - 1#): Call FMAdd(merit, st / currentWSD.fs - 1#)
            Call FMAdd(merit, reverseV(i) / (10# * revDep(i) * vLimit) - 1#)
            Call FMSpacing(merit, bar(i), WP_SP(sp(i)), d.TBase, cc, st)
        End If
    Next i
    FeasibilityMerit = merit
Finish:
    LastValidationReason = savedReason
    Exit Function
Failed:
    FeasibilityMerit = 1000000# + geometryDistance
    GoTo Finish
End Function
