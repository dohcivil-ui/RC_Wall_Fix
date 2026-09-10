Attribute VB_Name = "modProjectChecks"
Option Explicit
' Bounded project basis, NOT a declaration of full EIT or ACI code compliance.
' Pongnathee Ch.10 pp.348-351 member loads and WSD; retained ACI99 supplements.
' User research scope: no anchorage/lap checks or allowances; no formwork price.
' Main stem/toe/heel reinforcement only; geometric lengths; no other steel.
Public Const PROJECT_CHECK_BASIS As String = "PROJECT_BOOK_WSD_ACI99_V5_MEMBER_LOADS"
Public Const PROJECT_MEMBER_LOADS As String = "Book WSD loads: stem active only; toe upward bearing only; heel downward soil + footing only."
Public ProjectChecksEnabled As Boolean
Public ProjectTrialSummary As String
Private Const PSI_KSC As Double = 0.0703069579640175
Private Const IN_M As Double = 0.0254
Private Const CLEAR_GAP As Double = 0.0266666666666667
Private Const BASE_MOMENT_TOL As Double = 0.000000001

Public Type ProjectDetail
    OT As Double
    SL As Double
    BC As Double
    ' Moment/Shear: textbook main-face design loads. Reverse*: NET-load envelope.
    ReverseMoment(0 To 2) As Double  ' Magnitude, not a signed negative moment
    ReverseDepth(0 To 2) As Double
    ReverseShear(0 To 2) As Double
    ReverseFc(0 To 2) As Double
    ReverseFs(0 To 2) As Double
    Moment(0 To 2) As Double
    Shear(0 To 2) As Double
    Depth(0 To 2) As Double
    Steel(0 To 2) As Double
    Minimum(0 To 2) As Double
    FcBound(0 To 2) As Double
    FsBound(0 To 2) As Double
    Ld(0 To 2) As Double  ' Reserved CSV field: zero means excluded, not checked
    Length(0 To 2) As Double
    DB(0 To 2) As Integer
    SP(0 To 2) As Integer
    ShearLimit As Double
    MainWeight As Double
    ConcreteVolume As Double
    Cost As Double
End Type

Public Sub EnableProjectChecks()
    ProjectChecksEnabled = True
    WSDReviewed = True
    WSDSource = PROJECT_CHECK_BASIS
    ' Zero stem ratio means a member formula, not zero required steel.
    MinStemRatio = 0: MinBaseRatio = 0.002
End Sub

Private Function MaxV(a As Double, b As Double) As Double
    MaxV = a: If b > a Then MaxV = b
End Function
Private Function MinV(a As Double, b As Double) As Double
    MinV = a: If b < a Then MinV = b
End Function

Public Function ProjectShearLimit() As Double
    If currentWSD.fcPrime <= 0 Then Err.Raise 5, , "Material not initialized"
    ' Pongnathee pp.349/351: allowable shear in kgf/cm2, fcPrime in ksc.
    ProjectShearLimit = 0.29 * Sqr(currentWSD.fcPrime)
End Function

Public Function ProjectMinimum(thickness As Double, depth As Double, stem As Boolean) As Double
    Dim fcpsi As Double, fypsi As Double
    If thickness <= 0 Or depth <= 0 Or currentWSD.fy <= 0 Then Err.Raise 5
    fcpsi = currentWSD.fcPrime / PSI_KSC: fypsi = currentWSD.fy / PSI_KSC
    ' 10.5.1, without the 10.5.3 exemption. Uniform footing: 10.5.4/7.12.
    If stem Then
        ProjectMinimum = MaxV(3# * Sqr(fcpsi), 200#) / fypsi * 10000# * depth
    Else
        ' Retained footing minimum supplement. Book longitudinal temperature bars
        ' are separate steel, excluded from this three-main-set research model.
        ProjectMinimum = 0.002 * 10000# * thickness
    End If
End Function

Private Function SpacingOK(DB As Integer, SP As Integer, thickness As Double, stress As Double, tensionCover As Double) As Boolean
    Dim bar As Double, spacing As Double, maxSpacing As Double, fksi As Double
    If DB < DB_MIN Or DB > DB_MAX Or SP < SP_MIN Or SP > SP_MAX Then Exit Function
    If tensionCover < 0 Then Exit Function
    bar = WP_DB(DB) / 1000#: spacing = WP_SP(SP)
    If spacing - bar < MaxV(bar, CLEAR_GAP) Then Exit Function
    maxSpacing = MinV(3# * thickness, 18# * IN_M)
    ' 10.6.4 crack-control spacing, stress in ksi and clear cover in inches.
    If stress > 0 Then
        fksi = stress / PSI_KSC / 1000#
        maxSpacing = MinV(maxSpacing, MinV(540# / fksi - 2.5 * tensionCover / IN_M, 432# / fksi) * IN_M)
    End If
    SpacingOK = spacing <= maxSpacing
End Function

Public Function ProjectStemMoment(d As Design, y As Double) As Double
    Dim ha As Double
    ha = H - d.TBase - y
    If y < 0 Or ha < -0.00000001 Then Err.Raise 5
    ' Book p.351: member design does not credit passive pressure in front.
    ' Full active/passive forces remain in the overall stability calculations.
    ProjectStemMoment = gamma_soil * CalculateKa() * MaxV(ha, 0) ^ 3 / 6#
End Function

Private Sub MomentRange(d As Design, low As Double, high As Double, ByRef smallest As Double, ByRef largest As Double)
    ' Active-only triangular pressure gives a nonnegative, decreasing M(y).
    smallest = ProjectStemMoment(d, high)
    largest = ProjectStemMoment(d, low)
End Sub

Private Function NetStemMoment(d As Design, y As Double) As Double
    Dim ha As Double, hp As Double
    ha = H - d.TBase - y: hp = MaxV(H1 - d.TBase - y, 0)
    If y < 0 Or ha < -0.00000001 Then Err.Raise 5
    NetStemMoment = gamma_soil * (CalculateKa() * MaxV(ha, 0) ^ 3 - PassiveFactor * CalculateKp() * hp ^ 3) / 6#
End Function

Private Sub NetStemMomentRange(d As Design, low As Double, high As Double, ByRef smallest As Double, ByRef largest As Double)
    Dim y As Double, rootRatio As Double, value As Double, hs As Double, hp As Double
    hs = H - d.TBase: hp = H1 - d.TBase
    smallest = MinV(NetStemMoment(d, low), NetStemMoment(d, high))
    largest = MaxV(NetStemMoment(d, low), NetStemMoment(d, high))
    ' M'= -V. On the passive region V=0 has at most one root with both heights positive.
    If PassiveFactor > 0 Then
        rootRatio = Sqr(PassiveFactor * CalculateKp() / CalculateKa())
        If Abs(rootRatio - 1#) > 0.000000000001 Then
            y = (rootRatio * hp - hs) / (rootRatio - 1#)
            If y >= low And y <= high And y <= hp Then
                value = NetStemMoment(d, y)
                smallest = MinV(smallest, value): largest = MaxV(largest, value)
            End If
        End If
    End If
End Sub

Private Function StemInterval(d As Design, r As ProjectDetail, low As Double, high As Double, level As Integer) As Boolean
    Dim mn As Double, mx As Double, dep As Double, ca As Double, st As Double, j As Double, middle As Double
    Call MomentRange(d, low, high, mn, mx)
    If mn < -0.000000001 Then LastValidationReason = "REVERSED_STEM_FACE": Exit Function
    dep = r.Depth(0) - (d.tb - d.tt) * high / (H - d.TBase)
    If dep <= 0 Then LastValidationReason = "NONPOSITIVE_LOCAL_DEPTH": Exit Function
    ' Certified interval bound: max moment paired with minimum depth in interval.
    If Not SectionStresses(mx, dep, r.Steel(0), currentWSD.n, ca, st, j) Then Exit Function
    If ca <= currentWSD.fc And st <= currentWSD.fs Then
        r.FcBound(0) = MaxV(r.FcBound(0), ca): r.FsBound(0) = MaxV(r.FsBound(0), st)
        StemInterval = True: Exit Function
    End If
    middle = (low + high) / 2#
    dep = r.Depth(0) - (d.tb - d.tt) * middle / (H - d.TBase)
    Call SectionStresses(ProjectStemMoment(d, middle), dep, r.Steel(0), currentWSD.n, ca, st, j)
    If ca > currentWSD.fc Or st > currentWSD.fs Or level >= 22 Then
        LastValidationReason = "STEM_STRESS_OR_INTERVAL_BOUND": Exit Function
    End If
    If Not StemInterval(d, r, low, middle, level + 1) Then Exit Function
    StemInterval = StemInterval(d, r, middle, high, level + 1)
End Function

Public Function ProjectStemStressCheck(d As Design, r As ProjectDetail) As Boolean
    Dim hs As Double, hp As Double, minimum As Double, netMaximum As Double
    hs = H - d.TBase: hp = H1 - d.TBase
    If hs <= 0 Or hp < 0 Or hp > hs Then Err.Raise 5, , "Invalid stem heights"
    ' Book active-only design must not hide actual opposite-face stem tension.
    Call NetStemMomentRange(d, 0, hs, minimum, netMaximum)
    If minimum < -BASE_MOMENT_TOL Then LastValidationReason = "REVERSED_STEM_FACE": Exit Function
    r.Depth(0) = SectionDepth(d.tb, d.ASst_DB)
    r.Steel(0) = CalculateAsProv(d.ASst_DB, d.ASst_Sp)
    r.FcBound(0) = 0: r.FsBound(0) = 0
    Call MomentRange(d, 0, hs, minimum, r.Moment(0))
    If Not StemInterval(d, r, 0, hp, 0) Then Exit Function
    ProjectStemStressCheck = StemInterval(d, r, hp, hs, 0)
End Function

Private Function CantileverMoment(ByVal qfree As Double, ByVal slope As Double, ByVal position As Double) As Double
    CantileverMoment = qfree * position ^ 2 / 2# + slope * position ^ 3 / 6#
End Function

Private Function CantileverShear(ByVal qfree As Double, ByVal slope As Double, ByVal position As Double) As Double
    CantileverShear = qfree * position + slope * position ^ 2 / 2#
End Function

Private Sub BaseShearInterval(ByVal qfree As Double, ByVal slope As Double, ByVal low As Double, ByVal high As Double, ByRef normalShear As Double, ByRef reverseShear As Double)
    Dim peak As Double, position As Double, middleMoment As Double
    If high <= low Then Exit Sub
    peak = MaxV(Abs(CantileverShear(qfree, slope, low)), Abs(CantileverShear(qfree, slope, high)))
    If Abs(slope) > 0.000000000001 Then
        position = -qfree / slope
        If position > low And position < high Then peak = MaxV(peak, Abs(CantileverShear(qfree, slope, position)))
    End If
    middleMoment = CantileverMoment(qfree, slope, (low + high) / 2#)
    If middleMoment < 0 Then
        reverseShear = MaxV(reverseShear, peak)
    Else
        normalShear = MaxV(normalShear, peak)
    End If
End Sub

Public Sub ProjectCantileverEnvelope(ByVal qfree As Double, ByVal slope As Double, ByVal length As Double, ByRef normalMoment As Double, ByRef reverseMoment As Double, ByRef normalShear As Double, ByRef reverseShear As Double)
    ' q(x)=qfree+slope*x; x starts at the free end. Outputs are magnitudes.
    ' Force in ton/m of wall, moment in ton-m/m; no section properties here.
    Dim value As Double, position As Double, split As Double
    If length <= 0 Then Err.Raise 5, , "Nonpositive cantilever length"
    normalMoment = 0: reverseMoment = 0: normalShear = 0: reverseShear = 0
    value = CantileverMoment(qfree, slope, length)
    normalMoment = MaxV(0, value): reverseMoment = MaxV(0, -value)
    split = length
    If Abs(slope) > 0.000000000001 Then
        ' Interior moment extremum: V(x)=0, excluding the free end.
        position = -2# * qfree / slope
        If position > 0 And position < length Then
            value = CantileverMoment(qfree, slope, position)
            normalMoment = MaxV(normalMoment, value): reverseMoment = MaxV(reverseMoment, -value)
        End If
        ' Split shear checks where the tension face changes, M(x)=0.
        position = -3# * qfree / slope
        If position > 0 And position < length Then split = position
    End If
    Call BaseShearInterval(qfree, slope, 0, split, normalShear, reverseShear)
    If split < length Then Call BaseShearInterval(qfree, slope, split, length, normalShear, reverseShear)
End Sub

Private Function BaseMemberName(ByVal member As Integer) As String
    BaseMemberName = "TOE": If member = 2 Then BaseMemberName = "HEEL"
End Function

Private Function ReversedBaseChecks(d As Design, r As ProjectDetail, ByVal member As Integer, ByVal rhoMax As Double) As Boolean
    Dim tensionCover As Double, suffix As String
    ReversedBaseChecks = True
    If r.ReverseMoment(member) <= BASE_MOMENT_TOL Then Exit Function
    ReversedBaseChecks = False
    suffix = "_REVERSED_" & BaseMemberName(member)
    ' Same single layer: measure from the new compression/tension surfaces.
    tensionCover = d.TBase - cover - WP_DB(r.DB(member)) / 1000#
    If r.ReverseDepth(member) <= 0 Or tensionCover < 0 Then LastValidationReason = "MAIN_BAR_COVER" & suffix: Exit Function
    If r.Steel(member) > rhoMax * 10000# * r.ReverseDepth(member) Then LastValidationReason = "MAXIMUM_MAIN_STEEL" & suffix: Exit Function
    If r.ReverseFc(member) > currentWSD.fc Or r.ReverseFs(member) > currentWSD.fs Then LastValidationReason = "MEMBER_FLEXURE" & suffix: Exit Function
    If r.ReverseShear(member) > r.ShearLimit Then LastValidationReason = "MEMBER_SHEAR" & suffix: Exit Function
    If Not SpacingOK(r.DB(member), r.SP(member), d.TBase, r.ReverseFs(member), tensionCover) Then LastValidationReason = "MAIN_BAR_SPACING" & suffix: Exit Function
    ReversedBaseChecks = True
End Function

Private Function BaseEnvelope(d As Design, r As ProjectDetail) As Boolean
    Dim w(0 To 3) As Double, x(0 To 3) As Double, factor(0 To 3) As Double
    Dim i As Integer, mask As Integer, weight As Double, mr As Double, qt As Double, qh As Double
    Dim net As Double, slope As Double, qface As Double, wd As Double, length As Double, j As Double
    Dim normalMoment As Double, reverseMoment As Double, normalShear As Double, reverseShear As Double
    w(0) = CalculateW1(d, x(0)): w(1) = CalculateW2(d, x(1))
    w(2) = CalculateW3(d, x(2)): w(3) = CalculateW4(d, x(3))
    r.OT = 1E+30: r.SL = 1E+30: r.BC = 1E+30
    net = CalculateMO(d)
    ' A.2.3 bounded envelope: independent 0.85/1.0 vertical dead components.
    ' Pa/Pp remain FULL in every case. NET loads keep the SAME vertical factors.
    ' Book main-face design omits beneficial loads, separately from this equilibrium.
    For mask = 0 To 15
        weight = 0: mr = 0
        For i = 0 To 3
            factor(i) = 1#: If (mask And (2 ^ i)) <> 0 Then factor(i) = 0.85
            weight = weight + factor(i) * w(i): mr = mr + factor(i) * w(i) * x(i)
        Next i
        qt = 4# * weight / d.Base - 6# * (mr - net) / d.Base ^ 2
        qh = -2# * weight / d.Base + 6# * (mr - net) / d.Base ^ 2
        If qt < 0 Or qh < 0 Then LastValidationReason = "ENVELOPE_BASE_CONTACT": Exit Function
        r.OT = MinV(r.OT, (mr + CalculatePp() * H1 / 3#) / (CalculatePa() * H / 3#))
        r.SL = MinV(r.SL, (mu * weight + CalculatePp()) / CalculatePa())
        r.BC = MinV(r.BC, qa / MaxV(qt, qh))
        If r.OT < FS_OT_MIN Or r.SL < FS_SL_MIN Or r.BC < FS_BC_MIN Then
            LastValidationReason = "ENVELOPE_STABILITY": Exit Function
        End If
        slope = (qh - qt) / d.Base
        For i = 1 To 2
            If i = 1 Then
                length = d.LToe
                wd = factor(3) * gamma_concrete * d.TBase + factor(0) * gamma_soil * (H1 - d.TBase)
                qface = qt - wd
            Else
                length = d.LHeel
                wd = factor(3) * gamma_concrete * d.TBase + factor(1) * gamma_soil * (H - d.TBase)
                qface = wd - qh
            End If
            Call ProjectCantileverEnvelope(qface, slope, length, normalMoment, reverseMoment, normalShear, reverseShear)
            ' Preserve real net-load reversal checks for the existing single bar layer.
            r.ReverseMoment(i) = MaxV(r.ReverseMoment(i), reverseMoment)
            If reverseMoment > BASE_MOMENT_TOL Then r.ReverseShear(i) = MaxV(r.ReverseShear(i), reverseShear / (10# * r.ReverseDepth(i)))
            If i = 1 Then
                ' Book pp.348-349: upward bearing alone designs toe BOTTOM bars.
                ' Integrate the actual linear pressure; do not copy inconsistent qroot.
                normalMoment = CantileverMoment(qt, slope, length)
                normalShear = CantileverShear(qt, slope, length)
            Else
                ' Book pp.350-351: soil + footing downward load designs heel TOP bars.
                ' Full downward load governs. M and V use this SAME load;
                ' p.351's 14245 kg shear substitution conflicts with p.350's 24486 kg.
                wd = gamma_concrete * d.TBase + gamma_soil * (H - d.TBase)
                normalMoment = wd * length ^ 2 / 2#
                normalShear = wd * length
            End If
            r.Moment(i) = MaxV(r.Moment(i), normalMoment)
            r.Shear(i) = MaxV(r.Shear(i), normalShear / (10# * r.Depth(i)))
        Next i
    Next mask
    ' Calculate both directions before checking limits, including rejected designs.
    For i = 1 To 2
        If Not SectionStresses(r.Moment(i), r.Depth(i), r.Steel(i), currentWSD.n, r.FcBound(i), r.FsBound(i), j) Then
            LastValidationReason = "BASE_SECTION_STRESSES_" & BaseMemberName(i): Exit Function
        End If
        If r.ReverseMoment(i) > BASE_MOMENT_TOL Then
            If Not SectionStresses(r.ReverseMoment(i), r.ReverseDepth(i), r.Steel(i), currentWSD.n, r.ReverseFc(i), r.ReverseFs(i), j) Then
                LastValidationReason = "BASE_SECTION_STRESSES_REVERSED_" & BaseMemberName(i): Exit Function
            End If
        End If
    Next i
    BaseEnvelope = True
End Function

Public Function CheckProjectDesign(d As Design, r As ProjectDetail) As Boolean
    Dim blank As ProjectDetail, i As Integer, thickness As Double, j As Double, hs As Double, hp As Double
    Dim topDepth As Double, beta As Double, rhoMax As Double
    On Error GoTo BadData
    r = blank
    LastValidationReason = "PROJECT_CRITERIA_DISABLED"
    If Not WSDCriteriaReady() Then Exit Function
    LastValidationReason = "INVALID_GEOMETRY_OR_INPUT"
    If Not GeometryOK(d) Or Not CheckHeelLayout(d) Then Exit Function
    If currentWSD.fy <= 0 Or currentWSD.fcPrime <= 0 Or cover < 0.075 Then Exit Function
    If currentWSD.fy <> currentMaterial.fy Or currentWSD.fcPrime <> currentMaterial.fc Then LastValidationReason = "MATERIAL_STATE_MISMATCH": Exit Function
    If PassiveFactor <> PROJECT_PASSIVE_FACTOR Then Exit Function
    If d.UseDoubleStem Or d.UseDoubleToe Or d.UseDoubleHeel Then LastValidationReason = "UNSUPPORTED_MAIN_DOUBLE_LAYER": Exit Function
    hs = H - d.TBase: hp = H1 - d.TBase
    r.DB(0) = d.ASst_DB: r.SP(0) = d.ASst_Sp
    r.DB(1) = d.AStoe_DB: r.SP(1) = d.AStoe_Sp
    r.DB(2) = d.ASheel_DB: r.SP(2) = d.ASheel_Sp
    r.ShearLimit = ProjectShearLimit(): AllowableShear = r.ShearLimit
    For i = 0 To 2
        If r.DB(i) < DB_MIN Or r.DB(i) > DB_MAX Or r.SP(i) < SP_MIN Or r.SP(i) > SP_MAX Then Exit Function
        thickness = d.TBase: If i = 0 Then thickness = d.tb
        r.Depth(i) = SectionDepth(thickness, r.DB(i))
        If r.Depth(i) <= 0 Then LastValidationReason = "NONPOSITIVE_LOCAL_DEPTH": Exit Function
        If i > 0 Then r.ReverseDepth(i) = cover + WP_DB(r.DB(i)) / 2000#
        r.Steel(i) = CalculateAsProv(r.DB(i), r.SP(i))
        r.Minimum(i) = ProjectMinimum(thickness, r.Depth(i), i = 0)
        If r.Steel(i) < r.Minimum(i) Then LastValidationReason = "MINIMUM_MAIN_STEEL": Exit Function
    Next i
    If Not BaseEnvelope(d, r) Then Exit Function
    If Not ProjectStemStressCheck(d, r) Then Exit Function
    ' Active-only V(y)/d(y) decreases up this linear taper when top d > 0.
    ' Its full-height maximum is at the base; no passive-pressure credit.
    r.Shear(0) = gamma_soil * CalculateKa() * hs ^ 2 / (20# * r.Depth(0))
    beta = MaxV(0.65, 0.85 - 0.05 * MaxV(currentWSD.fcPrime / PSI_KSC - 4000#, 0) / 1000#)
    ' 10.3.3 uses ULTIMATE balanced strain, not the WSD balanced stress ratio.
    rhoMax = 0.75 * 0.85 * beta * currentWSD.fcPrime / currentWSD.fy * 0.003 / (0.003 + currentWSD.fy / 2040000#)
    For i = 0 To 2
        topDepth = r.Depth(i): thickness = d.TBase
        If i = 0 Then topDepth = SectionDepth(d.tt, r.DB(i)): thickness = d.tt
        If r.Steel(i) > rhoMax * 10000# * topDepth Then LastValidationReason = "MAXIMUM_MAIN_STEEL": Exit Function
        If r.FcBound(i) > currentWSD.fc Or r.FsBound(i) > currentWSD.fs Then LastValidationReason = "MEMBER_FLEXURE": Exit Function
        If r.Shear(i) > r.ShearLimit Then LastValidationReason = "MEMBER_SHEAR": Exit Function
        ' Conservative beam trigger for tapered stem; no transverse shear reinforcement designed.
        If i = 0 And d.tb <> d.tt And r.Shear(i) > r.ShearLimit / 2# Then LastValidationReason = "STEM_SHEAR_REINFORCEMENT_REQUIRED": Exit Function
        If thickness < 2# * cover + WP_DB(r.DB(i)) / 1000# Then LastValidationReason = "MAIN_BAR_COVER": Exit Function
        If Not SpacingOK(r.DB(i), r.SP(i), thickness, r.FsBound(i), cover) Then LastValidationReason = "MAIN_BAR_SPACING": Exit Function
        If i > 0 Then
            If Not ReversedBaseChecks(d, r, i, rhoMax) Then Exit Function
        End If
    Next i
    If Not ProjectQuantityCost(d, r) Then LastValidationReason = "INVALID_PROJECT_QUANTITIES": Exit Function
    LastValidationReason = "PASS_IMPLEMENTED_PROJECT_CHECKS"
    CheckProjectDesign = True
    Exit Function
BadData:
    LastValidationReason = "INVALID_PROJECT_DATA: " & Err.Description
End Function

Public Function ProjectQuantityCost(d As Design, r As ProjectDetail) As Boolean
    ' Price quantities only. This does not certify stability or member strength.
    Dim i As Integer, hs As Double
    On Error GoTo BadQuantity
    r.Cost = 0: r.MainWeight = 0: r.ConcreteVolume = 0
    If H <= d.TBase Or d.TBase <= 0 Or cover < 0 Then Exit Function
    If d.tt <= 0 Or d.tb < d.tt Or d.Base <= 0 Then Exit Function
    If Abs(d.Base - d.LToe - d.tb - d.LHeel) > 0.000001 Then Exit Function
    If d.UseDoubleStem Or d.UseDoubleToe Or d.UseDoubleHeel Then Exit Function
    hs = H - d.TBase
    r.DB(0) = d.ASst_DB: r.SP(0) = d.ASst_Sp
    r.DB(1) = d.AStoe_DB: r.SP(1) = d.AStoe_Sp
    r.DB(2) = d.ASheel_DB: r.SP(2) = d.ASheel_Sp
    ' Only the three main-bar selections are included in the research model.
    r.Length(0) = hs - cover
    r.Length(1) = d.LToe - cover
    r.Length(2) = d.LHeel - cover
    For i = 0 To 2
        If r.Length(i) <= 0 Then Exit Function
        If r.DB(i) < DB_MIN Or r.DB(i) > DB_MAX Or r.SP(i) < SP_MIN Or r.SP(i) > SP_MAX Then Exit Function
        r.MainWeight = r.MainWeight + ProvidedSteelWeight(r.DB(i), r.SP(i), r.Length(i))
    Next i
    ' Geometric quantities per metre; user excludes anchorage, laps and formwork.
    r.ConcreteVolume = (d.tt + d.tb) * hs / 2# + d.Base * d.TBase
    If currentMaterial.concretePrice <= 0 Or currentMaterial.SteelPrice <= 0 Then Exit Function
    r.Cost = r.ConcreteVolume * currentMaterial.concretePrice + r.MainWeight * currentMaterial.SteelPrice
    ProjectQuantityCost = True
    Exit Function
BadQuantity:
    ProjectQuantityCost = False
End Function

Private Function Bars(DB As Integer, SP As Integer) As String
    Bars = "DB" & WP_DB(DB) & " @ " & Format$(WP_SP(SP), "0.000") & " m"
End Function

Public Function ProjectDesignReport(d As Design, mat As MaterialProperties, algorithm As String) As String
    Dim r As ProjectDetail, s As String, i As Integer, label As String, ok As Boolean
    currentMaterial = mat: currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    s = "Basis: " & PROJECT_CHECK_BASIS & " (implemented checks; not full EIT/ACI certification)" & vbCrLf
    s = s & PROJECT_MEMBER_LOADS & vbCrLf
    s = s & "Research scope: main stem/toe/heel steel only; secondary steel excluded; anchorage/laps excluded from checks and steel quantities; no 0.40m allowance; formwork cost excluded." & vbCrLf
    If Not d.IsValid Then
        ProjectDesignReport = s & "NO_SOLUTION / no admissible design returned; no price." & vbCrLf
        Exit Function
    End If
    ok = CheckProjectDesign(d, r)
    s = s & "Validation: " & LastValidationReason & vbCrLf
    If Not ok Then ProjectDesignReport = s & "No admissible price.": Exit Function
    s = s & algorithm & "; H=" & H & "; H1=" & H1 & "; full active + passive=1" & vbCrLf
    s = s & "tt=" & d.tt & "; tb=" & d.tb & "; TBase=" & d.TBase & "; B=" & d.Base & "; toe=" & d.LToe & "; heel=" & d.LHeel & " m" & vbCrLf
    s = s & "Stem height H-TBase=" & H - d.TBase & " m; n=Es/Ec=" & currentWSD.n & vbCrLf
    s = s & "Envelope FS: overturning=" & Format$(r.OT, "0.0000") & "; sliding=" & Format$(r.SL, "0.0000") & "; qa/qmax=" & Format$(r.BC, "0.0000") & vbCrLf
    s = s & "Tapered stem without stirrups also limited to vc/2=" & Format$(r.ShearLimit / 2#, "0.0000") & " kg/cm^2 (retained conservative beam trigger at half the book shear limit)." & vbCrLf
    s = s & "16 vertical-dead-load combinations 0.85/1.0; Pa/Pp full in all cases." & vbCrLf
    s = s & "Book main tension: toe bottom, heel top. Net-load reversed bending is checked separately at the actual bar positions; no added steel." & vbCrLf
    For i = 0 To 2
        label = "Stem": If i = 1 Then label = "Toe"
        If i = 2 Then label = "Heel"
        s = s & label & ": " & Bars(r.DB(i), r.SP(i)) & "; As=" & Format$(r.Steel(i), "0.0000") & "; As_min=" & Format$(r.Minimum(i), "0.0000") & " cm2/m; d=" & r.Depth(i) & " m" & vbCrLf
        s = s & "M book design=" & Format$(r.Moment(i), "0.0000") & " ton-m; fc_bound=" & Format$(r.FcBound(i), "0.0000") & "/" & currentWSD.fc & "; fs_bound=" & Format$(r.FsBound(i), "0.0000") & "/" & currentWSD.fs & " kg/cm^2" & vbCrLf
        s = s & "v envelope=" & Format$(r.Shear(i), "0.0000") & "/" & Format$(r.ShearLimit, "0.0000") & " kg/cm^2; modeled main length=" & Format$(r.Length(i), "0.0000") & " m" & vbCrLf
        If i > 0 And r.ReverseMoment(i) > BASE_MOMENT_TOL Then
            s = s & "Reversed bending: |M|=" & Format$(r.ReverseMoment(i), "0.0000") & " ton-m; d=" & Format$(r.ReverseDepth(i), "0.0000") & " m; cc=" & Format$(d.TBase - cover - WP_DB(r.DB(i)) / 1000#, "0.0000") & " m" & vbCrLf
            s = s & "Reversed fc=" & Format$(r.ReverseFc(i), "0.0000") & "/" & currentWSD.fc & "; fs=" & Format$(r.ReverseFs(i), "0.0000") & "/" & currentWSD.fs & "; v=" & Format$(r.ReverseShear(i), "0.0000") & "/" & Format$(r.ShearLimit, "0.0000") & " kg/cm^2" & vbCrLf
        End If
    Next i
    s = s & "Concrete=" & Format$(r.ConcreteVolume, "0.0000") & " m3/m; main steel=" & Format$(r.MainWeight, "0.0000") & " kg/m" & vbCrLf
    s = s & "Concrete+main-steel estimate=" & Format$(r.Cost, "0.00") & " Baht/m; unit rates=" & mat.concretePrice & " Baht/m3, " & mat.SteelPrice & " Baht/kg" & vbCrLf
    s = s & "Detail assumptions: clear cover >=75mm; single main layer per member; normal weight, uncoated; aggregate<=20mm; geometric quantities only." & vbCrLf
    s = s & "Stem stress uses certified interval bounds over entire taper; shear uses full-height envelope. No compression-steel credit." & vbCrLf
    s = s & "Sources: Pongnathee Ch.10 pp.342-343,348-351 loads/material WSD; retained project ACI99 stability, maximum/minimum steel and spacing supplements. No full textbook-detailing claim." & vbCrLf
    s = s & "Not checked/costed: service deflection limit, site settlement/water/seismic, joints/end details, full construction BOQ. No global-optimum guarantee." & vbCrLf
    ProjectDesignReport = s
End Function

Public Sub RecordProjectTrial()
    Dim f As Integer, price As String
    If Not ProjectChecksEnabled Then Exit Sub
    If Len(ProjectTrialSummary) = 0 And Len(RunFolder) = 0 Then Exit Sub
    f = FreeFile
    If Len(ProjectTrialSummary) = 0 Then
        ProjectTrialSummary = RunFolder & "\trial-summary.csv"
        Open ProjectTrialSummary For Output As #f
        Print #f, "trial,status,evaluations,best_cost,best_evaluation,run_folder"
        Close #f
    End If
    If RunBest.IsValid Then price = CsvNumber(RunBestCost)
    Open ProjectTrialSummary For Append As #f
    Print #f, RunTrial & "," & RunStatus & "," & EvaluationCount & "," & price & "," & RunBestEvaluation & "," & Chr$(34) & Replace$(RunFolder, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
    Close #f
End Sub
