Attribute VB_Name = "modProjectChecks"
Option Explicit
' Bounded project basis, NOT a declaration of full EIT or ACI code compliance.
' WSD material stresses: supplied Pongnathee reference; ACI 318-99 supplements.
' User research scope: no anchorage/lap checks or allowances; no formwork price.
' Main stem/toe/heel reinforcement only; geometric lengths; no other steel.
Public Const PROJECT_CHECK_BASIS As String = "PROJECT_WSD_ACI99_V3_MAIN_ONLY"
Public ProjectChecksEnabled As Boolean
Public ProjectTrialSummary As String
Private Const PSI_KSC As Double = 0.0703069579640175
Private Const IN_M As Double = 0.0254
Private Const CLEAR_GAP As Double = 0.0266666666666667

Public Type ProjectDetail
    OT As Double
    SL As Double
    BC As Double
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
    ProjectShearLimit = 1.1 * Sqr(currentWSD.fcPrime / PSI_KSC) * PSI_KSC
End Function

Public Function ProjectMinimum(thickness As Double, depth As Double, stem As Boolean) As Double
    Dim fcpsi As Double, fypsi As Double
    If thickness <= 0 Or depth <= 0 Or currentWSD.fy <= 0 Then Err.Raise 5
    fcpsi = currentWSD.fcPrime / PSI_KSC: fypsi = currentWSD.fy / PSI_KSC
    ' 10.5.1, without the 10.5.3 exemption. Uniform footing: 10.5.4/7.12.
    If stem Then
        ProjectMinimum = MaxV(3# * Sqr(fcpsi), 200#) / fypsi * 10000# * depth
    Else
        ' .002 also retained by the supplied textbook for SD40; no Grade60 claim.
        ProjectMinimum = 0.002 * 10000# * thickness
    End If
End Function

Private Function SpacingOK(DB As Integer, SP As Integer, thickness As Double, stress As Double) As Boolean
    Dim bar As Double, spacing As Double, maxSpacing As Double, fksi As Double
    If DB < DB_MIN Or DB > DB_MAX Or SP < SP_MIN Or SP > SP_MAX Then Exit Function
    bar = WP_DB(DB) / 1000#: spacing = WP_SP(SP)
    If spacing - bar < MaxV(bar, CLEAR_GAP) Then Exit Function
    maxSpacing = MinV(3# * thickness, 18# * IN_M)
    ' 10.6.4 crack-control spacing, stress in ksi and clear cover in inches.
    If stress > 0 Then
        fksi = stress / PSI_KSC / 1000#
        maxSpacing = MinV(maxSpacing, MinV(540# / fksi - 2.5 * cover / IN_M, 432# / fksi) * IN_M)
    End If
    SpacingOK = spacing <= maxSpacing
End Function

Public Function ProjectStemMoment(d As Design, y As Double) As Double
    Dim ha As Double, hp As Double
    ha = H - d.TBase - y: hp = MaxV(H1 - d.TBase - y, 0)
    If y < 0 Or ha < -0.00000001 Then Err.Raise 5
    ProjectStemMoment = gamma_soil * (CalculateKa() * MaxV(ha, 0) ^ 3 - PassiveFactor * CalculateKp() * hp ^ 3) / 6#
End Function

Private Sub MomentRange(d As Design, low As Double, high As Double, ByRef smallest As Double, ByRef largest As Double)
    Dim y As Double, rootRatio As Double, value As Double, hs As Double, hp As Double
    hs = H - d.TBase: hp = H1 - d.TBase
    smallest = MinV(ProjectStemMoment(d, low), ProjectStemMoment(d, high))
    largest = MaxV(ProjectStemMoment(d, low), ProjectStemMoment(d, high))
    ' M'= -V. On the passive region V=0 has at most one root with both heights positive.
    If PassiveFactor > 0 Then
        rootRatio = Sqr(PassiveFactor * CalculateKp() / CalculateKa())
        If Abs(rootRatio - 1#) > 0.000000000001 Then
            y = (rootRatio * hp - hs) / (rootRatio - 1#)
            If y >= low And y <= high And y <= hp Then
                value = ProjectStemMoment(d, y)
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
    Dim hs As Double, hp As Double, minimum As Double
    hs = H - d.TBase: hp = H1 - d.TBase
    If hs <= 0 Or hp < 0 Or hp > hs Then Err.Raise 5, , "Invalid stem heights"
    r.Depth(0) = SectionDepth(d.tb, d.ASst_DB)
    r.Steel(0) = CalculateAsProv(d.ASst_DB, d.ASst_Sp)
    r.FcBound(0) = 0: r.FsBound(0) = 0
    Call MomentRange(d, 0, hs, minimum, r.Moment(0))
    If Not StemInterval(d, r, 0, hp, 0) Then Exit Function
    ProjectStemStressCheck = StemInterval(d, r, hp, hs, 0)
End Function

Private Function BaseEnvelope(d As Design, r As ProjectDetail) As Boolean
    Dim w(0 To 3) As Double, x(0 To 3) As Double, factor(0 To 3) As Double
    Dim i As Integer, mask As Integer, weight As Double, mr As Double, qt As Double, qh As Double
    Dim net As Double, slope As Double, qface As Double, wd As Double, shear As Double, moment As Double
    Dim qroot As Double, length As Double, rootShear As Double
    w(0) = CalculateW1(d, x(0)): w(1) = CalculateW2(d, x(1))
    w(2) = CalculateW3(d, x(2)): w(3) = CalculateW4(d, x(3))
    r.OT = 1E+30: r.SL = 1E+30: r.BC = 1E+30
    net = CalculateMO(d)
    ' A.2.3 bounded envelope: independent 0.85/1.0 vertical dead components.
    ' Pa/Pp remain FULL in every case; pressure and footing loads use the SAME factors.
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
            ' Coordinate starts at free end. Net load=qface+slope*x for both members.
            moment = qface * length ^ 2 / 2# + slope * length ^ 3 / 6#
            If qface < -0.000000001 Or moment < -0.000000001 Then LastValidationReason = "REVERSED_BASE_FACE": Exit Function
            r.Moment(i) = MaxV(r.Moment(i), moment)
            If Abs(slope) > 0.000000000001 Then
                qroot = -2# * qface / slope
                If qroot > 0 And qroot < length Then r.Moment(i) = MaxV(r.Moment(i), qface * qroot ^ 2 / 2# + slope * qroot ^ 3 / 6#)
            End If
            ' Entire cantilever checked, including support face (no d-section exemption).
            shear = Abs(qface * length + slope * length ^ 2 / 2#)
            If Abs(slope) > 0.000000000001 Then
                qroot = -qface / slope
                If qroot > 0 And qroot < length Then
                    rootShear = Abs(qface * qroot + slope * qroot ^ 2 / 2#)
                    shear = MaxV(shear, rootShear)
                End If
            End If
            r.Shear(i) = MaxV(r.Shear(i), shear / (10# * r.Depth(i)))
        Next i
    Next mask
    BaseEnvelope = True
End Function

Public Function CheckProjectDesign(d As Design, r As ProjectDetail) As Boolean
    Dim blank As ProjectDetail, i As Integer, thickness As Double, j As Double, hs As Double, hp As Double
    Dim topDepth As Double, beta As Double, rhoMax As Double, y As Double
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
        r.Steel(i) = CalculateAsProv(r.DB(i), r.SP(i))
        r.Minimum(i) = ProjectMinimum(thickness, r.Depth(i), i = 0)
        If r.Steel(i) < r.Minimum(i) Then LastValidationReason = "MINIMUM_MAIN_STEEL": Exit Function
    Next i
    If Not BaseEnvelope(d, r) Then Exit Function
    If Not ProjectStemStressCheck(d, r) Then Exit Function
    r.Shear(0) = StemShearStressEnvelope(d, r.Depth(0), y)
    beta = MaxV(0.65, 0.85 - 0.05 * MaxV(currentWSD.fcPrime / PSI_KSC - 4000#, 0) / 1000#)
    ' 10.3.3 uses ULTIMATE balanced strain, not the WSD balanced stress ratio.
    rhoMax = 0.75 * 0.85 * beta * currentWSD.fcPrime / currentWSD.fy * 0.003 / (0.003 + currentWSD.fy / 2040000#)
    For i = 0 To 2
        topDepth = r.Depth(i): thickness = d.TBase
        If i = 0 Then topDepth = SectionDepth(d.tt, r.DB(i)): thickness = d.tt
        If r.Steel(i) > rhoMax * 10000# * topDepth Then LastValidationReason = "MAXIMUM_MAIN_STEEL": Exit Function
        If i > 0 Then Call SectionStresses(r.Moment(i), r.Depth(i), r.Steel(i), currentWSD.n, r.FcBound(i), r.FsBound(i), j)
        If r.FcBound(i) > currentWSD.fc Or r.FsBound(i) > currentWSD.fs Then LastValidationReason = "MEMBER_FLEXURE": Exit Function
        If r.Shear(i) > r.ShearLimit Then LastValidationReason = "MEMBER_SHEAR": Exit Function
        ' Conservative beam trigger for tapered stem; no transverse shear reinforcement designed.
        If i = 0 And d.tb <> d.tt And r.Shear(i) > r.ShearLimit / 2# Then LastValidationReason = "STEM_SHEAR_REINFORCEMENT_REQUIRED": Exit Function
        If thickness < 2# * cover + WP_DB(r.DB(i)) / 1000# Then LastValidationReason = "MAIN_BAR_COVER": Exit Function
        If Not SpacingOK(r.DB(i), r.SP(i), thickness, r.FsBound(i)) Then LastValidationReason = "MAIN_BAR_SPACING": Exit Function
    Next i
    ' Only the three main-bar selections are included in the research model.
    r.Length(0) = hs - cover
    r.Length(1) = d.LToe - cover
    r.Length(2) = d.LHeel - cover
    For i = 0 To 2
        If r.Length(i) <= 0 Then LastValidationReason = "INVALID_MAIN_LENGTH": Exit Function
        r.MainWeight = r.MainWeight + ProvidedSteelWeight(r.DB(i), r.SP(i), r.Length(i))
    Next i
    ' Geometric quantities per metre; user excludes anchorage, laps and formwork.
    r.ConcreteVolume = (d.tt + d.tb) * hs / 2# + d.Base * d.TBase
    If currentMaterial.concretePrice <= 0 Or currentMaterial.SteelPrice <= 0 Then LastValidationReason = "INVALID_PRICE_INPUT": Exit Function
    r.Cost = r.ConcreteVolume * currentMaterial.concretePrice + r.MainWeight * currentMaterial.SteelPrice
    LastValidationReason = "PASS_IMPLEMENTED_PROJECT_CHECKS"
    CheckProjectDesign = True
    Exit Function
BadData:
    LastValidationReason = "INVALID_PROJECT_DATA: " & Err.Description
End Function

Private Function Bars(DB As Integer, SP As Integer) As String
    Bars = "DB" & WP_DB(DB) & " @ " & Format$(WP_SP(SP), "0.000") & " m"
End Function

Public Function ProjectDesignReport(d As Design, mat As MaterialProperties, algorithm As String) As String
    Dim r As ProjectDetail, s As String, i As Integer, label As String, ok As Boolean
    currentMaterial = mat: currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    s = "Basis: " & PROJECT_CHECK_BASIS & " (implemented checks; not full EIT/ACI certification)" & vbCrLf
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
    s = s & "Tapered stem without stirrups also limited to vc/2=" & Format$(r.ShearLimit / 2#, "0.0000") & " kgf/cm2 (ACI99 A.7.5.5.1 conservative beam interpretation)." & vbCrLf
    s = s & "16 vertical-dead-load combinations 0.85/1.0; Pa/Pp full in all cases." & vbCrLf
    For i = 0 To 2
        label = "Stem": If i = 1 Then label = "Toe"
        If i = 2 Then label = "Heel"
        s = s & label & ": " & Bars(r.DB(i), r.SP(i)) & "; As=" & Format$(r.Steel(i), "0.0000") & "; As_min=" & Format$(r.Minimum(i), "0.0000") & " cm2/m; d=" & r.Depth(i) & " m" & vbCrLf
        s = s & "M envelope=" & Format$(r.Moment(i), "0.0000") & " tf.m/m; fc_bound=" & Format$(r.FcBound(i), "0.0000") & "/" & currentWSD.fc & "; fs_bound=" & Format$(r.FsBound(i), "0.0000") & "/" & currentWSD.fs & " kgf/cm2" & vbCrLf
        s = s & "v envelope=" & Format$(r.Shear(i), "0.0000") & "/" & Format$(r.ShearLimit, "0.0000") & " kgf/cm2; modeled main length=" & Format$(r.Length(i), "0.0000") & " m" & vbCrLf
    Next i
    s = s & "Concrete=" & Format$(r.ConcreteVolume, "0.0000") & " m3/m; main steel=" & Format$(r.MainWeight, "0.0000") & " kg/m" & vbCrLf
    s = s & "Concrete+main-steel estimate=" & Format$(r.Cost, "0.00") & " Baht/m; unit rates=" & mat.concretePrice & " Baht/m3, " & mat.SteelPrice & " Baht/kg" & vbCrLf
    s = s & "Detail assumptions: clear cover >=75mm; single main layer per member; normal weight, uncoated; aggregate<=20mm; geometric quantities only." & vbCrLf
    s = s & "Stem stress uses certified interval bounds over entire taper; shear uses full-height envelope. No compression-steel credit." & vbCrLf
    s = s & "Sources: Pongnathee material WSD; ACI99 A.2.3/A.7,10.3.3/10.5/10.6,7.6/7.12. See audit/PROJECT_CHECKS_IMPLEMENTATION_TH.md." & vbCrLf
    s = s & "Not checked/costed: service deflection limit, site settlement/water/seismic, joints/end details, full construction BOQ. No global-optimum guarantee." & vbCrLf
    ProjectDesignReport = s
End Function

Public Sub RecordProjectTrial()
    Dim f As Integer, price As String
    If Not ProjectChecksEnabled Then Exit Sub
    f = FreeFile
    If Len(ProjectTrialSummary) = 0 Then
        ProjectTrialSummary = RunFolder & "\trial-summary.csv"
        Open ProjectTrialSummary For Output As #f
        Print #f, "trial,seed,status,evaluations,best_cost,best_evaluation,run_folder"
        Close #f
    End If
    If RunBest.IsValid Then price = CsvNumber(RunBestCost)
    Open ProjectTrialSummary For Append As #f
    Print #f, RunTrial & "," & RunSeed & "," & RunStatus & "," & EvaluationCount & "," & price & "," & RunBestEvaluation & "," & Chr$(34) & Replace$(RunFolder, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
    Close #f
End Sub
