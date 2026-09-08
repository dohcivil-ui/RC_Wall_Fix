from pathlib import Path
import re
root=Path(__file__).resolve().parent.parent
def read(n): return (root/n).read_bytes().decode('latin1').replace('\r\n','\n')
def write(n,s): (root/n).write_bytes(s.replace('\n','\r\n').encode('latin1'))
def fn(s,name,body):
    pat=rf'Public Function {name}\([\s\S]*?\nEnd Function'
    s,n=re.subn(pat,lambda _:body.strip(),s,count=1)
    assert n==1,name
    return s
s=read('modShared.bas')
s=s.replace('Option Explicit','''Option Explicit
Public BatchMode As Boolean
' H and H1 are elevations above the underside of the base; H1 is FRONT soil.
' Vertical back of stem, front taper, level dry cohesionless soil, no surcharge.
' PassiveFactor defaults to zero. Full passive requires permanent front soil
' and sufficient movement; it is not automatically available in service.
Public PassiveFactor As Double
Public LastValidationReason As String
Public Const NO_SOLUTION_COST As Double = 999999999
''',1)
s=s.replace('Public Const FS_BC_MIN As Double = 2#','Public Const FS_BC_MIN As Double = 1#')
s=s.replace('CalculatePp = 0.5 * gamma_soil * Kp * H1 * H1','CalculatePp = PassiveFactor * 0.5 * gamma_soil * Kp * H1 * H1')
s=s.replace('    If H1_toe < 0 Then H1_toe = 0','    If H1_toe < 0 Then Err.Raise 5, , "Front soil below base top: unsupported geometry"')
s=s.replace('    x1 = d.LToe / 2#\nEnd Function','''    If A_rect + A_tri > 0 Then
        x1 = (A_rect * d.LToe / 2# + A_tri * (d.LToe + base_triangle / 3#)) / (A_rect + A_tri)
    Else
        x1 = d.LToe / 2#
    End If
End Function''',1)
s=fn(s,'CalculateMomentStem','''Public Function CalculateMomentStem(d As Design) As Double
    Dim hs As Double, hp As Double
    hs = H - d.TBase: hp = H1 - d.TBase
    If hs <= 0 Or hp < 0 Or hp > hs Then Err.Raise 5, , "Invalid stem/front soil height"
    CalculateMomentStem = gamma_soil * (CalculateKa() * hs ^ 3 - PassiveFactor * CalculateKp() * hp ^ 3) / 6#
End Function''')
s=fn(s,'CalculateMomentToe','''Public Function CalculateMomentToe(d As Design) As Double
    Dim qt As Double, qh As Double, e As Double, qj As Double, wd As Double
    If Not BearingEdges(d, e, qt, qh) Then Err.Raise 5, , "Base contact invalid"
    qj = qt + (qh - qt) * d.LToe / d.Base
    wd = gamma_concrete * d.TBase + gamma_soil * (H1 - d.TBase)
    ' Positive = upward net loading: toe bottom tension.
    CalculateMomentToe = (qt - wd) * d.LToe ^ 2 / 2# + (qj - qt) * d.LToe ^ 2 / 6#
End Function''')
s=fn(s,'CalculateMomentHeel','''Public Function CalculateMomentHeel(d As Design) As Double
    Dim qt As Double, qh As Double, e As Double, qj As Double, wd As Double
    If Not BearingEdges(d, e, qt, qh) Then Err.Raise 5, , "Base contact invalid"
    qj = qt + (qh - qt) * (d.LToe + d.tb) / d.Base
    wd = gamma_concrete * d.TBase + gamma_soil * (H - d.TBase)
    ' Positive = downward net loading: heel top tension.
    CalculateMomentHeel = (wd - qh) * d.LHeel ^ 2 / 2# + (qh - qj) * d.LHeel ^ 2 / 6#
End Function''')
s=fn(s,'CheckFS_OT','''Public Function CheckFS_OT(d As Design, ByRef FS_OT As Double) As Boolean
    Dim driving As Double
    driving = CalculatePa() * H / 3#
    FS_OT = 0
    If driving <= 0 Then Exit Function
    FS_OT = (CalculateMR(d) + CalculatePp() * H1 / 3#) / driving
    CheckFS_OT = (FS_OT >= FS_OT_MIN)
End Function''')
s=fn(s,'CheckFS_BC','''Public Function CheckFS_BC(d As Design, ByRef FS_BC As Double, _
                           ByRef e As Double, ByRef q_max As Double, _
                           ByRef q_min As Double) As Boolean
    Dim qt As Double, qh As Double
    FS_BC = 0: e = 0: q_max = 0: q_min = 0
    If Not BearingEdges(d, e, qt, qh) Then Exit Function
    q_max = qt: q_min = qh
    If qh > qt Then q_max = qh: q_min = qt
    If q_max <= 0 Or qa <= 0 Then Exit Function
    ' qa is ALLOWABLE, as labelled by the existing input. Do not divide twice.
    FS_BC = qa / q_max
    CheckFS_BC = (FS_BC >= FS_BC_MIN)
End Function

Public Function BearingEdges(d As Design, ByRef e As Double, ByRef qt As Double, ByRef qh As Double) As Boolean
    Dim w As Double, w1 As Double, w2 As Double, w3 As Double, w4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    e = 0: qt = 0: qh = 0
    If Not GeometryOK(d) Then Exit Function
    w = CalculateWTotal(d, w1, w2, w3, w4, x1, x2, x3, x4)
    If w <= 0 Then Exit Function
    e = d.Base / 2# - (CalculateMR(d) - CalculateMO(d)) / w
    ' Positive e points toward TOE. Only full compressive contact is supported.
    ' Partial contact is rejected, never interpolated over the entire base.
    If Abs(e) > d.Base / 6# Then Exit Function
    qt = w / d.Base * (1# + 6# * e / d.Base)
    qh = w / d.Base * (1# - 6# * e / d.Base)
    BearingEdges = True
End Function

Public Function GeometryOK(d As Design) As Boolean
    If H <= 0 Or H1 < d.TBase Or H1 > H Then Exit Function
    If gamma_soil <= 0 Or gamma_concrete <= 0 Or qa <= 0 Then Exit Function
    If phi < 0 Or phi > 45 Or mu < 0 Or cover < 0 Then Exit Function
    If PassiveFactor < 0 Or PassiveFactor > 1 Then Exit Function
    If d.tt <= 0 Or d.tb < d.tt Or d.TBase <= 0 Or d.TBase >= H Then Exit Function
    If d.Base <= 0 Or d.LToe < 0 Or d.LHeel <= 0 Then Exit Function
    If Abs(d.Base - d.LToe - d.tb - d.LHeel) > 0.000001 Then Exit Function
    If d.tb <= cover Or d.TBase <= cover Then Exit Function
    GeometryOK = True
End Function

Public Function SectionDepth(thickness As Double, DB_idx As Integer) As Double
    If DB_idx < DB_MIN Or DB_idx > DB_MAX Then Err.Raise 5, , "Invalid steel index"
    SectionDepth = CalculateEffectiveDepth(thickness, cover, WP_DB(DB_idx))
End Function

Public Function SectionShear(d As Design, part As Integer, depth As Double) As Double
    Dim hs As Double, hp As Double, qt As Double, qh As Double, e As Double
    Dim l As Double, qcut As Double, wd As Double
    If depth <= 0 Then Err.Raise 5, , "Invalid effective depth"
    If part = 0 Then
        ' Face shear, conservative relative to a section d from support.
        hs = H - d.TBase: hp = H1 - d.TBase
        SectionShear = Abs(gamma_soil * (CalculateKa() * hs ^ 2 - PassiveFactor * CalculateKp() * hp ^ 2) / 2#)
    Else
        If Not BearingEdges(d, e, qt, qh) Then Err.Raise 5
        If part = 1 Then
            l = d.LToe - depth
            If l <= 0 Then Exit Function
            qcut = qt + (qh - qt) * l / d.Base
            wd = gamma_concrete * d.TBase + gamma_soil * (H1 - d.TBase)
            SectionShear = Abs(((qt + qcut) / 2# - wd) * l)
        Else
            l = d.LHeel - depth
            If l <= 0 Then Exit Function
            qcut = qh - (qh - qt) * l / d.Base
            wd = gamma_concrete * d.TBase + gamma_soil * (H - d.TBase)
            SectionShear = Abs((wd - (qh + qcut) / 2#) * l)
        End If
    End If
End Function''')
s=fn(s,'CheckSteelOK','''Public Function CheckSteelOK(M As Double, d_eff As Double, _
                             DB_idx As Integer, SP_idx As Integer) As Boolean
    Dim fcActual As Double, fsActual As Double, jActual As Double
    If Not SectionStresses(M, d_eff, CalculateAsProv(DB_idx, SP_idx), currentWSD.n, fcActual, fsActual, jActual) Then Exit Function
    CheckSteelOK = (fcActual <= currentWSD.fc And fsActual <= currentWSD.fs)
End Function''')
s=fn(s,'CheckDesignValid','''Public Function CheckDesignValid(d As Design, _
                                 StemDB As Integer, StemSP As Integer, _
                                 ToeDB As Integer, ToeSP As Integer, _
                                 HeelDB As Integer, HeelSP As Integer, _
                                 ByRef FS_OT As Double, ByRef FS_SL As Double, ByRef FS_BC As Double) As Boolean
    Dim e As Double, qmax As Double, qmin As Double
    Dim ds As Double, dt As Double, dh As Double
    Dim ms As Double, mt As Double, mh As Double
    On Error GoTo InvalidData
    FS_OT = 0: FS_SL = 0: FS_BC = 0: d.IsValid = False
    LastValidationReason = "INVALID_GEOMETRY_OR_INPUT"
    If Not GeometryOK(d) Then Exit Function
    If d.LHeel < 0.3 Or d.LHeel <= d.LToe Then Exit Function
    If d.UseDoubleStem Or d.UseDoubleToe Or d.UseDoubleHeel Then
        LastValidationReason = "UNSUPPORTED_DOUBLE_LAYER": Exit Function
    End If
    ds = SectionDepth(d.tb, StemDB): dt = SectionDepth(d.TBase, ToeDB): dh = SectionDepth(d.TBase, HeelDB)
    LastValidationReason = "STABILITY_OR_BASE_CONTACT"
    If Not CheckFS_OT(d, FS_OT) Then Exit Function
    If Not CheckFS_SL(d, FS_SL) Then Exit Function
    If Not CheckFS_BC(d, FS_BC, e, qmax, qmin) Then Exit Function
    ms = CalculateMomentStem(d): mt = CalculateMomentToe(d): mh = CalculateMomentHeel(d)
    LastValidationReason = "REVERSED_TENSION_FACE_UNSUPPORTED"
    If ms < 0 Or mt < 0 Or mh < 0 Then Exit Function
    LastValidationReason = "CONCRETE_OR_STEEL_STRESS"
    If Not CheckSteelOK(ms, ds, StemDB, StemSP) Then Exit Function
    If Not CheckSteelOK(mt, dt, ToeDB, ToeSP) Then Exit Function
    If Not CheckSteelOK(mh, dh, HeelDB, HeelSP) Then Exit Function
    LastValidationReason = "WSD_CRITERIA_UNVERIFIED"
    If Not WSDCriteriaReady() Then Exit Function
    LastValidationReason = "MINIMUM_STEEL"
    If CalculateAsProv(StemDB, StemSP) < MinStemRatio * 10000# * d.tb Then Exit Function
    If CalculateAsProv(ToeDB, ToeSP) < MinBaseRatio * 10000# * d.TBase Then Exit Function
    If CalculateAsProv(HeelDB, HeelSP) < MinBaseRatio * 10000# * d.TBase Then Exit Function
    LastValidationReason = "CONCRETE_SHEAR"
    If SectionShear(d, 0, ds) / (10# * ds) > AllowableShear Then Exit Function
    If SectionShear(d, 1, dt) / (10# * dt) > AllowableShear Then Exit Function
    If SectionShear(d, 2, dh) / (10# * dh) > AllowableShear Then Exit Function
    d.FS_OT = FS_OT: d.FS_SL = FS_SL: d.FS_BC = FS_BC
    d.IsValid = True
    LastValidationReason = "VERIFIED_CONFIGURED_CHECKS_ONLY"
    CheckDesignValid = True
    Exit Function
InvalidData:
    LastValidationReason = "INVALID_DATA: " & Err.Description
End Function''')
# No-solution must not look like a zero-cost optimum.
s=s.replace('    CalculateCost = CalculateCostFull', '    If Not GeometryOK(d) Then CalculateCost = NO_SOLUTION_COST: Exit Function\n    CalculateCost = CalculateCostFull',1)
# Keep the established report layout, but derive all section values from same engine.
s=s.replace('CalculateMomentStem()', 'CalculateMomentStem(d)')
s=s.replace('    wsd = CalculateWSDParameters(mat.fy, mat.fc)','''    If Not GeometryOK(d) Then
        FormatResults = "NO_SOLUTION / no admissible design returned. WSD criteria may still be unverified."
        Exit Function
    End If
    currentMaterial = mat
    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    wsd = currentWSD''')
for field,th,db in [('stem','tb','ASst_DB'),('toe','TBase','AStoe_DB'),('heel','TBase','ASheel_DB')]:
    s=re.sub(rf'    d_{field} = d\.[^\n]+',f'    d_{field} = SectionDepth(d.{th}, d.{db})',s)
    s=re.sub(rf'    As_req_{field} = [^\n]+',f'    As_req_{field} = RequiredSectionSteel(M_{field}, d_{field}, d.{th}, wsd, {str(field=="stem")})',s)
s=s.replace('    FormatResults = result','''    Dim checkOK As Boolean
    checkOK = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, FS_OT, FS_SL, FS_BC)
    result = "Design check: " & LastValidationReason & vbCrLf & result
    result = result & "Stem height H-TBase=" & Format(H - d.TBase, "0.000") & " m" & vbCrLf
    result = result & "Clear cover to bar surface=" & Format(cover, "0.000") & " m; d=t-cover-db/2" & vbCrLf
    result = result & "Passive fraction=" & PassiveFactor & "; signed e positive toward TOE; full contact required" & vbCrLf
    result = result & "qa is ALLOWABLE; shear and minimum steel require sourced WSD criteria." & vbCrLf
    result = result & "Heuristic best found only; no guarantee of global optimum." & vbCrLf
    FormatResults = result''')
s=s.replace('" >= 2.0 "\n    If FS_BC >= 2#','" >= 1.0 "\n    If FS_BC >= FS_BC_MIN')
write('modShared.bas',s)
s=read('RC_RT_HCA_v2.vbp').replace('C:\\Users\\moosu\\Downloads\\modShared.bas','modShared.bas')
write('RC_RT_HCA_v2.vbp',s)
s=read('modWSD.bas').replace("' Based on ACI 318-19 and Thai Standard", "' Legacy parameters, NOT verified against Thai Standard")
s=s.replace('    n As Integer', '    n As Double').replace('(Ec/Es) = 9','(Es/Ec); legacy assumed value = 9')
s=s.replace('Option Explicit','''Option Explicit
' No guessed EIT clauses: these remain unset until reviewed against 011007-19.
Public AllowableShear As Double   ' kgf/cm2, for this material and member
Public MinStemRatio As Double    ' main vertical steel / gross concrete area
Public MinBaseRatio As Double    ' main slab steel / gross concrete area
Public WSDSource As String       ' edition, clause/page and applicability
Public WSDReviewed As Boolean

Public Function WSDCriteriaReady() As Boolean
    WSDCriteriaReady = WSDReviewed And Len(WSDSource) > 0 And AllowableShear > 0 And MinStemRatio > 0 And MinBaseRatio > 0
End Function

Public Function SectionStresses(M As Double, depth As Double, steel As Double, n As Double, _
    ByRef fcActual As Double, ByRef fsActual As Double, ByRef jActual As Double) As Boolean
    Dim rho As Double, k As Double, dc As Double, moment As Double
    fcActual = 0: fsActual = 0: jActual = 0
    If depth <= 0 Or steel <= 0 Or n <= 0 Then Exit Function
    dc = depth * 100#: rho = steel / (100# * dc)
    ' Cracked transformed section: k^2/2 = n*rho*(1-k).
    k = Sqr((n * rho) ^ 2 + 2# * n * rho) - n * rho
    jActual = 1# - k / 3#: moment = Abs(M) * 100000#
    fcActual = 2# * moment / (100# * k * jActual * dc ^ 2)
    fsActual = moment / (steel * jActual * dc)
    SectionStresses = True
End Function

Public Function RequiredSectionSteel(M As Double, depth As Double, thickness As Double, wsd As WSDParams, stem As Boolean) As Double
    Dim lo As Double, hi As Double, mid As Double, i As Integer
    Dim c As Double, s As Double, j As Double, minimum As Double
    If depth <= 0 Then Err.Raise 5, , "Invalid effective depth"
    lo = 0.000001: hi = 10000#
    For i = 1 To 80
        mid = (lo + hi) / 2#
        Call SectionStresses(M, depth, mid, wsd.n, c, s, j)
        If c > wsd.fc Or s > wsd.fs Then lo = mid Else hi = mid
    Next i
    If stem Then minimum = MinStemRatio * 10000# * thickness Else minimum = MinBaseRatio * 10000# * thickness
    RequiredSectionSteel = hi
    If minimum > hi Then RequiredSectionSteel = minimum
End Function
''',1)
s=fn(s,'CalculateEffectiveDepth','''Public Function CalculateEffectiveDepth(thickness As Double, cover As Double, Optional bar_mm As Double = 0) As Double
    If thickness <= 0 Or cover < 0 Or bar_mm <= 0 Then Err.Raise 5, , "Thickness, clear cover and actual bar diameter required"
    CalculateEffectiveDepth = thickness - cover - bar_mm / 2000#
    If CalculateEffectiveDepth <= 0 Then Err.Raise 5, , "Nonpositive effective depth"
End Function''')
s=s.replace('DB_index > 6','DB_index > UBound(DBArray)').replace('SP_index > 10','SP_index > UBound(SPArray)')
s=s.replace('If d2 <= 0 Then d2 = 0.01', 'If d2 <= 0 Then Err.Raise 5, , "Invalid second layer depth"')
s=s.replace('CalculateRhoBalanced = (wsd.fc / wsd.fs) * wsd.k','CalculateRhoBalanced = 0.5 * (wsd.fc / wsd.fs) * wsd.k')
s=fn(s,'CalculateMinimumSteel','''Public Function CalculateMinimumSteel(b As Double, d As Double) As Double
    Err.Raise 5, , "Select sourced member-specific minimum steel on gross thickness"
End Function''')
s=s.replace('Format(As_required, "0.00") & " cm2 --> OK"', 'Format(As_required, "0.00") & " cm2 --> " & IIf(As_provided >= As_required, "OK (area only)", "FAIL")')
write('modWSD.bas',s)
