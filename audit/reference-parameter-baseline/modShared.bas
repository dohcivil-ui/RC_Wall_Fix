Attribute VB_Name = "modShared"
'================================================================================
' Module: modShared.bas
' Project: RC_RT_HCA v2.8 - Cantilever Retaining Wall Optimization
' Purpose: Shared Arrays, Variables, and Functions for HCA and BA
' Version: 1.0
' Date: 2567
'
' หลักการ:
' - รวม Arrays, ค่าจาก TextBox, Functions กลางไว้ที่เดียว
' - HCA และ BA เรียกใช้จาก module นี้
' - แก้ไขที่เดียว ใช้ได้ทุกที่
'================================================================================
Option Explicit
Public BatchMode As Boolean
' H and H1 are elevations above the underside of the base; H1 is FRONT soil.
' Vertical back of stem, front taper, level dry cohesionless soil, no surcharge.
' PassiveFactor defaults to zero. Full passive requires permanent front soil
' and sufficient movement; it is not automatically available in service.
Public PassiveFactor As Double
Public LastValidationReason As String
Public Const NO_SOLUTION_COST As Double = 999999999


'================================================================================
' SECTION 1: Constants
'================================================================================
Public Const PI As Double = 3.14159265358979
Public Const DEG_TO_RAD As Double = 1.74532925199433E-02

' Safety Factor Requirements
Public Const FS_OT_MIN As Double = 2#      ' Overturning
Public Const FS_SL_MIN As Double = 1.5     ' Sliding
Public Const FS_BC_MIN As Double = 1#      ' qa is allowable, not ultimate

' Steel Price
Public Const STEEL_PRICE_SD40 As Double = 24  ' Baht/kg


'================================================================================
' SECTION 2: Arrays - Index ไม่ต่อเนื่อง (เหมือน HCA เสา)
'================================================================================

' === มิติกำแพง ===
' tt:     Index 1-17   (0.200-0.600 m, step=0.025, 17 ค่า)
' tb:     Index 20-36  (0.20-1.00 m, step=0.05, 17 ค่า)
' TBase:  Index 40-54  (0.30-1.00 m, step=0.05, 15 ค่า)
' Base:   Index 60-71  (1.50-7.00 m, step=0.50, 12 ค่า)
' LToe:   Index 80-89  (0.30-1.20 m, step=0.10, 10 ค่า)

Public WP_tt(1 To 17) As Double
Public WP_tb(20 To 36) As Double
Public WP_TBase(40 To 54) As Double
Public WP_Base(60 To 71) As Double
Public WP_LToe(80 To 89) As Double

' === เหล็กเสริม ===
' DB:     Index 100-104 (12, 16, 20, 25, 28 mm, 5 ค่า)
' SP:     Index 110-113 (0.10, 0.15, 0.20, 0.25 m, 4 ค่า)

Public WP_DB(100 To 104) As Integer
Public WP_SP(110 To 113) As Double

' === Arrays สำหรับ modWSD (Index 1-based) ===
Public DBArray(1 To 5) As Integer
Public SPArray(1 To 4) As Double
Public fcArray(1 To 8) As Integer
Public concretePrice(1 To 8) As Double
Public fyArray(1 To 2) As Integer

' === Index Range Constants ===
Public Const TT_MIN As Integer = 1:     Public Const TT_MAX As Integer = 17
Public Const TB_MIN As Integer = 20:    Public Const tb_max As Integer = 36
Public Const TBASE_MIN As Integer = 40: Public Const TBase_max As Integer = 54
Public Const BASE_MIN As Integer = 60:  Public Const BASE_MAX As Integer = 71
Public Const LTOE_MIN As Integer = 80:  Public Const LTOE_MAX As Integer = 89
Public Const DB_MIN As Integer = 100:   Public Const DB_MAX As Integer = 104
Public Const SP_MIN As Integer = 110:   Public Const SP_MAX As Integer = 113

'================================================================================
' SECTION 3: Module-Level Variables (ค่าจาก TextBox)
'================================================================================
Public H As Double              ' Wall height (m)
Public H1 As Double             ' Backfill height (m)
Public gamma_soil As Double     ' Soil unit weight (ton/m3)
Public gamma_concrete As Double ' Concrete unit weight (ton/m3)
Public phi As Double            ' Friction angle (degrees)
Public mu As Double             ' Friction coefficient
Public qa As Double             ' Allowable bearing capacity (ton/m2)
Public cover As Double          ' Concrete cover (m)

' Material Properties (set from Form)
Public currentMaterial As MaterialProperties
Public currentWSD As WSDParams

'================================================================================
' SECTION 4: Tracking Variables
'================================================================================
Public EvaluationCount As Long
Public EvaluationBudget As Long
Public RunSeed As Long
Public RunTrial As Long
Public RunAlgorithm As String
Public RunBest As Design
Public RunBestCost As Double
Public RunBestEvaluation As Long
Public RunStatus As String
Public RunFolder As String
Public RunRecoveryCount As Long
Private EvaluationCSV As String
Public LastPrice As Double
Public CheckStop As Long

'================================================================================
' SECTION 5: Initialize Arrays
'================================================================================
Public Sub InitializeArrays()
    
    ' === tt (Index 1-17): 0.200-0.600 m, step=0.025 ===
    WP_tt(1) = 0.2:    WP_tt(2) = 0.225:  WP_tt(3) = 0.25
    WP_tt(4) = 0.275:  WP_tt(5) = 0.3:    WP_tt(6) = 0.325
    WP_tt(7) = 0.35:   WP_tt(8) = 0.375:  WP_tt(9) = 0.4
    WP_tt(10) = 0.425: WP_tt(11) = 0.45:  WP_tt(12) = 0.475
    WP_tt(13) = 0.5:   WP_tt(14) = 0.525: WP_tt(15) = 0.55
    WP_tt(16) = 0.575: WP_tt(17) = 0.6
    
    ' === tb (Index 20-36): 0.20-1.00 m, step=0.05 ===
    WP_tb(20) = 0.2:   WP_tb(21) = 0.25:  WP_tb(22) = 0.3
    WP_tb(23) = 0.35:  WP_tb(24) = 0.4:   WP_tb(25) = 0.45
    WP_tb(26) = 0.5:   WP_tb(27) = 0.55:  WP_tb(28) = 0.6
    WP_tb(29) = 0.65:  WP_tb(30) = 0.7:   WP_tb(31) = 0.75
    WP_tb(32) = 0.8:   WP_tb(33) = 0.85:  WP_tb(34) = 0.9
    WP_tb(35) = 0.95:  WP_tb(36) = 1#
    
    ' === TBase (Index 40-54): 0.30-1.00 m, step=0.05 ===
    WP_TBase(40) = 0.3:   WP_TBase(41) = 0.35:  WP_TBase(42) = 0.4
    WP_TBase(43) = 0.45:  WP_TBase(44) = 0.5:   WP_TBase(45) = 0.55
    WP_TBase(46) = 0.6:   WP_TBase(47) = 0.65:  WP_TBase(48) = 0.7
    WP_TBase(49) = 0.75:  WP_TBase(50) = 0.8:   WP_TBase(51) = 0.85
    WP_TBase(52) = 0.9:   WP_TBase(53) = 0.95:  WP_TBase(54) = 1#
    
    ' === Base (Index 60-71): 1.50-7.00 m, step=0.50 ===
    WP_Base(60) = 1.5:  WP_Base(61) = 2#:   WP_Base(62) = 2.5
    WP_Base(63) = 3#:   WP_Base(64) = 3.5:  WP_Base(65) = 4#
    WP_Base(66) = 4.5:  WP_Base(67) = 5#:   WP_Base(68) = 5.5
    WP_Base(69) = 6#:   WP_Base(70) = 6.5:  WP_Base(71) = 7#
    
    ' === LToe (Index 80-89): 0.30-1.20 m, step=0.10 ===
    WP_LToe(80) = 0.3:  WP_LToe(81) = 0.4:  WP_LToe(82) = 0.5
    WP_LToe(83) = 0.6:  WP_LToe(84) = 0.7:  WP_LToe(85) = 0.8
    WP_LToe(86) = 0.9:  WP_LToe(87) = 1#:   WP_LToe(88) = 1.1
    WP_LToe(89) = 1.2
    
    ' === DB (Index 100-104): 12, 16, 20, 25, 28 mm ===
    WP_DB(100) = 12:  WP_DB(101) = 16:  WP_DB(102) = 20
    WP_DB(103) = 25:  WP_DB(104) = 28
    
    ' === SP (Index 110-113): 0.10, 0.15, 0.20, 0.25 m ===
    WP_SP(110) = 0.1:   WP_SP(111) = 0.15
    WP_SP(112) = 0.2:   WP_SP(113) = 0.25
    
    ' === Arrays สำหรับ modWSD (Index 1-based) ===
    
    ' DBArray (mm)
    DBArray(1) = 12:  DBArray(2) = 16:  DBArray(3) = 20
    DBArray(4) = 25:  DBArray(5) = 28
    
    ' SPArray (m)
    SPArray(1) = 0.1:   SPArray(2) = 0.15
    SPArray(3) = 0.2:   SPArray(4) = 0.25
    
    ' fcArray (ksc)
    fcArray(1) = 180:  fcArray(2) = 210:  fcArray(3) = 240
    fcArray(4) = 280:  fcArray(5) = 300:  fcArray(6) = 320
    fcArray(7) = 350:  fcArray(8) = 400
    
    ' concretePrice (Baht/m3) - Maha Sarakham Province
    concretePrice(1) = 2100:  concretePrice(2) = 2150
    concretePrice(3) = 2200:  concretePrice(4) = 2265
    concretePrice(5) = 2300:  concretePrice(6) = 2337
    concretePrice(7) = 2380:  concretePrice(8) = 2437
    
    ' fyArray (ksc)
    fyArray(1) = 3000  ' SD30
    fyArray(2) = 4000  ' SD40
    
End Sub

'================================================================================
' SECTION 6: Utility Functions
'================================================================================

'--------------------------------------------------------------------------------
' สุ่มเลขจำนวนเต็มในช่วง Low ถึง High
'--------------------------------------------------------------------------------
Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

'--------------------------------------------------------------------------------
' คำนวณ LHeel จาก Base - LToe - tb
'--------------------------------------------------------------------------------
Public Function CalculateLHeel(Base As Double, LToe As Double, tb As Double) As Double
    CalculateLHeel = Base - LToe - tb
End Function

'--------------------------------------------------------------------------------
' Get Concrete Price by f'c (Maha Sarakham Province)Sep 68
'--------------------------------------------------------------------------------
'��Ȩԡ�¹ 2568
Public Function GetConcretePrice(fc As Integer) As Double
    Select Case fc
        Case 180: GetConcretePrice = 2337 '2010+327=2337
        Case 210: GetConcretePrice = 2384 '2057+327=2384
        Case 240: GetConcretePrice = 2430 '2103+327=2430
        Case 280: GetConcretePrice = 2524 '2197+327=2524
        Case 300: GetConcretePrice = 2570 '2243+327=2570
        Case 320: GetConcretePrice = 2617 '2290+327=2617
        Case 350: GetConcretePrice = 2783
        Case 400: GetConcretePrice = 2850
        Case Else: GetConcretePrice = 2430  ' Default fc=240 (�.�. 2568)
    End Select
End Function

'--------------------------------------------------------------------------------
' Get SD40 Material Properties
'--------------------------------------------------------------------------------
Public Function GetSD40Material(fc As Integer, _
                                concPrice As Double, _
                                steelPrice As Double) As MaterialProperties
    Dim mat As MaterialProperties
    
    mat.SteelGrade = "SD40"
    mat.fy = 4000
    mat.fc = fc
    mat.concretePrice = concPrice
    mat.steelPrice = steelPrice
    
    GetSD40Material = mat
End Function

'================================================================================
' SECTION 7: Earth Pressure Calculations
'================================================================================

'--------------------------------------------------------------------------------
' Calculate Ka (Active Earth Pressure Coefficient)
'--------------------------------------------------------------------------------
Public Function CalculateKa() As Double
    Dim phi_rad As Double
    phi_rad = phi * DEG_TO_RAD
    CalculateKa = (1# - Sin(phi_rad)) / (1# + Sin(phi_rad))
End Function

'--------------------------------------------------------------------------------
' Calculate Kp (Passive Earth Pressure Coefficient)
'--------------------------------------------------------------------------------
Public Function CalculateKp() As Double
    Dim phi_rad As Double
    phi_rad = phi * DEG_TO_RAD
    CalculateKp = (1# + Sin(phi_rad)) / (1# - Sin(phi_rad))
End Function

'--------------------------------------------------------------------------------
' Calculate Pa (Active Earth Pressure Force)
'--------------------------------------------------------------------------------
Public Function CalculatePa() As Double
    Dim Ka As Double
    Ka = CalculateKa()
    CalculatePa = 0.5 * gamma_soil * Ka * H * H
End Function

'--------------------------------------------------------------------------------
' Calculate Pp (Passive Earth Pressure Force)
'--------------------------------------------------------------------------------
Public Function CalculatePp() As Double
    Dim Kp As Double
    Kp = CalculateKp()
    CalculatePp = PassiveFactor * 0.5 * gamma_soil * Kp * H1 * H1
End Function

'================================================================================
' SECTION 8: Weight Calculations
'================================================================================

'--------------------------------------------------------------------------------
' W1: Soil on Toe
'--------------------------------------------------------------------------------
Public Function CalculateW1(d As Design, ByRef x1 As Double) As Double
    Dim H1_toe As Double, H_stem As Double
    Dim base_triangle As Double, A_rect As Double, A_tri As Double
    
    H_stem = H - d.TBase
    H1_toe = H1 - d.TBase
    If H1_toe < 0 Then Err.Raise 5, , "Front soil below base top: unsupported geometry"
    
    If H1_toe < 0.001 Then
        CalculateW1 = 0
        x1 = d.LToe / 2#
        Exit Function
    End If
    
    If H_stem > 0.001 Then
        base_triangle = (d.tb - d.tt) * H1_toe / H_stem
    Else
        base_triangle = 0
    End If
    
    A_rect = d.LToe * H1_toe
    A_tri = 0.5 * base_triangle * H1_toe
    
    CalculateW1 = (A_rect + A_tri) * gamma_soil
    If A_rect + A_tri > 0 Then
        x1 = (A_rect * d.LToe / 2# + A_tri * (d.LToe + base_triangle / 3#)) / (A_rect + A_tri)
    Else
        x1 = d.LToe / 2#
    End If
End Function

'--------------------------------------------------------------------------------
' W2: Soil on Heel
'--------------------------------------------------------------------------------
Public Function CalculateW2(d As Design, ByRef x2 As Double) As Double
    Dim H_wall As Double
    H_wall = H - d.TBase
    CalculateW2 = d.LHeel * H_wall * gamma_soil
    x2 = d.LToe + d.tb + d.LHeel / 2#
End Function

'--------------------------------------------------------------------------------
' W3: Stem (Concrete)
'--------------------------------------------------------------------------------
Public Function CalculateW3(d As Design, ByRef x3 As Double) As Double
    Dim H_stem As Double
    Dim A_rect As Double, A_tri As Double, A_total As Double
    Dim x_rect As Double, x_tri As Double, centroid_from_heel As Double
    
    H_stem = H - d.TBase
    CalculateW3 = 0.5 * (d.tt + d.tb) * H_stem * gamma_concrete
    
    A_rect = d.tt * H_stem
    x_rect = d.tt / 2#
    A_tri = 0.5 * (d.tb - d.tt) * H_stem
    x_tri = d.tt + (d.tb - d.tt) / 3#
    A_total = A_rect + A_tri
    
    If A_total > 0.001 Then
        centroid_from_heel = (A_rect * x_rect + A_tri * x_tri) / A_total
    Else
        centroid_from_heel = d.tb / 2#
    End If
    
    x3 = (d.LToe + d.tb) - centroid_from_heel
End Function

'--------------------------------------------------------------------------------
' W4: Base Slab (Concrete)
'--------------------------------------------------------------------------------
Public Function CalculateW4(d As Design, ByRef x4 As Double) As Double
    CalculateW4 = d.Base * d.TBase * gamma_concrete
    x4 = d.Base / 2#
End Function

'--------------------------------------------------------------------------------
' Total Weight
'--------------------------------------------------------------------------------
Public Function CalculateWTotal(d As Design, _
                                ByRef W1 As Double, ByRef W2 As Double, _
                                ByRef W3 As Double, ByRef W4 As Double, _
                                ByRef x1 As Double, ByRef x2 As Double, _
                                ByRef x3 As Double, ByRef x4 As Double) As Double
    W1 = CalculateW1(d, x1)
    W2 = CalculateW2(d, x2)
    W3 = CalculateW3(d, x3)
    W4 = CalculateW4(d, x4)
    CalculateWTotal = W1 + W2 + W3 + W4
End Function

'================================================================================
' SECTION 9: Moment Calculations
'================================================================================

'--------------------------------------------------------------------------------
' Moment Resisting (MR)
'--------------------------------------------------------------------------------
Public Function CalculateMR(d As Design) As Double
    Dim W1 As Double, W2 As Double, W3 As Double, W4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Call CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    CalculateMR = W1 * x1 + W2 * x2 + W3 * x3 + W4 * x4
End Function

'--------------------------------------------------------------------------------
' Moment Overturning (MO)
'--------------------------------------------------------------------------------
Public Function CalculateMO(d As Design) As Double
    Dim Pa As Double, Pp As Double
    Pa = CalculatePa()
    Pp = CalculatePp()
    CalculateMO = Pa * (H / 3#) - Pp * (H1 / 3#)
End Function

'--------------------------------------------------------------------------------
' Moment at Stem Base
'--------------------------------------------------------------------------------
Public Function CalculateMomentStem(d As Design) As Double
    Dim hs As Double, hp As Double
    hs = H - d.TBase: hp = H1 - d.TBase
    If hs <= 0 Or hp < 0 Or hp > hs Then Err.Raise 5, , "Invalid stem/front soil height"
    CalculateMomentStem = gamma_soil * (CalculateKa() * hs ^ 3 - PassiveFactor * CalculateKp() * hp ^ 3) / 6#
End Function

'--------------------------------------------------------------------------------
' Moment at Toe (using Bearing Pressure)
'--------------------------------------------------------------------------------
Public Function CalculateMomentToe(d As Design) As Double
    Dim qt As Double, qh As Double, e As Double, qj As Double, wd As Double
    If Not BearingEdges(d, e, qt, qh) Then Err.Raise 5, , "Base contact invalid"
    qj = qt + (qh - qt) * d.LToe / d.Base
    wd = gamma_concrete * d.TBase + gamma_soil * (H1 - d.TBase)
    ' Positive = upward net loading: toe bottom tension.
    CalculateMomentToe = (qt - wd) * d.LToe ^ 2 / 2# + (qj - qt) * d.LToe ^ 2 / 6#
End Function

'--------------------------------------------------------------------------------
' Moment at Heel (using Bearing Pressure)
'--------------------------------------------------------------------------------
Public Function CalculateMomentHeel(d As Design) As Double
    Dim qt As Double, qh As Double, e As Double, qj As Double, wd As Double
    If Not BearingEdges(d, e, qt, qh) Then Err.Raise 5, , "Base contact invalid"
    qj = qt + (qh - qt) * (d.LToe + d.tb) / d.Base
    wd = gamma_concrete * d.TBase + gamma_soil * (H - d.TBase)
    ' Positive = downward net loading: heel top tension.
    CalculateMomentHeel = (wd - qh) * d.LHeel ^ 2 / 2# + (qh - qj) * d.LHeel ^ 2 / 6#
End Function

'================================================================================
' SECTION 10: Safety Factor Checks
'================================================================================

'--------------------------------------------------------------------------------
' Check Overturning (FS_OT >= 2.0)
'--------------------------------------------------------------------------------
Public Function CheckFS_OT(d As Design, ByRef FS_OT As Double) As Boolean
    Dim driving As Double
    driving = CalculatePa() * H / 3#
    FS_OT = 0
    If driving <= 0 Then Exit Function
    FS_OT = (CalculateMR(d) + CalculatePp() * H1 / 3#) / driving
    CheckFS_OT = (FS_OT >= FS_OT_MIN)
End Function

'--------------------------------------------------------------------------------
' Check Sliding (FS_SL >= 1.5)
'--------------------------------------------------------------------------------
Public Function CheckFS_SL(d As Design, ByRef FS_SL As Double) As Boolean
    Dim Pa As Double, Pp As Double, W_total As Double
    Dim W1 As Double, W2 As Double, W3 As Double, W4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim Resistance As Double
    
    Pa = CalculatePa()
    Pp = CalculatePp()
    W_total = CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    
    Resistance = Pp + mu * W_total
    
    If Pa <= 0.001 Then
        FS_SL = 999
        CheckFS_SL = True
        Exit Function
    End If
    
    FS_SL = Resistance / Pa
    CheckFS_SL = (FS_SL >= FS_SL_MIN)
End Function

'--------------------------------------------------------------------------------
' Check ALLOWABLE bearing pressure (qa/qmax >= 1.0)
'--------------------------------------------------------------------------------
Public Function CheckFS_BC(d As Design, ByRef FS_BC As Double, _
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
        ' Face shear only. Passive cancellation can make a higher section govern.
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
End Function

'================================================================================
' SECTION 11: Steel Calculations
'================================================================================

'--------------------------------------------------------------------------------
' Calculate As Provided from DB and SP indices
'--------------------------------------------------------------------------------
Public Function CalculateAsProv(DB_idx As Integer, SP_idx As Integer) As Double
    Dim db_mm As Integer, spacing_m As Double
    Dim n_bars As Double, Ab As Double
    
    If DB_idx < DB_MIN Or DB_idx > DB_MAX Then
        CalculateAsProv = 0
        Exit Function
    End If
    If SP_idx < SP_MIN Or SP_idx > SP_MAX Then
        CalculateAsProv = 0
        Exit Function
    End If
    
    db_mm = WP_DB(DB_idx)
    spacing_m = WP_SP(SP_idx)
    n_bars = 1# / spacing_m
    Ab = PI * (db_mm / 2#) * (db_mm / 2#)  ' mm^2
    CalculateAsProv = n_bars * Ab / 100#    ' cm^2/m
End Function

'--------------------------------------------------------------------------------
' Check Steel Adequacy
'--------------------------------------------------------------------------------
Public Function CheckSteelOK(M As Double, d_eff As Double, _
                             DB_idx As Integer, SP_idx As Integer) As Boolean
    Dim fcActual As Double, fsActual As Double, jActual As Double
    If Not SectionStresses(M, d_eff, CalculateAsProv(DB_idx, SP_idx), currentWSD.n, fcActual, fsActual, jActual) Then Exit Function
    CheckSteelOK = (fcActual <= currentWSD.fc And fsActual <= currentWSD.fs)
End Function

'================================================================================
' SECTION 12: Cost Calculation
'================================================================================

'--------------------------------------------------------------------------------
' CalculateCost - รับแค่ Design (ดึงค่าเหล็กจาก Design.ASst_DB ฯลฯ)
'--------------------------------------------------------------------------------
Public Function CalculateCost(d As Design) As Double
    If Not GeometryOK(d) Then CalculateCost = NO_SOLUTION_COST: Exit Function
    CalculateCost = CalculateCostFull(d, d.ASst_DB, d.ASst_Sp, _
                                      d.AStoe_DB, d.AStoe_Sp, _
                                      d.ASheel_DB, d.ASheel_Sp)
End Function

'--------------------------------------------------------------------------------
' CalculateCostFull - รับ Design + Steel indices แยก
'--------------------------------------------------------------------------------
Public Function CalculateCostFull(d As Design, _
                              StemDB As Integer, StemSP As Integer, _
                              ToeDB As Integer, ToeSP As Integer, _
                              HeelDB As Integer, HeelSP As Integer) As Double
    Dim V_stem As Double, V_base As Double, V_total As Double
    Dim H_stem As Double
    Dim W_stem As Double, W_toe As Double, W_heel As Double, W_total_steel As Double
    Dim L_stem As Double, L_toe As Double, L_heel As Double
    
    H_stem = H - d.TBase
    
    ' Concrete Volume (m^3/m)
    V_stem = 0.5 * (d.tt + d.tb) * H_stem * 1#
    V_base = d.Base * d.TBase * 1#
    V_total = V_stem + V_base
    
    ' Steel Weight (kg/m)
    L_stem = H_stem + 0.4
    L_toe = d.LToe + 0.4
    L_heel = d.LHeel + 0.4
    
    W_stem = CalculateSteelWeight(StemDB, StemSP, L_stem)
    W_toe = CalculateSteelWeight(ToeDB, ToeSP, L_toe)
    W_heel = CalculateSteelWeight(HeelDB, HeelSP, L_heel)
    W_total_steel = W_stem + W_toe + W_heel
    
    ' Total Cost
    CalculateCostFull = V_total * currentMaterial.concretePrice + _
                    W_total_steel * currentMaterial.steelPrice
End Function

'--------------------------------------------------------------------------------
' Calculate Steel Weight
'--------------------------------------------------------------------------------
Private Function CalculateSteelWeight(DB_idx As Integer, SP_idx As Integer, _
                                      length As Double) As Double
    Dim db_mm As Integer, spacing_m As Double
    Dim n_bars As Double, weight_per_m As Double
    
    If DB_idx < DB_MIN Or DB_idx > DB_MAX Then
        CalculateSteelWeight = 0
        Exit Function
    End If
    If SP_idx < SP_MIN Or SP_idx > SP_MAX Then
        CalculateSteelWeight = 0
        Exit Function
    End If
    
    db_mm = WP_DB(DB_idx)
    spacing_m = WP_SP(SP_idx)
    n_bars = 1# / spacing_m
    weight_per_m = 0.00617 * db_mm * db_mm  ' kg/m
    CalculateSteelWeight = n_bars * weight_per_m * length
End Function

'================================================================================
' SECTION 13: Design Validity Check (รวมทุกเงื่อนไข)
'================================================================================

Public Function CheckDesignValid(d As Design, _
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
    If Not CheckHeelLayout(d) Then Exit Function
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
    If MemberShearStress(d, 0, ds) > AllowableShear Then Exit Function
    If MemberShearStress(d, 1, dt) > AllowableShear Then Exit Function
    If MemberShearStress(d, 2, dh) > AllowableShear Then Exit Function
    d.FS_OT = FS_OT: d.FS_SL = FS_SL: d.FS_BC = FS_BC
    d.IsValid = True
    LastValidationReason = "VERIFIED_CONFIGURED_CHECKS_ONLY"
    CheckDesignValid = True
    Exit Function
InvalidData:
    LastValidationReason = "INVALID_DATA: " & Err.Description
End Function

'================================================================================
' SECTION 14: Format Results for Display
'================================================================================

Public Function FormatResults(d As Design, mat As MaterialProperties, _
                              Optional AlgoName As String = "Hill Climbing", Optional ReportTrial As Long = 0) As String
    Dim result As String
    Dim wsd As WSDParams
    Dim M_stem As Double, M_toe As Double, M_heel As Double
    Dim FS_OT As Double, FS_SL As Double, FS_BC As Double
    Dim e As Double, q_max As Double, q_min As Double
    Dim cost As Double
    Dim Ka As Double, Kp As Double, Pa As Double, Pp As Double
    Dim W_total As Double, W1 As Double, W2 As Double, W3 As Double, W4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim d_stem As Double, d_toe As Double, d_heel As Double
    Dim As_req_stem As Double, As_req_toe As Double, As_req_heel As Double
    Dim As_prov_stem As Double, As_prov_toe As Double, As_prov_heel As Double
    
    If Not GeometryOK(d) Then
        FormatResults = "NO_SOLUTION / no admissible design returned. WSD criteria may still be unverified."
        Exit Function
    End If
    Dim reportOK As Boolean, reportQT As Double, reportQH As Double
    On Error GoTo InvalidReport
    currentMaterial = mat
    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    reportOK = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, FS_OT, FS_SL, FS_BC)
    If Not BearingEdges(d, e, reportQT, reportQH) Then
        FormatResults = "Design check: STABILITY_OR_BASE_CONTACT; no supported full-contact pressure diagram." & vbCrLf & BuildDesignCheckReport(d)
        Exit Function
    End If
    currentMaterial = mat
    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    wsd = currentWSD
    
    Ka = CalculateKa()
    Kp = CalculateKp()
    Pa = CalculatePa()
    Pp = CalculatePp()
    
    W_total = CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    
    M_stem = CalculateMomentStem(d)
    M_toe = CalculateMomentToe(d)
    M_heel = CalculateMomentHeel(d)
    
    Call CheckFS_OT(d, FS_OT)
    Call CheckFS_SL(d, FS_SL)
    Call CheckFS_BC(d, FS_BC, e, q_max, q_min)
    
    cost = CalculateCost(d)
    
    ' คำนวณ As_req และ As_prov
    d_stem = SectionDepth(d.tb, d.ASst_DB)
    d_toe = SectionDepth(d.TBase, d.AStoe_DB)
    d_heel = SectionDepth(d.TBase, d.ASheel_DB)
    
    As_req_stem = RequiredSectionSteel(M_stem, d_stem, d.tb, wsd, True)
    As_req_toe = RequiredSectionSteel(M_toe, d_toe, d.TBase, wsd, False)
    As_req_heel = RequiredSectionSteel(M_heel, d_heel, d.TBase, wsd, False)
    
    ' === MATERIAL PROPERTIES ===
    result = "=== MATERIAL PROPERTIES ===" & vbCrLf
    result = result & "Steel: " & mat.SteelGrade & " (fy=" & mat.fy & ", fs=" & wsd.fs & " ksc)" & vbCrLf
    result = result & "Concrete: f'c = " & mat.fc & " ksc (fc=" & Format(wsd.fc, "0.0") & " ksc)" & vbCrLf
    result = result & "Legacy WSD parameters (unverified unless sourced): n=" & wsd.n & ", k=" & Format(wsd.k, "0.000") & ", j=" & Format(wsd.j, "0.000") & ", R=" & Format(wsd.R, "0.00") & " ksc" & vbCrLf
    result = result & "Prices: Concrete=" & Format(mat.concretePrice, "#,##0") & " Baht/m3, Steel=" & mat.steelPrice & " Baht/kg" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === OPTIMIZATION RESULTS ===
    result = result & "=== OPTIMIZATION RESULTS ===" & vbCrLf
    result = result & "Algorithm: " & AlgoName & vbCrLf
    If ReportTrial = 0 Then ReportTrial = modDataStructures.BestTrial
    result = result & "Best Found at Trial: " & ReportTrial & vbCrLf
    result = result & "Best Found at Iteration: " & modDataStructures.BestCostIteration & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === TOTAL COST ===
    result = result & "=== TOTAL COST ===" & vbCrLf
    result = result & "Total: " & Format(cost, "#,##0.00") & " Baht/m" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === DIMENSIONS ===
    result = result & "=== DIMENSIONS ===" & vbCrLf
    result = result & "Stem_top (tt): " & Format(d.tt, "0.000") & " m" & vbCrLf
    result = result & "Stem_bottom (tb): " & Format(d.tb, "0.000") & " m" & vbCrLf
    result = result & "Base_width (B): " & Format(d.Base, "0.000") & " m" & vbCrLf
    result = result & "Base_thickness: " & Format(d.TBase, "0.000") & " m" & vbCrLf
    result = result & "Toe: " & Format(d.LToe, "0.000") & " m" & vbCrLf
    result = result & "Heel: " & Format(d.LHeel, "0.000") & " m" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === EARTH PRESSURES ===
    result = result & "=== EARTH PRESSURES ===" & vbCrLf
    result = result & "Ka = " & Format(Ka, "0.000") & ", Kp = " & Format(Kp, "0.000") & vbCrLf
    result = result & "Pa = " & Format(Pa, "0.000") & " ton, Pp = " & Format(Pp, "0.000") & " ton" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === WEIGHTS ===
    result = result & "=== WEIGHTS ===" & vbCrLf
    result = result & "W1 (Soil on Toe): " & Format(W1, "0.000") & " ton" & vbCrLf
    result = result & "W2 (Soil on Heel): " & Format(W2, "0.000") & " ton" & vbCrLf
    result = result & "W3 (Stem): " & Format(W3, "0.000") & " ton" & vbCrLf
    result = result & "W4 (Base): " & Format(W4, "0.000") & " ton" & vbCrLf
    result = result & "W_total: " & Format(W_total, "0.000") & " ton" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === STEEL REINFORCEMENT ===
    result = result & "=== STEEL REINFORCEMENT ===" & vbCrLf
    
    Dim stemDB_val As Integer, stemSP_val As Double
    Dim toeDB_val As Integer, toeSP_val As Double
    Dim heelDB_val As Integer, heelSP_val As Double
    
    ' Stem
    If d.ASst_DB >= DB_MIN And d.ASst_DB <= DB_MAX Then
        stemDB_val = WP_DB(d.ASst_DB)
    ElseIf d.ASst_DB >= 1 And d.ASst_DB <= 5 Then
        stemDB_val = DBArray(d.ASst_DB)
    Else
        stemDB_val = 20
    End If
    If d.ASst_Sp >= SP_MIN And d.ASst_Sp <= SP_MAX Then
        stemSP_val = WP_SP(d.ASst_Sp)
    ElseIf d.ASst_Sp >= 1 And d.ASst_Sp <= 4 Then
        stemSP_val = SPArray(d.ASst_Sp)
    Else
        stemSP_val = 0.15
    End If
    As_prov_stem = CalculateAsProv(d.ASst_DB, d.ASst_Sp)
    
    result = result & vbCrLf & "--- Stem ---" & vbCrLf
    result = result & "Moment: " & Format(M_stem, "0.00") & " ton-m" & vbCrLf
    result = result & "Steel: DB" & stemDB_val & " @ " & Format(stemSP_val, "0.00") & "m" & vbCrLf
    result = result & "As_req: " & Format(As_req_stem, "0.00") & ", As_prov: " & Format(As_prov_stem, "0.00") & " cm2/m" & vbCrLf
    
    ' Toe
    If d.AStoe_DB >= DB_MIN And d.AStoe_DB <= DB_MAX Then
        toeDB_val = WP_DB(d.AStoe_DB)
    ElseIf d.AStoe_DB >= 1 And d.AStoe_DB <= 5 Then
        toeDB_val = DBArray(d.AStoe_DB)
    Else
        toeDB_val = 20
    End If
    If d.AStoe_Sp >= SP_MIN And d.AStoe_Sp <= SP_MAX Then
        toeSP_val = WP_SP(d.AStoe_Sp)
    ElseIf d.AStoe_Sp >= 1 And d.AStoe_Sp <= 4 Then
        toeSP_val = SPArray(d.AStoe_Sp)
    Else
        toeSP_val = 0.15
    End If
    As_prov_toe = CalculateAsProv(d.AStoe_DB, d.AStoe_Sp)
    
    result = result & vbCrLf & "--- Toe ---" & vbCrLf
    result = result & "Moment: " & Format(M_toe, "0.00") & " ton-m" & vbCrLf
    result = result & "Steel: DB" & toeDB_val & " @ " & Format(toeSP_val, "0.00") & "m" & vbCrLf
    result = result & "As_req: " & Format(As_req_toe, "0.00") & ", As_prov: " & Format(As_prov_toe, "0.00") & " cm2/m" & vbCrLf
    
    ' Heel
    If d.ASheel_DB >= DB_MIN And d.ASheel_DB <= DB_MAX Then
        heelDB_val = WP_DB(d.ASheel_DB)
    ElseIf d.ASheel_DB >= 1 And d.ASheel_DB <= 5 Then
        heelDB_val = DBArray(d.ASheel_DB)
    Else
        heelDB_val = 20
    End If
    If d.ASheel_Sp >= SP_MIN And d.ASheel_Sp <= SP_MAX Then
        heelSP_val = WP_SP(d.ASheel_Sp)
    ElseIf d.ASheel_Sp >= 1 And d.ASheel_Sp <= 4 Then
        heelSP_val = SPArray(d.ASheel_Sp)
    Else
        heelSP_val = 0.15
    End If
    As_prov_heel = CalculateAsProv(d.ASheel_DB, d.ASheel_Sp)
    
    result = result & vbCrLf & "--- Heel ---" & vbCrLf
    result = result & "Moment: " & Format(M_heel, "0.00") & " ton-m" & vbCrLf
    result = result & "Steel: DB" & heelDB_val & " @ " & Format(heelSP_val, "0.00") & "m" & vbCrLf
    result = result & "As_req: " & Format(As_req_heel, "0.00") & ", As_prov: " & Format(As_prov_heel, "0.00") & " cm2/m" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === SAFETY FACTORS ===
    result = result & "=== SAFETY FACTORS ===" & vbCrLf
    result = result & "FS_Overturning: " & Format(FS_OT, "0.00") & " >= 2.0 "
    If FS_OT >= 2# Then result = result & "WITHIN_PROJECT_LIMIT" Else result = result & "FAIL"
    result = result & vbCrLf
    
    result = result & "FS_Sliding: " & Format(FS_SL, "0.00") & " >= 1.5 "
    If FS_SL >= 1.5 Then result = result & "WITHIN_PROJECT_LIMIT" Else result = result & "FAIL"
    result = result & vbCrLf
    
    result = result & "FS_Bearing: " & Format(FS_BC, "0.00") & " >= 1.0 "
    If FS_BC >= FS_BC_MIN Then result = result & "WITHIN_PROJECT_LIMIT" Else result = result & "FAIL"
    result = result & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === BEARING CAPACITY ===
    result = result & "=== BEARING CAPACITY ===" & vbCrLf
    result = result & "Eccentricity (e): " & Format(e, "0.000") & " m" & vbCrLf
    result = result & "q_max: " & Format(q_max, "0.00") & " ton/m2 (allowable: " & Format(qa, "0.00") & ")" & vbCrLf
    result = result & "q_min: " & Format(q_min, "0.00") & " ton/m2" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    Dim checkOK As Boolean
    checkOK = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, FS_OT, FS_SL, FS_BC)
    result = result & SectionCheckSummary("Stem", M_stem, d_stem, As_prov_stem, SectionShear(d, 0, d_stem), d.tb, True, MemberShearStress(d, 0, d_stem))
    result = result & SectionCheckSummary("Toe", M_toe, d_toe, As_prov_toe, SectionShear(d, 1, d_toe), d.TBase, False)
    result = result & SectionCheckSummary("Heel", M_heel, d_heel, As_prov_heel, SectionShear(d, 2, d_heel), d.TBase, False)
    result = "Design check: " & LastValidationReason & vbCrLf & result
    result = result & "Stem height H-TBase=" & Format(H - d.TBase, "0.000") & " m" & vbCrLf
    result = result & "Clear cover to bar surface=" & Format(cover, "0.000") & " m; d=t-cover-db/2" & vbCrLf
    result = result & "Passive fraction=" & PassiveFactor & "; signed e positive toward TOE; full contact required" & vbCrLf
    result = result & "qa is ALLOWABLE; shear and minimum steel require sourced WSD criteria." & vbCrLf
    result = result & "Heuristic best found only; no guarantee of global optimum." & vbCrLf
    FormatResults = result & vbCrLf & BuildDesignCheckReport(d)
    Exit Function
InvalidReport:
    FormatResults = "INVALID_DATA: " & Err.Description
End Function

'================================================================================
' END OF MODULE: modShared.bas
'================================================================================

' One evaluation = one complete candidate request (including invalid geometry).
' Initial, reset and neighbor requests all enter here; reporting does not count.
Public Sub BeginSearch(budget As Long, seed As Long, algorithm As String, Optional trial As Long = 1)
    Dim dummy As Single, emptyDesign As Design, suffix As Long, basePath As String
    If budget < 1 Then Err.Raise 5, , "Evaluation budget must be positive"
    EvaluationCount = 0: EvaluationBudget = budget: RunSeed = seed
    RunTrial = trial: RunAlgorithm = algorithm
    RunBest = emptyDesign: RunBestCost = NO_SOLUTION_COST: RunBestEvaluation = 0
    RunRecoveryCount = 0
    RunStatus = "NO_SOLUTION"
    dummy = Rnd(-1): Randomize seed
    If Dir$(App.Path & "\results", vbDirectory) = "" Then MkDir App.Path & "\results"
    basePath = App.Path & "\results\" & algorithm & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-seed" & CStr(seed)
    RunFolder = basePath
    Do While Dir$(RunFolder, vbDirectory) <> ""
        suffix = suffix + 1: RunFolder = basePath & "-" & CStr(suffix)
    Loop
    MkDir RunFolder
    EvaluationCSV = "evaluation,entry,valid,reason,cost,best_cost,tt,tb,TBase,Base,LToe,StemDB,StemSP,ToeDB,ToeSP,HeelDB,HeelSP" & vbCrLf
End Sub

Public Function EvaluateCandidate(d As Design, entry As String, ByRef candidateCost As Double) As Boolean
    Dim ot As Double, sl As Double, bc As Double, ok As Boolean
    If EvaluationCount >= EvaluationBudget Then Err.Raise 5, , "Evaluation budget exhausted"
    EvaluationCount = EvaluationCount + 1
    candidateCost = NO_SOLUTION_COST
    ok = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, ot, sl, bc)
    ' Price only feasible candidates; the sentinel is NOT a rejected design price.
    If ok Then candidateCost = CalculateCost(d)
    d.TotalCost = candidateCost: d.IsValid = ok
    If ok Then
        If Not RunBest.IsValid Or candidateCost < RunBestCost Then
            RunBest = d: RunBestCost = candidateCost: RunBestEvaluation = EvaluationCount
            RunStatus = "SOLUTION_FOUND_CONFIGURED_CHECKS"
        End If
    End If
    EvaluationCSV = EvaluationCSV & EvaluationCount & "," & entry & "," & CStr(ok) & "," & Replace(LastValidationReason, ",", ";") & "," & _
        CsvNumber(candidateCost) & "," & CsvNumber(RunBestCost) & "," & CsvNumber(d.tt) & "," & CsvNumber(d.tb) & "," & _
        CsvNumber(d.TBase) & "," & CsvNumber(d.Base) & "," & CsvNumber(d.LToe) & "," & _
        d.ASst_DB & "," & d.ASst_Sp & "," & d.AStoe_DB & "," & d.AStoe_Sp & "," & d.ASheel_DB & "," & d.ASheel_Sp & vbCrLf
    EvaluateCandidate = ok
End Function

Public Function CsvNumber(value As Double) As String
    CsvNumber = Replace(Format$(value, "0.0000000000"), ",", ".")
End Function

Public Sub FinishSearch()
    Dim f As Integer
    modDataStructures.BestCostIteration = RunBestEvaluation
    If Not RunBest.IsValid Then RunBest.TotalCost = NO_SOLUTION_COST
    f = FreeFile
    Open RunFolder & "\evaluations.csv" For Output As #f
    Print #f, EvaluationCSV;
    Close #f
    f = FreeFile
    Open RunFolder & "\run.txt" For Output As #f
    Print #f, "Status=" & RunStatus
    Print #f, "Algorithm=" & RunAlgorithm & "; Trial=" & RunTrial
    Print #f, "Seed=" & RunSeed & "; Evaluations=" & EvaluationCount & "; Budget=" & EvaluationBudget
    Print #f, "BestEvaluation=" & RunBestEvaluation & "; Recoveries=" & RunRecoveryCount
    Print #f, "H=" & H & "; H1=" & H1 & "; gamma_soil=" & gamma_soil & "; gamma_concrete=" & gamma_concrete
    Print #f, "phi=" & phi & "; mu=" & mu & "; qa_allowable=" & qa & "; clear_cover=" & cover
    Print #f, "fc_prime=" & currentMaterial.fc & "; fy=" & currentMaterial.fy & "; passive_fraction=" & PassiveFactor
    Print #f, "WSDReviewed=" & WSDReviewed & "; source=" & WSDSource
    Print #f, "AllowableShear=" & AllowableShear & "; MinStemRatio=" & MinStemRatio & "; MinBaseRatio=" & MinBaseRatio
    Print #f, FormatResults(RunBest, currentMaterial, RunAlgorithm, RunTrial)
    Close #f
End Sub

Public Function UniqueExportPath(stem As String) As String
    Dim n As Long, p As String
    p = RunFolder & "\" & stem & ".csv"
    Do While Len(Dir$(p)) > 0
        n = n + 1: p = RunFolder & "\" & stem & "-" & n & ".csv"
    Loop
    UniqueExportPath = p
End Function

Public Sub SeedSearchRandom(seed As Long)
    Dim dummy As Single
    dummy = Rnd(-1): Randomize seed
End Sub

Private Function SectionCheckSummary(label As String, M As Double, depth As Double, steel As Double, shear As Double, thickness As Double, stem As Boolean, Optional nominalStress As Double = -1) As String
    Dim c As Double, s As Double, j As Double, minimum As Double
    Call SectionStresses(M, depth, steel, currentWSD.n, c, s, j)
    If stem Then minimum = MinStemRatio * thickness * 10000# Else minimum = MinBaseRatio * thickness * 10000#
    If nominalStress < 0 Then nominalStress = shear / (10# * depth)
    SectionCheckSummary = label & ": d=" & Format(depth, "0.0000") & " m; fc_actual=" & Format(c, "0.00") & _
        "/" & currentWSD.fc & "; fs_actual=" & Format(s, "0.00") & "/" & currentWSD.fs & " kgf/cm2" & vbCrLf & _
        "v=" & Format(nominalStress, "0.000") & "/" & AllowableShear & " kgf/cm2; As_min=" & minimum & " cm2/m" & vbCrLf
    If Not WSDCriteriaReady() Then SectionCheckSummary = SectionCheckSummary & "Allowables/minimums UNVERIFIED (zero means unset, never a passing criterion)." & vbCrLf
End Function


' Read-only audit: continue independent checks after a failure. All arithmetic
' uses the same helpers as CheckDesignValid; unset limits are never zero limits.
Private Function AuditRow(item As String, actual As String, criterion As String, outcome As String, source As String) As String
    AuditRow = "| " & item & " | " & actual & " | " & criterion & " | " & outcome & " | " & source & " |" & vbCrLf
End Function

Private Function AuditCompare(item As String, actual As Double, limit As Double, minimum As Boolean, units As String, source As String, ByRef failed As Boolean) As String
    Dim within As Boolean, relation As String, outcome As String
    If minimum Then
        within = actual >= limit: relation = ">= "
    Else
        within = actual <= limit: relation = "<= "
    End If
    outcome = "WITHIN_LISTED_LIMIT"
    If Not within Then outcome = "FAIL_LISTED_LIMIT": failed = True
    AuditCompare = AuditRow(item, Format$(actual, "0.0000") & " " & units, relation & Format$(limit, "0.0000"), outcome, source)
End Function

Private Function AuditMember(d As Design, part As Integer, label As String, thickness As Double, db As Integer, sp As Integer, ByRef failed As Boolean, ByRef yieldFailed As Boolean) As String
    Dim depth As Double, steel As Double, moment As Double, shear As Double
    Dim c As Double, st As Double, jActual As Double, ratio As Double, result As String
    Dim source As String, bound As Double
    On Error GoTo InvalidMember
    depth = SectionDepth(thickness, db): steel = CalculateAsProv(db, sp)
    If steel <= 0 Or currentMaterial.fy <= 0 Then Err.Raise 5, , "Invalid steel/material"
    Select Case part
        Case 0: moment = CalculateMomentStem(d): ratio = MinStemRatio
        Case 1: moment = CalculateMomentToe(d): ratio = MinBaseRatio
        Case 2: moment = CalculateMomentHeel(d): ratio = MinBaseRatio
    End Select
    result = AuditRow(label & " bars", "DB" & WP_DB(db) & " @ " & Format$(WP_SP(sp), "0.00") & " m", "specified candidate", "INPUT", "Design")
    result = result & AuditRow(label & " effective depth", Format$(depth, "0.0000") & " m", "t-cover-db/2 > 0", "CALCULATED", "Main bar outermost; single layer assumption")
    result = result & AuditRow(label & " As", Format$(steel, "0.0000") & " cm2/m", "area per metre", "CALCULATED", "CalculateAsProv")
    result = result & AuditCompare(label & " signed M", moment, 0, True, "tf.m/m", "Supported tension face; modShared", failed)
    ' A necessary condition only: M=T*z, T<=As*fy and z<=d for the
    ' singly reinforced pure-bending model. NOT a WSD allowable capacity.
    bound = steel * currentMaterial.fy * depth / 1000#
    result = result & AuditCompare(label & " necessary yield bound", Abs(moment), bound, False, "tf.m/m", "M<=As*fy*d; no modular ratio or code allowable used", yieldFailed)
    If Not SectionStresses(moment, depth, steel, currentWSD.n, c, st, jActual) Then Err.Raise 5, , "Invalid section"
    source = "Legacy n=" & currentWSD.n & "; EIT clause UNVERIFIED"
    If WSDCriteriaReady() Then source = "Configured screening only: " & WSDSource
    result = result & AuditCompare(label & " concrete stress", c, currentWSD.fc, False, "kgf/cm2", source, failed)
    result = result & AuditCompare(label & " steel stress", st, currentWSD.fs, False, "kgf/cm2", source, failed)
    shear = MemberShearStress(d, part, depth)
    If part = 0 Then
        Dim shearHeight As Double
        shear = StemShearStressEnvelope(d, depth, shearHeight)
        result = result & AuditRow("Stem governing shear height", Format$(shearHeight, "0.0000") & " m", "above base top", "CALCULATED", "Maximum nominal abs(V)/(b*d) along tapered stem; EIT critical-section rule UNVERIFIED")
    End If
    If AllowableShear > 0 Then
        result = result & AuditCompare(label & " nominal shear V/bd", shear, AllowableShear, False, "kgf/cm2", source & "; shear section/definition require review", failed)
    Else
        result = result & AuditRow(label & " nominal shear V/bd", Format$(shear, "0.0000") & " kgf/cm2", "UNSET", "UNVERIFIED", "EIT shear limit/definition/critical section missing")
    End If
    If ratio > 0 Then
        result = result & AuditCompare(label & " minimum steel", steel, ratio * 10000# * thickness, True, "cm2/m", source & "; gross-area ratio", failed)
    Else
        result = result & AuditRow(label & " minimum steel", Format$(steel, "0.0000") & " cm2/m", "UNSET", "UNVERIFIED", "EIT applicable minimum ratio missing")
    End If
    AuditMember = result
    Exit Function
InvalidMember:
    failed = True
    AuditMember = result & AuditRow(label, Err.Description, "valid supported section required", "INVALID_DATA", "No depth clamping")
End Function

Public Function BuildDesignCheckReport(d As Design, Optional stemOnly As Boolean = False) As String
    Dim result As String, failed As Boolean, yieldFailed As Boolean, ok As Boolean
    Dim e As Double, qt As Double, qh As Double, ot As Double, sl As Double, bc As Double
    Dim qmax As Double, qmin As Double, status As String
    On Error GoTo InvalidAudit
    result = "PER-CHECK AUDIT; no EIT compliance claim" & vbCrLf
    result = result & "H=" & H & "; H1=" & H1 & "; gamma_soil=" & gamma_soil & "; gamma_concrete=" & gamma_concrete & "; phi=" & phi & "; mu=" & mu & "; qa=" & qa & vbCrLf
    result = result & "tt=" & d.tt & "; tb=" & d.tb & "; TBase=" & d.TBase & "; Base=" & d.Base & "; LToe=" & d.LToe & "; LHeel=" & d.LHeel & vbCrLf
    result = result & "fc_prime=" & currentMaterial.fc & "; fy=" & currentMaterial.fy & "; cover=" & cover & "; passive fraction=" & PassiveFactor & vbCrLf
    result = result & vbCrLf & "| Check | Actual | Limit / rule | Comparison | Basis / verification |" & vbCrLf
    result = result & "| --- | --- | --- | --- | --- |" & vbCrLf
    If Not GeometryOK(d) Then Err.Raise 5, , "INVALID_GEOMETRY_OR_INPUT"
    If d.UseDoubleStem Or d.UseDoubleToe Or d.UseDoubleHeel Then Err.Raise 5, , "UNSUPPORTED_DOUBLE_LAYER"
    result = result & AuditRow("Stem height", Format$(H - d.TBase, "0.0000") & " m", "H-TBase", "CALCULATED", "H measured from base underside")
    result = result & AuditMember(d, 0, "Stem", d.tb, d.ASst_DB, d.ASst_Sp, failed, yieldFailed)
    If stemOnly Then
        result = result & AuditRow("Toe / heel / wall stability", "Not evaluated", "Full candidate required", "OUT_OF_SCOPE", "Reference specifies stem steel only")
    Else
        result = result & AuditCompare("Heel length", d.LHeel, 0.3, True, "m", "Existing project geometry rule", failed)
        If Not CheckHeelLayout(d) Then
            failed = True
            result = result & AuditRow("Heel/toe", CStr(d.LHeel), "heel > toe", "FAIL_LISTED_LIMIT", "Existing project geometry rule")
        End If
        ok = CheckFS_OT(d, ot): ok = CheckFS_SL(d, sl)
        result = result & AuditCompare("Overturning FS", ot, FS_OT_MIN, True, "", "Existing project stability criterion", failed)
        result = result & AuditCompare("Sliding FS", sl, FS_SL_MIN, True, "", "Existing project stability criterion", failed)
        If BearingEdges(d, e, qt, qh) Then
            result = result & AuditCompare("Full contact abs(e)", Abs(e), d.Base / 6#, False, "m", "Linear full-compression model; signed e=" & Format$(e, "0.0000") & " positive to toe", failed)
            result = result & AuditRow("q_toe / q_heel", Format$(qt, "0.0000") & " / " & Format$(qh, "0.0000") & " tf/m2", "signed pressure diagram", "CALCULATED", "BearingEdges")
            ok = CheckFS_BC(d, bc, e, qmax, qmin)
            result = result & AuditCompare("Bearing qa/qmax", bc, FS_BC_MIN, True, "", "qa INPUT interpreted as ALLOWABLE", failed)
            result = result & AuditMember(d, 1, "Toe", d.TBase, d.AStoe_DB, d.AStoe_Sp, failed, yieldFailed)
            result = result & AuditMember(d, 2, "Heel", d.TBase, d.ASheel_DB, d.ASheel_Sp, failed, yieldFailed)
        Else
            failed = True
            result = result & AuditRow("Base contact", "Unsupported pressure diagram", "abs(e)<=B/6", "FAIL_MODEL", "Toe/heel forces NOT_EVALUATED; no full-width extrapolation")
        End If
    End If
    result = result & AuditRow("Complete EIT 011007-19 compliance", "Not established", "Verified applicable clauses and complete detailing", "UNVERIFIED", "See audit/WSD_PRIMARY_SOURCE_SEARCH.md")
    result = result & AuditRow("Cover, bar order, spacing, anchorage, distribution steel", "Not fully checked", "Applicable detailing provisions", "UNVERIFIED", "Current d assumes main bar outermost; no interposed transverse bar")
    status = "INDETERMINATE_WSD"
    If failed Then status = "FAIL_LISTED_CHECKS; EIT compliance still UNVERIFIED"
    If yieldFailed Then status = "FAIL_NECESSARY_YIELD_BOUND; EIT compliance still UNVERIFIED"
    BuildDesignCheckReport = "AUDIT RESULT: " & status & vbCrLf & result
    Exit Function
InvalidAudit:
    BuildDesignCheckReport = "AUDIT RESULT: INVALID_DATA; " & Err.Description & vbCrLf & result
End Function

' Stable comparison for dimensions represented by binary floating point.
' The existing project rule is heel >= 0.3 m and strictly greater than toe.
Public Function CheckHeelLayout(d As Design) As Boolean
    Const dimensionTolerance As Double = 0.000000001
    CheckHeelLayout = (d.LHeel >= 0.3 - dimensionTolerance And d.LHeel - d.LToe > dimensionTolerance)
End Function

' Nominal shear envelope, distinct from the still-unverified EIT critical section.
Public Function MemberShearStress(d As Design, part As Integer, depth As Double) As Double
    Dim criticalHeight As Double
    If part = 0 Then
        MemberShearStress = StemShearStressEnvelope(d, depth, criticalHeight)
    Else
        MemberShearStress = SectionShear(d, part, depth) / (10# * depth)
    End If
End Function

Public Function StemShearStressEnvelope(d As Design, baseDepth As Double, ByRef criticalHeight As Double) As Double
    Dim hs As Double, hp As Double, slope As Double, peak As Double
    Dim a As Double, b As Double, c As Double, region As Integer, lower As Double, upper As Double
    Dim qaRoot As Double, qbRoot As Double, qcRoot As Double, discriminant As Double
    If Not GeometryOK(d) Then Err.Raise 5, , "Invalid geometry for stem shear envelope"
    hs = H - d.TBase: hp = H1 - d.TBase: slope = (d.tb - d.tt) / hs
    If baseDepth <= 0 Or baseDepth - slope * hs <= 0 Then Err.Raise 5, , "Invalid local effective depth"
    criticalHeight = 0
    For region = 0 To 1
        a = gamma_soil * CalculateKa() / 2#
        b = -gamma_soil * CalculateKa() * hs
        c = gamma_soil * CalculateKa() * hs ^ 2 / 2#
        lower = hp: upper = hs
        If region = 0 Then
            lower = 0: upper = hp
            a = a - PassiveFactor * gamma_soil * CalculateKp() / 2#
            b = b + PassiveFactor * gamma_soil * CalculateKp() * hp
            c = c - PassiveFactor * gamma_soil * CalculateKp() * hp ^ 2 / 2#
        End If
        Call ConsiderStemShearPoint(lower, lower, upper, a, b, c, slope, baseDepth, peak, criticalHeight)
        Call ConsiderStemShearPoint(upper, lower, upper, a, b, c, slope, baseDepth, peak, criticalHeight)
        ' d[V/(d0-s*y)]/dy = 0 gives this quadratic. Evaluate both soil regions.
        qaRoot = -slope * a: qbRoot = 2# * a * baseDepth: qcRoot = b * baseDepth + slope * c
        If Abs(qaRoot) < 0.00000000000001 Then
            If Abs(qbRoot) > 0.00000000000001 Then Call ConsiderStemShearPoint(-qcRoot / qbRoot, lower, upper, a, b, c, slope, baseDepth, peak, criticalHeight)
        Else
            discriminant = qbRoot ^ 2 - 4# * qaRoot * qcRoot
            If discriminant >= 0 Then
                Call ConsiderStemShearPoint((-qbRoot + Sqr(discriminant)) / (2# * qaRoot), lower, upper, a, b, c, slope, baseDepth, peak, criticalHeight)
                Call ConsiderStemShearPoint((-qbRoot - Sqr(discriminant)) / (2# * qaRoot), lower, upper, a, b, c, slope, baseDepth, peak, criticalHeight)
            End If
        End If
    Next region
    StemShearStressEnvelope = peak
End Function

Private Sub ConsiderStemShearPoint(y As Double, lower As Double, upper As Double, a As Double, b As Double, c As Double, slope As Double, baseDepth As Double, ByRef peak As Double, ByRef criticalHeight As Double)
    Dim value As Double
    If y < lower Or y > upper Then Exit Sub
    value = Abs(a * y ^ 2 + b * y + c) / (10# * (baseDepth - slope * y))
    If value > peak Then peak = value: criticalHeight = y
End Sub
