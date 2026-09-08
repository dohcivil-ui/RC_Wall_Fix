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

' v2.8: Batch experiment flag - suppresses per-iteration CSV writes when True
Public BatchMode As Boolean

'================================================================================
' SECTION 1: Constants
'================================================================================
Public Const PI As Double = 3.14159265358979
Public Const DEG_TO_RAD As Double = 1.74532925199433E-02

' Safety Factor Requirements
Public Const FS_OT_MIN As Double = 2#      ' Overturning
Public Const FS_SL_MIN As Double = 1.5     ' Sliding
Public Const FS_BC_MIN As Double = 2#      ' Bearing Capacity

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
' Get Concrete Price by f'c (Maha Sarakham Province)
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
    CalculatePp = 0.5 * gamma_soil * Kp * H1 * H1
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
    If H1_toe < 0 Then H1_toe = 0
    
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
    x1 = d.LToe / 2#
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
Public Function CalculateMomentStem() As Double
    Dim Ka As Double
    Ka = CalculateKa()
    CalculateMomentStem = 0.5 * gamma_soil * Ka * H1 ^ 3 / 3#
End Function

'--------------------------------------------------------------------------------
' Moment at Toe (using Bearing Pressure)
'--------------------------------------------------------------------------------
Public Function CalculateMomentToe(d As Design) As Double
    Dim W_total As Double, W1 As Double, W2 As Double, W3 As Double, W4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim MR As Double, MO As Double
    Dim e As Double, q_max As Double, q_min As Double
    Dim q_toe As Double, q_junction As Double
    Dim M_bearing As Double, M_self As Double, w_self As Double, L_eff As Double
    
    W_total = CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    MR = W1 * x1 + W2 * x2 + W3 * x3 + W4 * x4
    MO = CalculateMO(d)
    
    If W_total > 0.1 And d.Base > 0.1 Then
        e = Abs((d.Base / 2#) - ((MR - MO) / W_total))
        If e <= d.Base / 6# Then
            q_max = (W_total / d.Base) * (1# + (6# * e / d.Base))
            q_min = (W_total / d.Base) * (1# - (6# * e / d.Base))
        Else
            L_eff = 3# * (d.Base / 2# - e)
            If L_eff > 0.1 Then
                q_max = 2# * W_total / L_eff
            Else
                q_max = W_total / d.Base * 2#
            End If
            q_min = 0
        End If
    Else
        q_max = 1#: q_min = 1#
    End If
    
    q_toe = q_max
    If d.Base > 0.01 Then
        q_junction = q_max - (q_max - q_min) * (d.LToe / d.Base)
    Else
        q_junction = q_max
    End If
    M_bearing = q_junction * d.LToe * (d.LToe / 2#) + _
                (q_toe - q_junction) * (d.LToe / 2#) * (2# * d.LToe / 3#)
    w_self = gamma_concrete * d.TBase
    M_self = w_self * d.LToe * d.LToe / 2#
    
    CalculateMomentToe = Abs(M_bearing - M_self)
End Function

'--------------------------------------------------------------------------------
' Moment at Heel (using Bearing Pressure)
'--------------------------------------------------------------------------------
Public Function CalculateMomentHeel(d As Design) As Double
    Dim W_total As Double, W1 As Double, W2 As Double, W3 As Double, W4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim MR As Double, MO As Double
    Dim e As Double, q_max As Double, q_min As Double
    Dim q_junction As Double, q_heel As Double
    Dim M_downward As Double, M_bearing As Double
    Dim H_soil As Double, w_down As Double, L_eff As Double
    
    W_total = CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    MR = W1 * x1 + W2 * x2 + W3 * x3 + W4 * x4
    MO = CalculateMO(d)
    
    If W_total > 0.1 And d.Base > 0.1 Then
        e = Abs((d.Base / 2#) - ((MR - MO) / W_total))
        If e <= d.Base / 6# Then
            q_max = (W_total / d.Base) * (1# + (6# * e / d.Base))
            q_min = (W_total / d.Base) * (1# - (6# * e / d.Base))
        Else
            L_eff = 3# * (d.Base / 2# - e)
            If L_eff > 0.1 Then
                q_max = 2# * W_total / L_eff
            Else
                q_max = W_total / d.Base * 2#
            End If
            q_min = 0
        End If
    Else
        q_max = 1#: q_min = 1#
    End If
    
    If d.Base > 0.01 Then
        q_junction = q_max - (q_max - q_min) * ((d.LToe + d.tb) / d.Base)
    Else
        q_junction = q_max
    End If
    q_heel = q_min
    H_soil = H - d.TBase
    w_down = (d.TBase * gamma_concrete) + (H_soil * gamma_soil)
    M_downward = w_down * d.LHeel * d.LHeel / 2#
    M_bearing = q_heel * d.LHeel * (d.LHeel / 2#) + _
                (q_junction - q_heel) * (d.LHeel / 2#) * (d.LHeel / 3#)
    
    CalculateMomentHeel = Abs(M_downward - M_bearing)
End Function

'================================================================================
' SECTION 10: Safety Factor Checks
'================================================================================

'--------------------------------------------------------------------------------
' Check Overturning (FS_OT >= 2.0)
'--------------------------------------------------------------------------------
Public Function CheckFS_OT(d As Design, ByRef FS_OT As Double) As Boolean
    Dim MR As Double, MO As Double
    MR = CalculateMR(d)
    MO = CalculateMO(d)
    
    If MO <= 0.001 Then
        FS_OT = 999
        CheckFS_OT = True
        Exit Function
    End If
    
    FS_OT = MR / MO
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
' Check Bearing Capacity (FS_BC >= 2.0)
'--------------------------------------------------------------------------------
Public Function CheckFS_BC(d As Design, ByRef FS_BC As Double, _
                           ByRef e As Double, ByRef q_max As Double, _
                           ByRef q_min As Double) As Boolean
    Dim W_total As Double, W1 As Double, W2 As Double, W3 As Double, W4 As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim MR As Double, MO As Double, L_eff As Double
    
    W_total = CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    
    If W_total < 0.1 Then
        FS_BC = 0
        CheckFS_BC = False
        Exit Function
    End If
    
    If d.Base < 0.1 Then
        FS_BC = 0
        CheckFS_BC = False
        Exit Function
    End If
    
    MR = CalculateMR(d)
    MO = CalculateMO(d)
    
    e = Abs((d.Base / 2#) - ((MR - MO) / W_total))
    
    If e > d.Base / 3# Then
        FS_BC = 0
        CheckFS_BC = False
        Exit Function
    End If
    
    If e <= d.Base / 6# Then
        q_max = (W_total / d.Base) * (1# + (6# * e / d.Base))
        q_min = (W_total / d.Base) * (1# - (6# * e / d.Base))
    Else
        L_eff = 3# * (d.Base / 2# - e)
        If L_eff > 0.01 Then
            q_max = 2# * W_total / L_eff
        Else
            q_max = W_total / d.Base * 2#
        End If
        q_min = 0
    End If
    
    If q_min < 0 Or q_max <= 0.001 Then
        FS_BC = 0
        CheckFS_BC = False
        Exit Function
    End If
    
    FS_BC = qa / q_max
    CheckFS_BC = (FS_BC >= FS_BC_MIN)
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
    Dim As_req As Double, As_prov As Double
    Dim M_kg_cm As Double, d_cm As Double
    
    ' Calculate As_required
    M_kg_cm = M * 100000   ' ton-m to kg-cm
    d_cm = d_eff * 100     ' m to cm
    If d_cm < 1 Then d_cm = 1
    As_req = M_kg_cm / (currentWSD.fs * currentWSD.j * d_cm)
    
    ' Calculate As_provided
    As_prov = CalculateAsProv(DB_idx, SP_idx)
    
    CheckSteelOK = (As_prov >= As_req)
End Function

'================================================================================
' SECTION 12: Cost Calculation
'================================================================================

'--------------------------------------------------------------------------------
' CalculateCost - รับแค่ Design (ดึงค่าเหล็กจาก Design.ASst_DB ฯลฯ)
'--------------------------------------------------------------------------------
Public Function CalculateCost(d As Design) As Double
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
                                 ByRef FS_OT As Double, _
                                 ByRef FS_SL As Double, _
                                 ByRef FS_BC As Double) As Boolean
    
    Dim e As Double, q_max As Double, q_min As Double
    Dim M_stem As Double, M_toe As Double, M_heel As Double
    Dim d_stem As Double, d_toe As Double, d_heel As Double
    
    CheckDesignValid = False
    
    ' === Constraint Checks ===
    
    ' tb >= tt
    If d.tb < d.tt Then Exit Function
    
    ' LHeel >= 0.3m
    If d.LHeel < 0.3 Then Exit Function
    
    ' LHeel > LToe
    If d.LHeel <= d.LToe Then Exit Function
    
    ' === Safety Factor Checks ===
    
    If Not CheckFS_OT(d, FS_OT) Then Exit Function
    If Not CheckFS_SL(d, FS_SL) Then Exit Function
    If Not CheckFS_BC(d, FS_BC, e, q_max, q_min) Then Exit Function
    
    ' === Steel Adequacy Checks ===
    
    M_stem = CalculateMomentStem()
    M_toe = CalculateMomentToe(d)
    M_heel = CalculateMomentHeel(d)
    
    d_stem = d.tb - cover
    d_toe = d.TBase - cover
    d_heel = d.TBase - cover
    
    If d_stem <= 0.05 Then d_stem = 0.05
    If d_toe <= 0.05 Then d_toe = 0.05
    If d_heel <= 0.05 Then d_heel = 0.05
    
    If Not CheckSteelOK(M_stem, d_stem, StemDB, StemSP) Then Exit Function
    If Not CheckSteelOK(M_toe, d_toe, ToeDB, ToeSP) Then Exit Function
    If Not CheckSteelOK(M_heel, d_heel, HeelDB, HeelSP) Then Exit Function
    
    ' === All Checks Passed ===
    CheckDesignValid = True
    
End Function

'================================================================================
' SECTION 14: Format Results for Display
'================================================================================

Public Function FormatResults(d As Design, mat As MaterialProperties, _
                              Optional AlgoName As String = "Hill Climbing v5.0") As String
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
    
    wsd = CalculateWSDParameters(mat.fy, mat.fc)
    
    Ka = CalculateKa()
    Kp = CalculateKp()
    Pa = CalculatePa()
    Pp = CalculatePp()
    
    W_total = CalculateWTotal(d, W1, W2, W3, W4, x1, x2, x3, x4)
    
    M_stem = CalculateMomentStem()
    M_toe = CalculateMomentToe(d)
    M_heel = CalculateMomentHeel(d)
    
    Call CheckFS_OT(d, FS_OT)
    Call CheckFS_SL(d, FS_SL)
    Call CheckFS_BC(d, FS_BC, e, q_max, q_min)
    
    cost = CalculateCost(d)
    
    ' คำนวณ As_req และ As_prov
    d_stem = d.tb - cover: If d_stem < 0.05 Then d_stem = 0.05
    d_toe = d.TBase - cover: If d_toe < 0.05 Then d_toe = 0.05
    d_heel = d.TBase - cover: If d_heel < 0.05 Then d_heel = 0.05
    
    As_req_stem = (M_stem * 100000) / (wsd.fs * wsd.j * (d_stem * 100))
    As_req_toe = (M_toe * 100000) / (wsd.fs * wsd.j * (d_toe * 100))
    As_req_heel = (M_heel * 100000) / (wsd.fs * wsd.j * (d_heel * 100))
    
    ' === MATERIAL PROPERTIES ===
    result = "=== MATERIAL PROPERTIES ===" & vbCrLf
    result = result & "Steel: " & mat.SteelGrade & " (fy=" & mat.fy & ", fs=" & wsd.fs & " ksc)" & vbCrLf
    result = result & "Concrete: f'c = " & mat.fc & " ksc (fc=" & Format(wsd.fc, "0.0") & " ksc)" & vbCrLf
    result = result & "WSD: n=" & wsd.n & ", k=" & Format(wsd.k, "0.000") & ", j=" & Format(wsd.j, "0.000") & ", R=" & Format(wsd.R, "0.00") & " ksc" & vbCrLf
    result = result & "Prices: Concrete=" & Format(mat.concretePrice, "#,##0") & " Baht/m3, Steel=" & mat.steelPrice & " Baht/kg" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === OPTIMIZATION RESULTS ===
    result = result & "=== OPTIMIZATION RESULTS ===" & vbCrLf
    result = result & "Algorithm: " & AlgoName & vbCrLf
    result = result & "Best Found at Trial: " & modDataStructures.BestTrial & vbCrLf
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
    If FS_OT >= 2# Then result = result & "PASS" Else result = result & "FAIL"
    result = result & vbCrLf
    
    result = result & "FS_Sliding: " & Format(FS_SL, "0.00") & " >= 1.5 "
    If FS_SL >= 1.5 Then result = result & "PASS" Else result = result & "FAIL"
    result = result & vbCrLf
    
    result = result & "FS_Bearing: " & Format(FS_BC, "0.00") & " >= 2.0 "
    If FS_BC >= 2# Then result = result & "PASS" Else result = result & "FAIL"
    result = result & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    ' === BEARING CAPACITY ===
    result = result & "=== BEARING CAPACITY ===" & vbCrLf
    result = result & "Eccentricity (e): " & Format(e, "0.000") & " m" & vbCrLf
    result = result & "q_max: " & Format(q_max, "0.00") & " ton/m2 (allowable: " & Format(qa, "0.00") & ")" & vbCrLf
    result = result & "q_min: " & Format(q_min, "0.00") & " ton/m2" & vbCrLf
    result = result & "-------------------------------" & vbCrLf
    
    FormatResults = result
End Function

'================================================================================
' END OF MODULE: modShared.bas
'================================================================================
