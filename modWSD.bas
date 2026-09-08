Attribute VB_Name = "modWSD"
Option Explicit
' No guessed EIT clauses: these remain unset until reviewed against 011007-19.
Public AllowableShear As Double   ' kgf/cm2, for this material and member
Public MinStemRatio As Double    ' main vertical steel / gross concrete area
Public MinBaseRatio As Double    ' main slab steel / gross concrete area
Public WSDSource As String       ' edition, clause/page and applicability
Public WSDReviewed As Boolean
Public Const PROJECT_MODULAR_RATIO As Double = 9#  ' User-selected design assumption

' ========================================
' Working Stress Design (WSD) Module
' RC_RT_HCA v2.0
' Legacy parameters, NOT verified against Thai Standard (วสท. 2562)
' ========================================

' WSD Parameters Structure
Public Type WSDParams
    n As Double          ' Modular ratio adopted for this project: n=9
    fs As Double         ' Steel working stress (ksc): 1500 or 1700
    fc As Double         ' Concrete working stress (ksc): 0.45 × f'c
    k As Double          ' Neutral axis factor
    j As Double          ' Lever arm factor
    R As Double          ' Flexural constant (ksc)
End Type

' ========================================
' Calculate WSD Parameters
' ========================================
Public Function CalculateWSDParameters(fy As Integer, fc_prime As Integer) As WSDParams
    Dim params As WSDParams
    
    If fc_prime <= 0 Or (fy <> 3000 And fy <> 4000) Then Err.Raise 5, , "Invalid/unsupported material"
    ' User explicitly selected n=9 as the project design assumption.
    ' Do not replace this choice with an automatic Es/Ec calculation.
    ' Balanced k/j/R still depend on material allowables; actual bars use As/d.
    params.n = PROJECT_MODULAR_RATIO
    
    ' Steel working stress (fs)
    ' SD30 (fy=3000): fs = 1500 ksc
    ' SD40 (fy=4000): fs = 1700 ksc
    If fy <= 3000 Then
        params.fs = 1500
    Else
        params.fs = 1700
    End If
    
    ' Concrete working stress
    params.fc = 0.45 * fc_prime
    
    ' Balanced allowable-stress section only; actual bars use SectionStresses.
    ' Neutral axis factor: k = 1 / (1 + (fs / (n × fc)))
    params.k = 1 / (1 + (params.fs / (params.n * params.fc)))
    
    ' Lever arm factor: j = 1 - (k/3)
    params.j = 1 - (params.k / 3)
    
    ' Flexural constant: R = 0.5 × fc × k × j
    params.R = 0.5 * params.fc * params.k * params.j
    
    CalculateWSDParameters = params
End Function

' ========================================
' Calculate Moment Capacity
' Mr = R × b × d²  (ton-m)
' ========================================
Public Function CalculateMomentCapacity(R As Double, d As Double) As Double
    Dim b As Double
    
    ' Width = 1.00 m (per meter width)
    b = 1#
    
    ' Mr = R × b × d² (ksc × m × m² = ton-m)
    ' Note: R is in ksc (kg/cm²), need to convert
    ' R (ksc) × 10 = ton/m²
    ' Mr = (R × 10) × b × d²
    CalculateMomentCapacity = (R * 10) * b * d * d
End Function

' ========================================
' Calculate Required Steel Area
' As = M / (fs × j × d)  (cm²)
' ========================================
Public Function CalculateAsRequired(M As Double, fs As Double, j As Double, d As Double) As Double
    Dim M_kg_cm As Double
    Dim d_cm As Double
    
    ' แปลงหน่วยให้ถูกต้อง
    M_kg_cm = M * 1000 * 100    ' ton-m → kg-cm
    d_cm = d * 100              ' m → cm
    
    ' ตรวจสอบค่า
    If M_kg_cm <= 0 Or d_cm <= 0 Or fs <= 0 Or j <= 0 Then
        CalculateAsRequired = 0
        Exit Function
    End If
    
    ' As = M(kg-cm) / (fs(kg/cm2) × j × d(cm))
    CalculateAsRequired = M_kg_cm / (fs * j * d_cm)
End Function

' ========================================
' Calculate Provided Steel Area
' As = (Area per bar) × (Number of bars in 1.00 m)
' ========================================
Public Function CalculateAsProvided(DB_index As Integer, SP_index As Integer) As Double
    Dim db_mm As Integer
    Dim db_cm As Double
    Dim spacing_m As Double
    Dim area_per_bar As Double
    Dim n_bars As Double
    
    ' เพิ่มการตรวจสอบขอบเขต
    If DB_index < 1 Or DB_index > UBound(DBArray) Then
        CalculateAsProvided = 0
        Exit Function
    End If
    
    ' FIXED: Changed validation from >5 to >10 (SPArray has 10 items)
    If SP_index < 1 Or SP_index > UBound(SPArray) Then
        CalculateAsProvided = 0
        Exit Function
    End If
    
    ' Get values from arrays
    db_mm = DBArray(DB_index)
    spacing_m = SPArray(SP_index)
    
    ' Convert diameter from mm to cm
    db_cm = db_mm / 10
    
    ' Area of one bar: A = π × (d/2)²
    area_per_bar = 3.14159265358979 * (db_cm / 2) * (db_cm / 2)
    
    ' Number of bars in 1.00 m width
    If spacing_m > 0 Then
        n_bars = 1# / spacing_m
    Else
        n_bars = 0
    End If
    
    ' Total provided area
    CalculateAsProvided = area_per_bar * n_bars
End Function

' ========================================
' Check Steel Adequacy
' Returns True if As_provided >= As_required
' ========================================
Public Function CheckSteelAdequacy(As_provided As Double, As_required As Double) As Boolean
    CheckSteelAdequacy = (As_provided >= As_required)
End Function

' ========================================
' Format Steel Check Result
' Returns formatted string with comparison
' ========================================
Public Function FormatSteelCheckResult(As_required As Double, As_provided As Double) As String
    Dim result As String
    
    result = "As_required: " & Format(As_required, "0.00") & " cm2" & vbCrLf
    result = result & "As_provided: " & Format(As_provided, "0.00") & " cm2 >= " & _
             Format(As_required, "0.00") & " cm2 --> " & IIf(As_provided >= As_required, "OK (area only)", "FAIL")
    
    FormatSteelCheckResult = result
End Function

' ========================================
' Calculate Effective Depth
' d = thickness - cover
' ========================================
Public Function CalculateEffectiveDepth(thickness As Double, cover As Double, Optional ByVal bar_mm As Double = 0) As Double
    If thickness <= 0 Or cover < 0 Or bar_mm <= 0 Then Err.Raise 5, , "Thickness, clear cover and actual bar diameter required"
    CalculateEffectiveDepth = thickness - cover - bar_mm / 2000#
    If CalculateEffectiveDepth <= 0 Then Err.Raise 5, , "Nonpositive effective depth"
End Function

' ========================================
' Get Minimum Steel Ratio (ρmin)
' As_min = 0.0015 × b × d  (for temperature/shrinkage)
' ========================================
Public Function CalculateMinimumSteel(b As Double, d As Double) As Double
    Err.Raise 5, , "Select sourced member-specific minimum steel on gross thickness"
End Function

' ========================================
' Format Steel Designation
' Returns string like "DB20 @ 15cm"
' ========================================
Public Function FormatSteelDesignation(DB_index As Integer, SP_index As Integer) As String
    Dim db_mm As Integer
    Dim spacing_cm As Double
    
    ' Bounds check
    If DB_index < 1 Or DB_index > UBound(DBArray) Then
        FormatSteelDesignation = "Invalid DB"
        Exit Function
    End If
    
    If SP_index < 1 Or SP_index > UBound(SPArray) Then
        FormatSteelDesignation = "Invalid Spacing"
        Exit Function
    End If
    
    ' Get values
    db_mm = DBArray(DB_index)
    spacing_cm = SPArray(SP_index) * 100  ' m → cm
    
    ' Format: "DB20 @ 15cm"
    FormatSteelDesignation = "DB" & db_mm & " @ " & Format(spacing_cm, "0.0") & " cm"
End Function

' ========================================
' Calculate Balanced Steel Ratio (ρ_balanced)
' Elastic balanced section (not a code maximum): ρ_b = 0.5 * (fc/fs) × k
' ========================================
Public Function CalculateRhoBalanced(wsd As WSDParams) As Double
    ' ρ_balanced = 0.5 * (fc/fs) × k
    If wsd.fs > 0 Then
        CalculateRhoBalanced = 0.5 * (wsd.fc / wsd.fs) * wsd.k
    Else
        CalculateRhoBalanced = 0
    End If
End Function

' ========================================
' Calculate Maximum Steel Ratio (ρ_max)
' ρ_max = 0.75 × ρ_balanced
' ========================================
Public Function CalculateRhoMax(wsd As WSDParams) As Double
    Err.Raise 5, , "No sourced WSD maximum-steel rule; 0.75*rho_balanced is not assumed"
End Function

' ========================================
' Calculate Maximum Steel Area (As_max)
' As_max = ρ_max × b × d  (cm²)
' ========================================
Public Function CalculateAsMax(wsd As WSDParams, b As Double, d As Double) As Double
    Dim rho_max As Double
    Dim b_cm As Double
    Dim d_cm As Double
    
    ' Convert to cm
    b_cm = b * 100  ' m → cm
    d_cm = d * 100  ' m → cm
    
    ' Get ρ_max
    rho_max = CalculateRhoMax(wsd)
    
    ' As_max = ρ_max × b × d
    CalculateAsMax = rho_max * b_cm * d_cm
End Function

' ========================================
' Check If Double Layer Needed
' Returns True if As_required > As_max (single layer)
' ========================================
Public Function CheckIfDoubleLayerNeeded(As_required As Double, wsd As WSDParams, _
                                         b As Double, d As Double) As Boolean
    Dim As_max As Double
    
    ' Calculate As_max for single layer
    As_max = CalculateAsMax(wsd, b, d)
    
    ' Check if double layer needed
    If As_required > As_max Then
        CheckIfDoubleLayerNeeded = True
    Else
        CheckIfDoubleLayerNeeded = False
    End If
End Function

' ========================================
' Calculate Required Steel for Double Layer
' Layer 1: As1 = As_max
' Layer 2: As2 = As_required - As_max
' d2 = d1 - 2.5cm - db (cm)
' Returns total As_required for validation
' ========================================
Public Function CalculateDoubleLayerSteel(As_required As Double, wsd As WSDParams, _
                                          d1 As Double, db_cm As Double, _
                                          ByRef As1 As Double, ByRef As2 As Double, _
                                          ByRef d2 As Double) As Double
    Dim As_max As Double
    Dim b As Double
    
    b = 1#  ' 1.00 m width
    
    ' Calculate As_max for layer 1
    As_max = CalculateAsMax(wsd, b, d1)
    
    ' Layer 1: As1 = As_max
    As1 = As_max
    
    ' Layer 2: As2 = As_required - As_max
    As2 = As_required - As_max
    If As2 < 0 Then As2 = 0
    
    ' Calculate d2 (effective depth for layer 2)
    ' d2 = d1 - 2.5cm - db
    d2 = d1 - 0.025 - (db_cm / 100)  ' Convert to meters
    If d2 <= 0 Then Err.Raise 5, , "Invalid second layer depth"  ' Minimum
    
    ' Return total As_required
    CalculateDoubleLayerSteel = As1 + As2
End Function



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

