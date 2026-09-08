"""Native H5 recalculation, independent reinforcement per member.

Production BA retains all feasibility gates. The separate finite-grid study
is explicitly a relaxed diagnostic while EIT criteria remain unverified.
"""
from pathlib import Path
import re
import math
import json
P = Path(__file__).resolve().parent
form = (P.parent/'Form1.frm').read_bytes().decode('latin1')
defaults = {}
for name, body in re.findall(r'Begin VB.TextBox (\w+)(.*?)\bEnd\b', form, re.S):
    value = re.search(r'Text\s*=\s*"([^"]*)"', body)
    if value:
        defaults[name] = float(value[1])
inputs = dict(H=5.0, H1=defaults['txtH1'], gamma_soil=defaults['txtGammaSoil'],
              gamma_concrete=defaults['txtGammaCon'], phi=defaults['txtPhi'], mu=defaults['txtMu'],
              qa_allowable=defaults['txtQa'], cover=defaults['txtCover']/100,
              seed=int(defaults['txtSeed']), budget=int(defaults['txtMaxIter']), fc=320, fy=4000, passive_factor=1, modular_ratio_model="Es/Ec", n=2040000/(15100*math.sqrt(320)))
(P/'full-sizing-inputs.json').write_text(json.dumps(dict(source='Form1.frm design-time TextBox values; H=5 and fc=320 retained from requested study; one project trial with active and full passive, not the default research trial count', inputs=inputs),indent=2),encoding='ascii')
def put(name, text):
    for key,value in inputs.items():
        text = text.replace('__'+key+'__', str(value))
    (P/name).write_bytes(text.replace('\n', '\r\n').encode('ascii'))
put('FullSizing.vbp', '''Type=Exe
Module=FullSizingMain; FullSizingMain.bas
Module=modDataStructures; ..\\modDataStructures.bas
Module=modWSD; ..\\modWSD.bas
Module=modProjectChecks; ..\\modProjectChecks.bas
Module=modShared; ..\\modShared.bas
Module=modBA; ..\\modBA.bas
Startup="Sub Main"
Name="NativeFullSizing"
ExeName32="FullSizing.exe"
''')
put('FullSizingMain.bas', '''Attribute VB_Name = "FullSizingMain"
Option Explicit
Private steelChecks As Long

Private Function SelectFlexuralBars(moment As Double, thickness As Double, ByRef bar As Integer, ByRef spacing As Integer) As Boolean
    Dim db As Integer, sp As Integer, area As Double, bestArea As Double, depth As Double
    bestArea = 1E+30
    For db = DB_MIN To DB_MAX
        depth = SectionDepth(thickness, db)
        For sp = SP_MIN To SP_MAX
            steelChecks = steelChecks + 1
            If CheckSteelOK(moment, depth, db, sp) Then
                area = CalculateAsProv(db, sp)
                If area < bestArea Then
                    bestArea = area: bar = db: spacing = sp
                    SelectFlexuralBars = True
                End If
            End If
        Next sp
    Next db
End Function

Private Function Fields(d As Design) As String
    Fields = CsvNumber(d.tt) & "," & CsvNumber(d.tb) & "," & CsvNumber(d.TBase) & "," & CsvNumber(d.Base) & "," & CsvNumber(d.LToe) & "," & CsvNumber(d.LHeel) & "," & _
        WP_DB(d.ASst_DB) & "," & CsvNumber(WP_SP(d.ASst_Sp)) & "," & WP_DB(d.AStoe_DB) & "," & CsvNumber(WP_SP(d.AStoe_Sp)) & "," & WP_DB(d.ASheel_DB) & "," & CsvNumber(WP_SP(d.ASheel_Sp))
End Function

Private Sub WriteBarOptions(d As Design, eta As Integer, output As Integer)
    Dim part As Integer, db As Integer, sp As Integer, depth As Double, thickness As Double, moment As Double
    Dim c As Double, st As Double, j As Double, area As Double, selected As Boolean, ok As Boolean
    For part = 0 To 2
        Select Case part
            Case 0: thickness = d.tb: moment = CalculateMomentStem(d)
            Case 1: thickness = d.TBase: moment = CalculateMomentToe(d)
            Case 2: thickness = d.TBase: moment = CalculateMomentHeel(d)
        End Select
        For db = DB_MIN To DB_MAX
            depth = SectionDepth(thickness, db)
            For sp = SP_MIN To SP_MAX
                area = CalculateAsProv(db, sp)
                ok = SectionStresses(moment, depth, area, currentWSD.n, c, st, j)
                selected = False
                If part = 0 Then selected = db = d.ASst_DB And sp = d.ASst_Sp
                If part = 1 Then selected = db = d.AStoe_DB And sp = d.AStoe_Sp
                If part = 2 Then selected = db = d.ASheel_DB And sp = d.ASheel_Sp
                Print #output, eta & "," & part & "," & WP_DB(db) & "," & CsvNumber(WP_SP(sp)) & "," & CsvNumber(depth) & "," & CsvNumber(area) & "," & CsvNumber(moment) & "," & CsvNumber(c) & "," & CsvNumber(st) & "," & CStr(CheckSteelOK(moment, depth, db, sp)) & "," & CStr(selected)
            Next sp
        Next db
    Next part
End Sub

Public Sub Main()
    Dim report As Integer, trace As Integer, options As Integer, summary As Integer, flagFile As Integer
    Dim flags As String
    Dim eta As Integer, it As Integer, ib As Integer, iz As Integer, iw As Integer, il As Integer
    Dim d As Design, best As Design, emptyDesign As Design, ba As Design, mat As MaterialProperties
    Dim rows As Long, stable As Long, screened As Long, bestRow As Long
    Dim ot As Double, sl As Double, bc As Double, e As Double, qmax As Double, qmin As Double
    Dim ms As Double, mt As Double, mh As Double, estimate As Double, bestEstimate As Double, rejected As Boolean
    On Error GoTo Fatal
    InitializeArrays
    H = __H__: H1 = __H1__: gamma_soil = __gamma_soil__: gamma_concrete = __gamma_concrete__
    phi = __phi__: mu = __mu__: qa = __qa_allowable__: cover = __cover__
    mat = GetSD40Material(320, GetConcretePrice(320), 24): currentMaterial = mat
    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    WSDReviewed = False: WSDSource = "": AllowableShear = 0: MinStemRatio = 0: MinBaseRatio = 0
    report = FreeFile: Open App.Path & "\\full-sizing-vb6.md" For Output As #report
    trace = FreeFile: Open App.Path & "\\full-sizing-improvements.csv" For Output As #trace
    options = FreeFile: Open App.Path & "\\full-sizing-bar-options.csv" For Output As #options
    summary = FreeFile: Open App.Path & "\\full-sizing-summary.csv" For Output As #summary
    Print #report, "# Actual VB6 H5 recalculation: independent stem/toe/heel reinforcement"
    Print #report, "Other inputs read from Form1.frm defaults; H=5 and fc=320 retained study. See full-sizing-inputs.json. One project trial with active and full passive."
    Print #report, "Project load case: active from H and full Rankine passive from front H1, as specified by the user."
    Print #report, "Production BA requires all criteria. Separate relaxed grid: stability + supported bending + legacy stress only; NOT an accepted EIT design."
    Print #report, "Missing minimum/shear/detailing are NOT silently passed. No synthetic limits. Concrete unit price=" & mat.concretePrice & "; steel=" & mat.SteelPrice
    Print #report, "Cost model includes concrete and listed main bars, with original +0.4 m bar length allowances; not a complete construction BOQ."
    Print #trace, "eta,geometry_row,screening_estimate,tt,tb,TBase,Base,toe,heel,stemDB,stemSP,toeDB,toeSP,heelDB,heelSP"
    Print #summary, "eta,geometry_rows,stable_rows,screened_rows,steel_checks,best_row,screening_estimate,tt,tb,TBase,Base,toe,heel,stemDB,stemSP,toeDB,toeSP,heelDB,heelSP,qa_allowable"
    Print #options, "eta,part,DB,spacing,depth,As,moment,fc,fs,legacy_flexure,selected"
    For eta = 1 To 1
        PassiveFactor = eta
        Print #report, ""
        Print #report, "## Passive fraction=" & eta
        ba = BisectionOptimization(__budget__, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, mat, RandomSeed:=__seed__, TrialNumber:=1)
        Print #report, "Actual production BA: " & RunStatus & "; evaluations=" & EvaluationCount & "; best evaluation=" & RunBestEvaluation & "; folder=" & RunFolder
        If ba.IsValid Or EvaluationCount <> __budget__ Or WSDCriteriaReady() Then Err.Raise 5, , "Unexpected full acceptance or BA budget"
        best = emptyDesign: bestEstimate = NO_SOLUTION_COST: rows = 0: stable = 0: screened = 0: steelChecks = 0
        flags = String$(520200, "0")
        For it = TT_MIN To TT_MAX
        For ib = TB_MIN To tb_max
        For iz = TBASE_MIN To TBase_max
        For iw = BASE_MIN To BASE_MAX
        For il = LTOE_MIN To LTOE_MAX
            rows = rows + 1
            d.tt = WP_tt(it): d.tb = WP_tb(ib): d.TBase = WP_TBase(iz)
            d.Base = WP_Base(iw): d.LToe = WP_LToe(il): d.LHeel = d.Base - d.LToe - d.tb
            If Not GeometryOK(d) Then GoTo NextGeometry
            If Not CheckHeelLayout(d) Then GoTo NextGeometry
            If Not CheckFS_OT(d, ot) Then GoTo NextGeometry
            If Not CheckFS_SL(d, sl) Then GoTo NextGeometry
            If Not CheckFS_BC(d, bc, e, qmax, qmin) Then GoTo NextGeometry
            stable = stable + 1
            Mid$(flags, rows, 1) = "1"
            ms = CalculateMomentStem(d): mt = CalculateMomentToe(d): mh = CalculateMomentHeel(d)
            If ms < 0 Or mt < 0 Or mh < 0 Then GoTo NextGeometry
            If Not SelectFlexuralBars(ms, d.tb, d.ASst_DB, d.ASst_Sp) Then GoTo NextGeometry
            If Not SelectFlexuralBars(mt, d.TBase, d.AStoe_DB, d.AStoe_Sp) Then GoTo NextGeometry
            If Not SelectFlexuralBars(mh, d.TBase, d.ASheel_DB, d.ASheel_Sp) Then GoTo NextGeometry
            screened = screened + 1
            Mid$(flags, rows, 1) = "2"
            ' Diagnostic relaxed estimate only; NEVER update production RunBest.
            estimate = CalculateCost(d)
            If estimate < bestEstimate Then
                best = d: bestEstimate = estimate: bestRow = rows
                Print #trace, eta & "," & rows & "," & CsvNumber(estimate) & "," & Fields(d)
            End If
NextGeometry:
        Next il
        Next iw
        Next iz
        Next ib
        Next it
        flagFile = FreeFile
        Open App.Path & "\\full-sizing-flags-" & eta & ".bin" For Binary As #flagFile
        Put #flagFile, , flags
        Close #flagFile
        If rows <> 520200 Or screened = 0 Then Err.Raise 5, , "Incomplete grid"
        rejected = Not CheckDesignValid(best, best.ASst_DB, best.ASst_Sp, best.AStoe_DB, best.AStoe_Sp, best.ASheel_DB, best.ASheel_Sp, ot, sl, bc)
        If Not rejected Or RunBest.IsValid Then Err.Raise 5, , "Diagnostic candidate promoted to verified design"
        Print #summary, eta & "," & rows & "," & stable & "," & screened & "," & steelChecks & "," & bestRow & "," & CsvNumber(bestEstimate) & "," & Fields(best) & "," & CsvNumber(qa)
        Print #report, "Relaxed grid: rows=" & rows & "; all 3 stability=" & stable & "; screened=" & screened & "; steel checks=" & steelChecks
        Print #report, "Lowest relaxed grid estimate=" & CsvNumber(bestEstimate) & " baht/m; complete validator=" & LastValidationReason
        Print #report, BuildDesignCheckReport(best)
        Call WriteBarOptions(best, eta, options)
    Next eta
    Print #report, "NATIVE COMPLETE: one project grid completed; production accepted designs=0; WSDReviewed=False"
    Close #summary: Close #options: Close #trace: Close #report
    Exit Sub
Fatal:
    If report > 0 Then Print #report, "FATAL " & Err.Number & ": " & Err.Description
    Close
End Sub
''')
