"""Build a bounded native section-sizing study; production BA is unchanged.

The six geometric alternatives and common main-bar layouts are a transparent
manual sizing comparison, NOT a BA run, research batch or full WSD optimization.
"""
from pathlib import Path
P = Path(__file__).resolve().parent

def put(name, text):
    (P / name).write_bytes(text.replace('\n', '\r\n').encode('ascii'))

put('Sizing.vbp', '''Type=Exe
Module=SizingMain; SizingMain.bas
Module=modDataStructures; ..\\modDataStructures.bas
Module=modWSD; ..\\modWSD.bas
Module=modShared; ..\\modShared.bas
Startup="Sub Main"
Name="NativeSizing"
ExeName32="Sizing.exe"
''')
put('SizingMain.bas', '''Attribute VB_Name = "SizingMain"
Option Explicit

Public Sub Main()
    Dim d As Design, mat As MaterialProperties, i As Integer, db As Integer, sp As Integer
    Dim f As Integer, report As Integer, count As Long, eligible As Long
    Dim ot As Double, sl As Double, bc As Double, e As Double, qmax As Double, qmin As Double
    Dim ms As Double, mt As Double, mh As Double, ds As Double, dt As Double
    Dim ok As Boolean, rejected As Boolean, s As String, audit As String
    On Error GoTo Fatal
    InitializeArrays
    H = 5: H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4
    phi = 30: mu = 0.6: qa = 20: cover = 0.075: PassiveFactor = 0
    mat = GetSD40Material(320, 2601, 24): currentMaterial = mat
    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    WSDReviewed = False: WSDSource = ""
    AllowableShear = 0: MinStemRatio = 0: MinBaseRatio = 0
    f = FreeFile
    Open App.Path & "\\sizing-options.csv" For Output As #f
    Print #f, "option,tt,tb,TBase,Base,toe,heel,DB,spacing,stem_height,d,Mstem,Mtoe,Mheel,FSot,FSsl,FSbc,known_checks_only,complete_WSD"
    report = FreeFile
    Open App.Path & "\\sizing-vb6-report.md" For Output As #report
    Print #report, "# Actual VB6 main-steel sizing comparison"
    Print #report, "Six explicit geometry alternatives x 20 common main-bar layouts. Not a BA run or full WSD optimization."
    Print #report, "Common bars are a practical comparison choice, not a restriction in production BA. No synthetic criteria."
    Print #report, "Legacy flexural limits: n=9, fc=144, fs=1700 kgf/cm2. Code verification remains outstanding."
    For i = 1 To 6
        d.tt = 0.2: d.Base = 3: d.LToe = 0.6
        d.tb = 0.25 + 0.05 * i: d.TBase = d.tb
        If i >= 4 Then d.tt = 0.25
        If i = 6 Then d.tb = 0.6: d.TBase = 0.6: d.Base = 3.5
        d.LHeel = d.Base - d.LToe - d.tb
        For db = DB_MIN To DB_MAX
            For sp = SP_MIN To SP_MAX
                count = count + 1
                d.ASst_DB = db: d.AStoe_DB = db: d.ASheel_DB = db
                d.ASst_Sp = sp: d.AStoe_Sp = sp: d.ASheel_Sp = sp
                ds = SectionDepth(d.tb, db): dt = SectionDepth(d.TBase, db)
                ms = CalculateMomentStem(d): mt = CalculateMomentToe(d): mh = CalculateMomentHeel(d)
                ok = GeometryOK(d) And CheckHeelLayout(d)
                ok = CheckFS_OT(d, ot) And ok
                ok = CheckFS_SL(d, sl) And ok
                ok = CheckFS_BC(d, bc, e, qmax, qmin) And ok
                ok = ms >= 0 And mt >= 0 And mh >= 0 And ok
                ok = CheckSteelOK(ms, ds, db, sp) And ok
                ok = CheckSteelOK(mt, dt, db, sp) And ok
                ok = CheckSteelOK(mh, dt, db, sp) And ok
                If ok Then eligible = eligible + 1
                ' The real validator MUST continue refusing all these candidates.
                rejected = Not CheckDesignValid(d, db, sp, db, sp, db, sp, ot, sl, bc)
                If Not rejected Then Err.Raise 5, , "Unsourced candidate accepted"
                Call CheckFS_OT(d, ot): Call CheckFS_SL(d, sl)
                Call CheckFS_BC(d, bc, e, qmax, qmin)
                s = CStr(i) & "," & CsvNumber(d.tt) & "," & CsvNumber(d.tb) & "," & CsvNumber(d.TBase) & "," & CsvNumber(d.Base)
                s = s & "," & CsvNumber(d.LToe) & "," & CsvNumber(d.LHeel) & "," & WP_DB(db) & "," & CsvNumber(WP_SP(sp))
                s = s & "," & CsvNumber(H - d.TBase) & "," & CsvNumber(ds) & "," & CsvNumber(ms) & "," & CsvNumber(mt) & "," & CsvNumber(mh)
                s = s & "," & CsvNumber(ot) & "," & CsvNumber(sl) & "," & CsvNumber(bc) & "," & CStr(ok) & ",UNVERIFIED"
                Print #f, s
                ' Report both the practical example and a weaker steel layout
                ' at the SAME dimensions, demonstrating selection by calculation.
                If i = 4 And db = 102 And (sp = 111 Or sp = 112) Then
                    Print #report, ""
                    Print #report, "## Option 4: common DB20 @ " & WP_SP(sp) & " m"
                    audit = BuildDesignCheckReport(d)
                    Print #report, audit
                    If sp = 111 And Not ok Then Err.Raise 5, , "Expected example fails listed checks"
                    If sp = 112 And ok Then Err.Raise 5, , "Expected wider spacing incorrectly qualifies"
                End If
            Next sp
        Next db
    Next i
    If count <> 120 Or WSDCriteriaReady() Then Err.Raise 5, , "Count or criteria changed"
    Print #report, "Native completion: layouts=" & count & "; within known checks=" & eligible & "; accepted by full validator=0; WSDReviewed=False"
    Close #f: Close #report
    Exit Sub
Fatal:
    If report > 0 Then Print #report, "FATAL " & Err.Number & ": " & Err.Description
    If f > 0 Then Close #f
    If report > 0 Then Close #report
End Sub
''')
