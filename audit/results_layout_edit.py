"""Presentation-only VB6 edit; preserve source bytes and CRLF."""
from pathlib import Path
P=Path(__file__).resolve().parent.parent
before=P/'audit'/'results-layout'/'before';before.mkdir(parents=True,exist_ok=True)
for name in ('Form1.frm','Form1.frx','modShared.bas'):
    assert not (before/name).exists()
    (before/name).write_bytes((P/name).read_bytes())

screen='''Private Function ScreenResultRow(label As String, value As String) As String
    ScreenResultRow = Left$(label & Space$(18), 18) & value & vbCrLf
End Function

Public Function FormatScreenResults(d As Design, mat As MaterialProperties, algorithm As String) As String
    Dim r As ProjectDetail, s As String, i As Integer, member As String
    If Not ProjectChecksEnabled Then
        FormatScreenResults = FormatResults(d, mat, algorithm)
        Exit Function
    End If
    s = "DESIGN RESULTS | " & algorithm & vbCrLf
    If Not d.IsValid Then
        FormatScreenResults = s & "NO_SOLUTION: no admissible design." & vbCrLf
        Exit Function
    End If
    currentMaterial = mat: currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    If Not CheckProjectDesign(d, r) Then
        FormatScreenResults = s & "Check failed: " & LastValidationReason & vbCrLf
        Exit Function
    End If
    s = s & "Status: PASS (configured project checks)" & vbCrLf
    s = s & ScreenResultRow("Cost", Format$(r.Cost, "#,##0.00") & " Baht/m")
    s = s & vbCrLf & "DIMENSIONS" & vbCrLf
    s = s & ScreenResultRow("H / H1", Format$(H, "0.00") & " / " & Format$(H1, "0.00") & " m")
    s = s & ScreenResultRow("tt / tb", Format$(d.tt, "0.00") & " / " & Format$(d.tb, "0.00") & " m")
    s = s & ScreenResultRow("TBase / Base", Format$(d.TBase, "0.00") & " / " & Format$(d.Base, "0.00") & " m")
    s = s & ScreenResultRow("Toe / Heel", Format$(d.LToe, "0.00") & " / " & Format$(d.LHeel, "0.00") & " m")
    s = s & ScreenResultRow("Stem height", Format$(H - d.TBase, "0.00") & " m")
    s = s & ScreenResultRow("Clear cover", Format$(cover, "0.000") & " m")
    s = s & vbCrLf & "STABILITY" & vbCrLf
    s = s & ScreenResultRow("FS overturning", Format$(r.OT, "0.00"))
    s = s & ScreenResultRow("FS sliding", Format$(r.SL, "0.00"))
    s = s & ScreenResultRow("Bearing qa/qmax", Format$(r.BC, "0.00"))
    s = s & vbCrLf & "MAIN REINFORCEMENT" & vbCrLf
    For i = 0 To 2
        member = "Stem (back)": If i = 1 Then member = "Toe (bottom)"
        If i = 2 Then member = "Heel (top)"
        s = s & ScreenResultRow(member, "DB" & WP_DB(r.DB(i)) & " @ " & Format$(WP_SP(r.SP(i)), "0.00") & " m")
    Next i
    s = s & vbCrLf & "MEMBER CHECKS - actual / allowable" & vbCrLf
    For i = 0 To 2
        member = "STEM": If i = 1 Then member = "TOE"
        If i = 2 Then member = "HEEL"
        s = s & member & vbCrLf
        s = s & ScreenResultRow("Moment", Format$(r.Moment(i), "0.00") & " tf.m/m")
        s = s & ScreenResultRow("As / minimum", Format$(r.Steel(i), "0.00") & " / " & Format$(r.Minimum(i), "0.00") & " cm2/m")
        s = s & ScreenResultRow("Effective depth", Format$(r.Depth(i), "0.000") & " m")
        s = s & ScreenResultRow("Concrete stress", Format$(r.FcBound(i), "0.00") & " / " & Format$(currentWSD.fc, "0.00") & " kgf/cm2")
        s = s & ScreenResultRow("Steel stress", Format$(r.FsBound(i), "0.00") & " / " & Format$(currentWSD.fs, "0.00") & " kgf/cm2")
        s = s & ScreenResultRow("Shear stress", Format$(r.Shear(i), "0.00") & " / " & Format$(r.ShearLimit, "0.00") & " kgf/cm2")
        s = s & vbCrLf
    Next i
    s = s & "MATERIAL QUANTITIES / WSD" & vbCrLf
    s = s & ScreenResultRow("Concrete", Format$(r.ConcreteVolume, "0.000") & " m3/m")
    s = s & ScreenResultRow("Main steel", Format$(r.MainWeight, "0.00") & " kg/m")
    s = s & ScreenResultRow("n = Es/Ec", Format$(currentWSD.n, "0.00"))
    s = s & vbCrLf & "NOTES" & vbCrLf
    s = s & "Full active and passive pressure." & vbCrLf
    s = s & "Concrete + main steel only; no secondary steel." & vbCrLf
    s = s & "No anchorage, laps or formwork included." & vbCrLf
    s = s & "Basis: " & PROJECT_CHECK_BASIS & vbCrLf
    s = s & "Full checks, limits and assumptions: run.txt" & vbCrLf
    FormatScreenResults = s
End Function

'''
p=P/'modShared.bas';b=p.read_bytes();pos=b.index(b'Public Function FormatResults(')
p.write_bytes(b[:pos]+screen.replace('\n','\r\n').encode('ascii')+b[pos:])
p=P/'Form1.frm';b=p.read_bytes()
start=b.index(b'      Begin VB.ListBox lstResults');end=b.index(b'      End\r\n',start)
part=b[start:end].replace(b'"CordiaUPC"',b'"Consolas"').replace(b'14.25',b'9.75').replace(b'Charset         =   222',b'Charset         =   0')
b=b[:start]+part+b[end:]
b=b.replace(b'modShared.FormatResults(bestDesign, selectedMaterial, "Bisection Algorithm v1.0")',b'modShared.FormatScreenResults(bestDesign, selectedMaterial, "BA")')
b=b.replace(b'FormatResults(bestDesign, selectedMaterial)',b'FormatScreenResults(bestDesign, selectedMaterial, "HCA")')
# Route all list output (including file paths and comparisons) through the wrapper.
b=b.replace(b'lstResults.AddItem ',b'AddResultLine ')
wrapper='''Public Sub AddResultLine(ByVal text As String)
    Const MAX_COLUMNS As Long = 48
    Dim cut As Long
    ' A ListBox does not wrap long rows. Preserve every character, including paths.
    Do While Len(text) > MAX_COLUMNS
        cut = InStrRev(Left$(text, MAX_COLUMNS), " ")
        If cut <= 1 Then cut = MAX_COLUMNS
        lstResults.AddItem Left$(text, cut)
        text = Mid$(text, cut + 1)
    Loop
    lstResults.AddItem text
End Sub

'''
pos=b.index(b'Private Sub Form_Load()')
b=b[:pos]+wrapper.replace('\n','\r\n').encode('ascii')+b[pos:]
p.write_bytes(b)
