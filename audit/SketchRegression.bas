Attribute VB_Name = "SketchRegression"
Option Explicit
Private f As Integer, failures As Long, checks As Long, folder As String

Private Sub Check(label As String, condition As Boolean)
    checks = checks + 1
    If Not condition Then failures = failures + 1
    Print #f, IIf(condition, "OK: ", "FAIL: ") & label
End Sub

Private Sub Snapshot(name As String)
    SavePicture frmBestDesign.picSketch.Image, folder & "\" & name & ".bmp"
    SaveWindowImage frmBestDesign.hWnd, folder & "\" & name & "-window.bmp"
    Print #f, name & " | " & frmBestDesign.Tag & " | " & frmBestDesign.picSketch.Tag
End Sub

Private Sub CheckSession(algorithm As String)
    Dim summary As Integer, line As String, row() As String, bestCost As Double
    Dim barParts() As String, barLabel As String
    Dim bestTrial As Long, bestEval As Long, bestFolder As String, cost As Double, ev As Long
    If algorithm = "BA" Then Form1.cmdBA.Value = True Else Form1.cmdRun.Value = True
    Check algorithm & " completed single trial", RunTrial = 1 And EvaluationCount = 64
    Check algorithm & " popup visible", frmBestDesign.Visible
    bestCost = 1E+30
    summary = FreeFile: Open ProjectTrialSummary For Input As #summary
    Line Input #summary, line
    Do Until EOF(summary)
        Line Input #summary, line: row = Split(line, ",")
        If Len(row(4)) > 0 Then
            cost = CDbl(row(4)): ev = CLng(row(5))
            If cost < bestCost Or (cost = bestCost And ev < bestEval) Then
                bestCost = cost: bestTrial = CLng(row(0)): bestEval = ev
                bestFolder = Replace$(row(6), Chr$(34), "")
            End If
        End If
    Loop
    Close #summary
    Check algorithm & " displays best trial from whole session", bestTrial > 0 And InStr(frmBestDesign.Tag, algorithm & "; trial=" & bestTrial & ";") = 1
    Check algorithm & " labels completed trial count", InStr(frmBestDesign.Tag, "; trials=" & Form1.txtTrials.Text & ";") > 0
    Check algorithm & " labels first best evaluation", InStr(frmBestDesign.Tag, "; evaluation=" & bestEval & ";") > 0
    If bestTrial > 0 Then
        summary = FreeFile: Open bestFolder & "\run.txt" For Input As #summary
        Do Until EOF(summary)
            Line Input #summary, line
            If Left$(line, 3) = "tt=" Then Exit Do
        Loop
        Close #summary
        Check algorithm & " picture dimensions equal winning report", Left$(line, 3) = "tt=" And Left$(line, Len(line) - 2) = Mid$(frmBestDesign.picSketch.Tag, InStr(frmBestDesign.picSketch.Tag, "tt="))
        Check algorithm & " price equals session best", InStr(frmBestDesign.Tag, "; price=" & bestCost & ";") > 0
        summary = FreeFile: Open bestFolder & "\run.txt" For Input As #summary
        Do Until EOF(summary)
            Line Input #summary, line
            If Left$(line, 5) = "Stem:" Or Left$(line, 4) = "Toe:" Or Left$(line, 5) = "Heel:" Then
                barParts = Split(Split(line, ";")(0), " @ ")
                barLabel = barParts(0) & " @ " & Format$(CDbl(Replace$(barParts(1), " m", "")), "0.00") & " m"
                Check algorithm & " bar label matches winning report to 2 decimals", InStr(frmBestDesign.Tag, barLabel) > 0
            End If
        Loop
        Close #summary
        Print #f, algorithm & " winning trial=" & bestTrial & "; final trial=" & RunTrial & "; best folder=" & bestFolder
    End If
    Snapshot "popup-" & algorithm
End Sub

Public Sub Main()
    Dim d As Design, bad As Design, height As Integer, i As Integer, expected As Single
    Dim window As Form, stillOpen As Boolean
    On Error GoTo Failed
    folder = Trim$(Replace$(Command$, Chr$(34), ""))
    f = FreeFile: Open folder & "\sketch-regression.txt" For Output As #f
    Load Form1: Form1.Show
    d.tt = 0.25: d.tb = 0.4: d.TBase = 0.45: d.LToe = 1: d.IsValid = True
    d.ASst_DB = 101: d.ASst_Sp = 110: d.AStoe_DB = 101: d.AStoe_Sp = 110: d.ASheel_DB = 101: d.ASheel_Sp = 110
    For height = 3 To 5
        d.Base = 0.5 * height + 0.5: d.LHeel = d.Base - d.LToe - d.tb
        frmBestDesign.ShowBest d, CDbl(height), 1.2, "H" & height & " FIXTURE", 1
        Check "H" & height & " displays dimensions", InStr(frmBestDesign.picSketch.Tag, "H=" & height) = 1 And InStr(frmBestDesign.picSketch.Tag, "B=" & d.Base) > 0
        Snapshot "popup-H" & height
    Next height
    Call SeedSearchRandom(12345): expected = Rnd
    Call SeedSearchRandom(12345)
    frmBestDesign.ShowBest d, 5, 1.2, "RNG CHECK", 1
    Check "drawing preserves optimizer random sequence", Rnd = expected
    Check "drawing leaves design unchanged", d.tt = 0.25 And d.Base = 3 And d.IsValid
    bad = d: bad.IsValid = False
    frmBestDesign.ShowBest bad, 5, 1.2, "NO SOLUTION", 0
    Check "no solution replaces old geometry", frmBestDesign.picSketch.Tag = "NO_SOLUTION"
    Snapshot "popup-no-solution"
    bad = d: bad.TBase = 5
    frmBestDesign.ShowBest bad, 5, 1.2, "INVALID", 0
    Check "invalid geometry is not drawn", frmBestDesign.picSketch.Tag = "NO_SOLUTION"
    Check "OK button is visible and supports Enter/Esc", frmBestDesign.cmdClose.Visible And frmBestDesign.cmdClose.Caption = "OK" And frmBestDesign.cmdClose.Default And frmBestDesign.cmdClose.Cancel
    Check "title bar has system Close", (GetWindowLong(frmBestDesign.hWnd, -16) And &H80000) <> 0
    frmBestDesign.cmdClose.Value = True
    For Each window In Forms
        If window.Name = "frmBestDesign" Then stillOpen = True
    Next window
    Check "OK button unloads popup", Not stillOpen
    frmBestDesign.ShowBest d, 5, 1.2, "CLOSE CHECK", 1
    Call SendMessage(frmBestDesign.hWnd, &H112, &HF060&, 0)
    stillOpen = False
    For Each window In Forms
        If window.Name = "frmBestDesign" Then stillOpen = True
    Next window
    Check "system Close command unloads popup", Not stillOpen
    Form1.txtH.Text = "5": Form1.txtH1.Text = "1.2": Form1.txtQa.Text = "30"
    Form1.txtMaxIter.Text = "64": Form1.txtTrials.Text = "1": Form1.txtSeed.Text = "20260908"
    For i = 0 To Form1.cboConcreteStrength.ListCount - 1
        If Form1.cboConcreteStrength.List(i) = "320" Then Form1.cboConcreteStrength.ListIndex = i
    Next i
    CheckSession "BA"
    Unload Form1
    stillOpen = False
    For Each window In Forms
        If window.Name = "frmBestDesign" Then stillOpen = True
    Next window
    Check "closing owner also closes popup", Not stillOpen
    Print #f, "SKETCH checks=" & checks & "; failures=" & failures
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL " & Err.Number & ": " & Err.Description
    Close #f
    Unload Form1
End Sub
