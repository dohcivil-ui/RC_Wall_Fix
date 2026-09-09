Attribute VB_Name = "ResultsLayout"
Option Explicit
Private f As Integer, failures As Long, clipped As Long
Public Sub Main()
    Dim d As Design, mat As MaterialProperties, s As String, joined As String, start As Integer, i As Integer
    On Error GoTo Failed
    Load Form1
    Form1.Caption = "Results layout before - native VB6 test"
    Call InitializeArrays
    Call EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 30: cover = 0.075
    f = FreeFile
    Open App.Path & "\layout.txt" For Output As #f
    H = 5: mat = GetSD40Material(320, GetConcretePrice(320), 24)
    d.tt = 0.2000000000
    d.tb = 0.4000000000
    d.TBase = 0.3000000000
    d.Base = 3.0000000000
    d.LToe = 1.0000000000
    d.LHeel = d.Base - d.LToe - d.tb
    d.ASst_DB = 101
    d.ASst_Sp = 110
    d.AStoe_DB = 101
    d.AStoe_Sp = 112
    d.ASheel_DB = 103
    d.ASheel_Sp = 113
    d.IsValid = True
    Call ShowFixture(d, mat)
    H = 4: mat = GetSD40Material(280, GetConcretePrice(280), 24)
    d.tt = 0.2000000000
    d.tb = 0.3000000000
    d.TBase = 0.3000000000
    d.Base = 2.0000000000
    d.LToe = 0.8000000000
    d.LHeel = d.Base - d.LToe - d.tb
    d.ASst_DB = 101
    d.ASst_Sp = 111
    d.AStoe_DB = 100
    d.AStoe_Sp = 111
    d.ASheel_DB = 100
    d.ASheel_Sp = 111
    d.IsValid = True
    Call ShowFixture(d, mat)
    Print #f, "CLIPPED_ROWS=" & clipped & "; FAILURES=" & failures
    Close #f
    Form1.txtH.Text = "4": Form1.txtH1.Text = "1.2"
    Form1.Show
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Description
    Close #f
End Sub

Private Sub ShowFixture(d As Design, mat As MaterialProperties)
    Dim s As String, line As Variant, i As Integer, width As Single, available As Single, output As Integer
    s = FormatResults(d, mat, "BA")
    Print #f, "H=" & H
    Print #f, s
    output = FreeFile
    Open App.Path & "\full-report-H" & H & ".txt" For Output As #output
    Print #output, FormatResults(d, mat, "BA")
    Close #output
    Form1.lstResults.Clear
    For Each line In Split(s, vbCrLf)
        Form1.lstResults.AddItem CStr(line)
    Next line
    Set Form1.picGraph.Font = Form1.lstResults.Font
    Form1.picGraph.ScaleMode = vbPixels
    available = Form1.ScaleX(Form1.lstResults.Width, vbTwips, vbPixels) - 24
    For i = 0 To Form1.lstResults.ListCount - 1
        width = Form1.picGraph.TextWidth(Form1.lstResults.List(i))
        If width > available Then clipped = clipped + 1
    Next i
    Print #f, "ROWS=" & Form1.lstResults.ListCount & "; AVAILABLE_PX=" & available
    Form1.lstResults.TopIndex = 0
End Sub
