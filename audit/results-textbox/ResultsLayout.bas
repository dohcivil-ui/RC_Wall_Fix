Attribute VB_Name = "ResultsLayout"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal index As Long) As Long
Private f As Integer, failures As Long
Public Sub Main()
    Dim d As Design, mat As MaterialProperties, s As String, joined As String, start As Integer, i As Integer
    On Error GoTo Failed
    Load Form1
    Form1.Caption = "Results layout after - native VB6 test"
    Call InitializeArrays
    Call EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 30: cover = 0.075
    f = FreeFile
    Open App.Path & "\runtime.txt" For Output As #f
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
    modDataStructures.BestTrial = 1
    modDataStructures.BestCostIteration = 279
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
    modDataStructures.BestTrial = 1
    modDataStructures.BestCostIteration = 279
    Call ShowFixture(d, mat)
    If Not Form1.txtResults.MultiLine Then failures = failures + 1
    If Form1.txtResults.ScrollBars <> 2 Then failures = failures + 1
    If Not Form1.txtResults.Locked Or Not Form1.txtResults.Enabled Then failures = failures + 1
    ' Native edit style: ES_AUTOHSCROLL must be off for word wrapping.
    If (GetWindowLong(Form1.txtResults.hwnd, -16) And &H80&) <> 0 Then failures = failures + 1
    Form1.txtResults.Text = vbNullString
    s = String$(150, "x") & " C:\folder with spaces\trial-summary.csv"
    Form1.AddResultLine s
    Form1.AddResultLine ""
    Form1.AddResultLine "Last row"
    If Form1.txtResults.Text <> s & vbCrLf & vbCrLf & "Last row" & vbCrLf Then failures = failures + 1
    i = SendMessage(Form1.txtResults.hwnd, &HBA, 0, 0)
    Print #f, "WRAPPED_LINE_COUNT=" & i
    If i <= 4 Then failures = failures + 1
    Form1.txtResults.Width = Form1.txtResults.Width / 2
    Print #f, "NARROW_LINE_COUNT=" & SendMessage(Form1.txtResults.hwnd, &HBA, 0, 0)
    If SendMessage(Form1.txtResults.hwnd, &HBA, 0, 0) <= i Then failures = failures + 1
    Form1.txtResults.SelStart = 0
    Form1.txtResults.SelLength = Len(Form1.txtResults.Text)
    If Form1.txtResults.SelText <> Form1.txtResults.Text Then failures = failures + 1
    Print #f, "FAILURES=" & failures
    Close #f
    Unload Form1
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Description
    Close #f
    Unload Form1
End Sub

Private Sub ShowFixture(d As Design, mat As MaterialProperties)
    Dim s As String, line As Variant, output As Integer
    s = FormatScreenResults(d, mat, "BA")
    Form1.txtResults.Text = vbNullString
    For Each line In Split(s, vbCrLf)
        Form1.AddResultLine CStr(line)
    Next line
    If Form1.txtResults.Text <> s & vbCrLf Then failures = failures + 1
    Print #f, "H=" & H & "; REPORT_CHARACTERS=" & Len(s) & "; TEXT_PRESERVED=" & CStr(Form1.txtResults.Text = s & vbCrLf)
    output = FreeFile
    Open App.Path & "\full-report-H" & H & ".txt" For Output As #output
    Print #output, FormatResults(d, mat, "BA")
    Close #output
End Sub
