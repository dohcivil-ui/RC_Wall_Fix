from pathlib import Path

p = Path(__file__).resolve().parent
old = (p.parent / 'results-layout/after/ResultsLayout.bas').read_text(encoding='latin1')
src = old[:old.index('    d.IsValid = False')]
src = src.replace('Private f As Integer, failures As Long, clipped As Long', '''Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal index As Long) As Long
Private f As Integer, failures As Long''')
src = src.replace('"\\layout.txt"', '"\\runtime.txt"')
src += '''    If Not Form1.txtResults.MultiLine Then failures = failures + 1
    If Form1.txtResults.ScrollBars <> 2 Then failures = failures + 1
    If Not Form1.txtResults.Locked Or Not Form1.txtResults.Enabled Then failures = failures + 1
    ' Native edit style: ES_AUTOHSCROLL must be off for word wrapping.
    If (GetWindowLong(Form1.txtResults.hwnd, -16) And &H80&) <> 0 Then failures = failures + 1
    Form1.txtResults.Text = vbNullString
    s = String$(150, "x") & " C:\\folder with spaces\\trial-summary.csv"
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
    Open App.Path & "\\full-report-H" & H & ".txt" For Output As #output
    Print #output, FormatResults(d, mat, "BA")
    Close #output
End Sub
'''
(p / 'ResultsLayout.bas').write_bytes(src.replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii'))
vbp = (p/'MainCompile.vbp').read_bytes().replace(b'Startup="Form1"', b'Startup="Sub Main"')
vbp += b'\r\nModule=ResultsLayout; ResultsLayout.bas\r\nExeName32="TextBoxProbe.exe"\r\n'
(p/'TextBoxProbe.vbp').write_bytes(vbp)
