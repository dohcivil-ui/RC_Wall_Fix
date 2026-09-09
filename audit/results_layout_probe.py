"""Build native before/after result-panel fixtures; no optimization or research CSV writes."""
from pathlib import Path
import json,csv,re
P=Path(__file__).resolve().parent.parent
records=json.loads((P/'audit/h5-search/release/verification.json').read_text())['records']
fixtures=[]
for height in (5,4):
    record=next(r for r in records if r['height']==height and r['method']=='HCA' and r['seed']==(12374 if height==5 else 12345))
    rows=list(csv.DictReader((Path(record['path'])/'evaluations.csv').open()))
    row=rows[record['evaluation']-1]
    fixtures.append((height,row))
for phase in ('before','after'):
    dest=P/'audit/results-layout'/phase;dest.mkdir(exist_ok=True,parents=True)
    if phase=='after':
        for name in ('Form1.frm','Form1.frx','modShared.bas'):dest.joinpath(name).write_bytes(P.joinpath(name).read_bytes())
    def write(name,s):dest.joinpath(name).write_bytes(s.replace('\r\n','\n').replace('\n','\r\n').encode('latin1'))
    v=P.joinpath('RC_RT_HCA_v2.vbp').read_bytes().decode('latin1')
    v=re.sub(r'^(Module=[^;]+; )([^\r\n]+)',lambda m:m[1]+str((dest if m[2]=='modShared.bas' else P)/m[2]),v,flags=re.M)
    v=re.sub(r'^Form=([^\r\n]+)',lambda m:'Form='+str((dest if m[1]=='Form1.frm' else P)/m[1]),v,flags=re.M)
    v=v.replace('Startup="Form1"','Startup="Sub Main"')
    write('ResultsLayout.vbp',v+'\nModule=ResultsLayout; ResultsLayout.bas\nExeName32="ResultsLayout.exe"\n')
    cases=''
    for height,row in fixtures:
        cases+=f'    H = {height}: mat = GetSD40Material({120+40*height}, GetConcretePrice({120+40*height}), 24)\n'
        for key in ('tt','tb','TBase','Base','LToe'):cases+=f'    d.{key} = {row[key]}\n'
        cases+='    d.LHeel = d.Base - d.LToe - d.tb\n'
        for key,field in [('StemDB','ASst_DB'),('StemSP','ASst_Sp'),('ToeDB','AStoe_DB'),('ToeSP','AStoe_Sp'),('HeelDB','ASheel_DB'),('HeelSP','ASheel_Sp')]:cases+=f'    d.{field} = {row[key]}\n'
        cases+='    d.IsValid = True\n    Call ShowFixture(d, mat)\n'
    format_call='FormatResults(d, mat, "BA")' if phase=='before' else 'FormatScreenResults(d, mat, "BA")'
    add='Form1.lstResults.AddItem CStr(line)' if phase=='before' else 'Form1.AddResultLine CStr(line)'
    after_checks='' if phase=='before' else '''    d.IsValid = False
    s = FormatScreenResults(d, mat, "HCA")
    If InStr(s, "NO_SOLUTION") = 0 Or InStr(s, "Cost") > 0 Then failures = failures + 1
    d.IsValid = True: d.tb = 0.2: d.LHeel = d.Base - d.LToe - d.tb
    d.ASst_DB = 100: d.ASst_Sp = 113
    s = FormatScreenResults(d, mat, "HCA")
    If InStr(s, "Check failed:") = 0 Or InStr(s, "Status: PASS") > 0 Then failures = failures + 1
    ' Verify wrapping is lossless, including a long path without spaces.
    s = "C:\\" & String$(120, "x") & "\\trial-summary.csv"
    start = Form1.lstResults.ListCount
    Form1.AddResultLine s
    For i = start To Form1.lstResults.ListCount - 1
        joined = joined & Form1.lstResults.List(i)
    Next i
    If joined <> s Then failures = failures + 1
    For i = Form1.lstResults.ListCount - 1 To start Step -1
        Form1.lstResults.RemoveItem i
    Next i
'''
    write('ResultsLayout.bas',f'''Attribute VB_Name = "ResultsLayout"
Option Explicit
Private f As Integer, failures As Long, clipped As Long
Public Sub Main()
    Dim d As Design, mat As MaterialProperties, s As String, joined As String, start As Integer, i As Integer
    On Error GoTo Failed
    Load Form1
    Form1.Caption = "Results layout {phase} - native VB6 test"
    Call InitializeArrays
    Call EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 30: cover = 0.075
    f = FreeFile
    Open App.Path & "\\layout.txt" For Output As #f
{cases}{after_checks}    Print #f, "CLIPPED_ROWS=" & clipped & "; FAILURES=" & failures
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
    s = {format_call}
    Print #f, "H=" & H
    Print #f, s
    output = FreeFile
    Open App.Path & "\\full-report-H" & H & ".txt" For Output As #output
    Print #output, FormatResults(d, mat, "BA")
    Close #output
    Form1.lstResults.Clear
    For Each line In Split(s, vbCrLf)
        {add}
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
''')
    print(dest)
