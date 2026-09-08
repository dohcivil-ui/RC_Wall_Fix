"""Build a VB6 harness that invokes the real Form1 BA button for 30 trials."""
from pathlib import Path
import re
P = Path(__file__).resolve().parent
project = (P.parent/'RC_RT_HCA_v2.vbp').read_bytes().decode('latin1')
project = project.replace('Startup="Form1"','Startup="Sub Main"').replace('Name="Project1"','Name="Trial30Project"')
project = re.sub(r'(?m)^(Module=[^;]+; )([^\r\n]+)',r'\1..\\\2',project)
project = project.replace('Form=Form1.frm','Form=..\\Form1.frm').replace('Form=frmBestDesign.frm','Form=..\\frmBestDesign.frm')
project += '\r\nModule=Trial30; Trial30.bas\r\nExeName32="Trial30.exe"\r\n'
(P/'Trial30.vbp').write_bytes(project.encode('latin1'))
code = '''Attribute VB_Name = "Trial30"
Option Explicit

Public Sub Main()
    Dim output As Integer, i As Integer, started As Double, elapsed As Double
    On Error GoTo Failed
    output = FreeFile
    Open App.Path & "\\trial30-native.txt" For Output As #output
    Load Form1
    Form1.Show
    Form1.txtH.Text = "5": Form1.txtH1.Text = "1.2"
    Form1.txtMu.Text = "0.6": Form1.txtGammaSoil.Text = "1.8"
    Form1.txtGammaCon.Text = "2.4": Form1.txtPhi.Text = "30"
    Form1.txtQa.Text = "30": Form1.txtCover.Text = "7.5"
    Form1.txtMaxIter.Text = "5000": Form1.txtTrials.Text = "30"
    Form1.txtSeed.Text = "12345"
    For i = 0 To Form1.cboConcreteStrength.ListCount - 1
        If Form1.cboConcreteStrength.List(i) = "320" Then Form1.cboConcreteStrength.ListIndex = i
    Next i
    If Form1.cboConcreteStrength.Text <> "320" Then Err.Raise 5, , "Concrete selection failed"
    If PassiveFactor <> 1# Then Err.Raise 5, , "Project passive setting changed"
    Print #output, "ACTUAL VB6 Form1 BA button: 30 trials, 5000 evaluations each, seeds 12345..12374."
    Print #output, "H=5; H1=1.2; fc_prime=320; fy=4000; qa_allowable=30; clear_cover=0.075; full passive=1."
    Print #output, "Production criteria unchanged: WSDReviewed=" & WSDReviewed & "; source=" & WSDSource
    Print #output, "AllowableShear=" & AllowableShear & "; MinStemRatio=" & MinStemRatio & "; MinBaseRatio=" & MinBaseRatio
    started = Timer
    Form1.cmdBA.Value = True
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    If RunTrial <> 30 Or EvaluationCount <> 5000 Or PassiveFactor <> 1# Then Err.Raise 5, , "Incomplete or inconsistent trial run"
    If Not Form1.cmdBA.Enabled Or Not Form1.cmdRun.Enabled Then Err.Raise 5, , "Form controls not restored"
    If Abs(currentWSD.n - 2040000# / (15100# * Sqr(320#))) > 0.000000001 Then Err.Raise 5, , "Incorrect modular ratio"
    Print #output, "NativeElapsedSeconds=" & CsvNumber(elapsed)
    Print #output, "FinalTrial=" & RunTrial & "; final seed=" & RunSeed & "; n=" & CsvNumber(currentWSD.n)
    For i = 0 To Form1.lstResults.ListCount - 1
        Print #output, Form1.lstResults.List(i)
    Next i
    Print #output, "NATIVE_TRIAL30_COMPLETE"
    Unload Form1
    Close #output
    Exit Sub
Failed:
    Print #output, "FATAL " & Err.Number & ": " & Err.Description
    Close #output
    Unload Form1
End Sub
'''
(P/'Trial30.bas').write_bytes(code.replace('\n','\r\n').encode('ascii'))
print('Trial30 harness generated from real .vbp; no production code or WSD settings changed.')
