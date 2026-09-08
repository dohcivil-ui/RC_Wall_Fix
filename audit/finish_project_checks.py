"""One-time finishing patch; do not rerun after it has been applied."""
from pathlib import Path
root = Path(__file__).resolve().parent.parent
p = root / 'audit/modProjectChecks.source'
s = p.read_text(encoding='ascii')
s = s.replace('Public ProjectChecksEnabled As Boolean', 'Public ProjectChecksEnabled As Boolean\nPublic ProjectTrialSummary As String')
s = s.replace('If frontLd <= d.TBase - cover Then', 'If frontLd <= d.TBase - cover And length - cover + frontLd <= 12# Then')
s = s.replace('    r.DB(0) = d.ASst_DB: r.SP(0) = d.ASst_Sp', '    If hs <= 2# * cover Or d.Base <= 2# * cover Then Exit Function\n    r.DB(0) = d.ASst_DB: r.SP(0) = d.ASst_Sp')
s += '''
Public Sub RecordProjectTrial()
    Dim f As Integer, price As String
    If Not ProjectChecksEnabled Then Exit Sub
    f = FreeFile
    If Len(ProjectTrialSummary) = 0 Then
        ProjectTrialSummary = RunFolder & "\\trial-summary.csv"
        Open ProjectTrialSummary For Output As #f
        Print #f, "trial,seed,status,evaluations,best_cost,best_evaluation,run_folder"
        Close #f
    End If
    If RunBest.IsValid Then price = CsvNumber(RunBestCost)
    Open ProjectTrialSummary For Append As #f
    Print #f, RunTrial & "," & RunSeed & "," & RunStatus & "," & EvaluationCount & "," & price & "," & RunBestEvaluation & "," & Chr$(34) & Replace$(RunFolder, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
    Close #f
End Sub
'''
p.write_text(s, encoding='ascii')
(root/'modProjectChecks.bas').write_bytes(s.replace('\n','\r\n').encode('ascii'))
p = root/'Form1.frm'
s = p.read_bytes().decode('latin1')
s = s.replace('    globalBestIteration = 0\r\n', '    globalBestIteration = 0\r\n    ProjectTrialSummary = ""\r\n')
for line in ['        trialCost = modShared.CalculateCost(trialDesign)', '        currentCost = CalculateCost(bestDesign)']:
    assert s.count(line) == 1
    s = s.replace(line, line + '\r\n        Call RecordProjectTrial')
line = '    lstResults.AddItem "Per-trial reports: " & App.Path & "\\results"'
assert s.count(line) == 2
s = s.replace(line, line + '\r\n    If ProjectChecksEnabled Then lstResults.AddItem "Trial summary: " & ProjectTrialSummary')
p.write_bytes(s.encode('latin1'))
p = root/'audit/build_project_regression.py'
s = p.read_text(encoding='utf8')
s = s.replace('    Dim sr As ProjectDetail', '    Dim summaryFile As Integer, summaryHeader As String, summaryRow As String, noSolutionSummary As String\n    Dim sr As ProjectDetail')
s = s.replace('    Form1.cboConcreteStrength.ListIndex = Form1.cboConcreteStrength.ListCount - 2', '''    Form1.txtH.Text = "5": Form1.txtH1.Text = "1.2"
    For i = 0 To Form1.cboConcreteStrength.ListCount - 1
        If Form1.cboConcreteStrength.List(i) = "320" Then Form1.cboConcreteStrength.ListIndex = i
    Next i''')
s = s.replace('    Unload Form1', '''    noSolutionSummary = ProjectTrialSummary
    summaryFile = FreeFile: Open ProjectTrialSummary For Input As #summaryFile
    Line Input #summaryFile, summaryHeader: Line Input #summaryFile, summaryRow
    Call Verify("no solution summary has blank price", InStr(summaryRow, ",NO_SOLUTION,4,,0,") > 0)
    Call Verify("summary contains one trial", EOF(summaryFile))
    Close #summaryFile
    Form1.txtMaxIter.Text = "64": Form1.txtSeed.Text = "20260908"
    Form1.cmdBA.Value = True
    text = ""
    For i = 0 To Form1.lstResults.ListCount - 1: text = text & Form1.lstResults.List(i) & vbCrLf: Next i
    Call Verify("actual UI accepted detail", InStr(text, "PASS_IMPLEMENTED_PROJECT_CHECKS") > 0 And InStr(text, "Stem horizontal EACH face") > 0)
    Call Verify("new UI session preserves old summary", ProjectTrialSummary <> noSolutionSummary And Len(Dir$(noSolutionSummary)) > 0)
    summaryFile = FreeFile: Open ProjectTrialSummary For Input As #summaryFile
    Line Input #summaryFile, summaryHeader: Line Input #summaryFile, summaryRow
    Call Verify("accepted summary has actual price", InStr(summaryRow, ",SOLUTION_FOUND_CONFIGURED_CHECKS,64," & CsvNumber(RunBestCost) & ",") > 0)
    Close #summaryFile
    Print #output, "GUI64: " & text
    Unload Form1''')
p.write_text(s, encoding='utf8')
