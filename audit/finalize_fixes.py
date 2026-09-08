from pathlib import Path
import re
P=Path(__file__).resolve().parent.parent
def read(n): return (P/n).read_bytes().decode('latin1').replace('\r\n','\n')
def write(n,s): (P/n).write_bytes(s.replace('\n','\r\n').encode('latin1'))
s=read('modShared.bas')
s=s.replace('Public RunSeed As Long','Public RunSeed As Long\nPublic RunTrial As Long\nPublic RunAlgorithm As String')
s=s.replace('Public Sub BeginSearch(budget As Long, seed As Long, algorithm As String)','Public Sub BeginSearch(budget As Long, seed As Long, algorithm As String, Optional trial As Long = 1)')
s=s.replace('    EvaluationCount = 0: EvaluationBudget = budget: RunSeed = seed','    EvaluationCount = 0: EvaluationBudget = budget: RunSeed = seed\n    RunTrial = trial: RunAlgorithm = algorithm')
s=s.replace('    Print #f, "Status=" & RunStatus','    Print #f, "Status=" & RunStatus\n    Print #f, "Algorithm=" & RunAlgorithm & "; Trial=" & RunTrial')
s=s.replace('    Print #f, FormatResults(RunBest, currentMaterial)','    Print #f, FormatResults(RunBest, currentMaterial, RunAlgorithm, RunTrial)')
s=s.replace('Optional AlgoName As String = "Hill Climbing v5.0") As String','Optional AlgoName As String = "Hill Climbing", Optional ReportTrial As Long = 0) As String')
s=s.replace('    Dim reportOK As Boolean, reportQT As Double, reportQH As Double','    Dim reportOK As Boolean, reportQT As Double, reportQH As Double\n    On Error GoTo InvalidReport\n    currentMaterial = mat\n    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)')
s=s.replace('    result = result & "Best Found at Trial: " & modDataStructures.BestTrial & vbCrLf','    If ReportTrial = 0 Then ReportTrial = modDataStructures.BestTrial\n    result = result & "Best Found at Trial: " & ReportTrial & vbCrLf')
s=s.replace('    FormatResults = result\nEnd Function','    FormatResults = result\n    Exit Function\nInvalidReport:\n    FormatResults = "INVALID_DATA: " & Err.Description\nEnd Function')
s=s.replace("' Check Bearing Capacity (FS_BC >= 2.0)","' Check ALLOWABLE bearing pressure (qa/qmax >= 1.0)")
s=s.replace('Public Const FS_BC_MIN As Double = 1#      \' Bearing Capacity','Public Const FS_BC_MIN As Double = 1#      \' qa is allowable, not ultimate')
s=s.replace('result = result & "WSD: n="','result = result & "Legacy WSD parameters (unverified unless sourced): n="')
write('modShared.bas',s)
for name,algo in [('modBA.bas','BA'),('modHillClimbing.bas','HCA')]:
    s=read(name)
    s=s.replace('Optional RandomSeed As Long = 12345) As Design','Optional RandomSeed As Long = 12345, _\n                       Optional TrialNumber As Long = 1) As Design')
    s=s.replace(f'BeginSearch(MaxIterations, RandomSeed, "{algo}")',f'BeginSearch(MaxIterations, RandomSeed, "{algo}", TrialNumber)')
    # Keep legacy CSV entrypoints useful and aligned to the authoritative trace.
    start=s.index('Public Sub SaveAcceptCSV')
    end=s.index('\nEnd Sub',start)+len('\nEnd Sub')
    signature=s[start:s.index('\n',start)]
    s=s[:start]+signature+'\n    \' Per-trial evaluations.csv is saved by FinishSearch.\nEnd Sub'+s[end:]
    write(name,s)
s=read('modWSD.bas')
s=re.sub(r'Public Function CalculateRhoMax\([\s\S]*?\nEnd Function','''Public Function CalculateRhoMax(wsd As WSDParams) As Double
    Err.Raise 5, , "No sourced WSD maximum-steel rule; 0.75*rho_balanced is not assumed"
End Function''',s,count=1)
s=s.replace("' For WSD: ","' Elastic balanced section (not a code maximum): ")
s=s.replace('    \' Modular ratio (constant)','    If fc_prime <= 0 Or (fy <> 3000 And fy <> 4000) Then Err.Raise 5, , "Invalid/unsupported material"\n    \' Modular ratio (legacy assumed constant, pending EIT verification)')
write('modWSD.bas',s)
s=read('Form1.frm')
s=s.replace('RandomSeed:=CLng(txtSeed.Text) + CLng(trial) - 1)', 'RandomSeed:=CLng(txtSeed.Text) + CLng(trial) - 1, TrialNumber:=CLng(trial))')
# Replace redundant modal completion dialog with persistent result-list location.
s=re.sub(r'    MsgBox [^\n]+ & vbCrLf & vbCrLf & _\n[\s\S]*?vbInformation, [^\n]+',
    '    lstResults.AddItem "Per-trial reports: " & App.Path & "\\\\results"\n    lstResults.AddItem "Seed start: " & txtSeed.Text & "; budget includes initial/reset/neighbor."',s)
s=s.replace('    cmdBA.Enabled = False','    cmdBA.Enabled = False: cmdRun.Enabled = False: cmdCommand1.Enabled = False: cmdCompare.Enabled = False')
s=s.replace('    cmdRun.Enabled = False\n','    cmdBA.Enabled = False: cmdRun.Enabled = False: cmdCommand1.Enabled = False: cmdCompare.Enabled = False\n')
s=s.replace('    cmdBA.Enabled = True','    cmdBA.Enabled = True: cmdRun.Enabled = True: cmdCommand1.Enabled = True: cmdCompare.Enabled = True')
s=s.replace('    cmdRun.Enabled = True\n','    cmdBA.Enabled = True: cmdRun.Enabled = True: cmdCommand1.Enabled = True: cmdCompare.Enabled = True\n')
write('Form1.frm',s)
s=read('modBatch.bas').replace('RandomSeed:=12345 + CLng(trial) - 1)', 'RandomSeed:=12345 + CLng(trial) - 1, TrialNumber:=CLng(trial))')
write('modBatch.bas',s)
