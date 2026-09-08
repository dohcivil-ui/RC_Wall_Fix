"""Install the bounded project WSD/ACI check path, preserving legacy VB6 bytes."""
from pathlib import Path
P = Path(__file__).resolve().parent.parent
def read(name): return (P/name).read_bytes().decode('latin1').replace('\r\n','\n')
def write(name,text): (P/name).write_bytes(text.replace('\n','\r\n').encode('latin1'))
s=read('modWSD.bas')
s=s.replace("' No guessed EIT clauses: these remain unset until reviewed against 011007-19.","' Legacy test thresholds; production uses the explicit PROJECT_WSD_ACI99_V1 basis.")
s=s.replace('Public Type WSDParams\n','Public Type WSDParams\n    fy As Double\n    fcPrime As Double\n')
s=s.replace('    params.n = 2040000#', '    params.fy = fy: params.fcPrime = fc_prime\n    params.n = 2040000#',1)
s=s.replace('Public Function WSDCriteriaReady() As Boolean\n','Public Function WSDCriteriaReady() As Boolean\n    If ProjectChecksEnabled Then\n        WSDCriteriaReady = WSDReviewed And WSDSource = PROJECT_CHECK_BASIS\n        Exit Function\n    End If\n')
write('modWSD.bas',s)
s=read('modShared.bas')
s=s.replace('Public Function CalculateCost(d As Design) As Double\n','Public Function CalculateCost(d As Design) As Double\n    Dim detail As ProjectDetail\n    If ProjectChecksEnabled Then\n        If CheckProjectDesign(d, detail) Then CalculateCost = detail.Cost Else CalculateCost = NO_SOLUTION_COST\n        Exit Function\n    End If\n')
anchor='    LastValidationReason = "INVALID_GEOMETRY_OR_INPUT"\n'
replacement='''    If ProjectChecksEnabled Then
        Dim detail As ProjectDetail
        d.ASst_DB = StemDB: d.ASst_Sp = StemSP
        d.AStoe_DB = ToeDB: d.AStoe_Sp = ToeSP
        d.ASheel_DB = HeelDB: d.ASheel_Sp = HeelSP
        CheckDesignValid = CheckProjectDesign(d, detail)
        FS_OT = detail.OT: FS_SL = detail.SL: FS_BC = detail.BC
        d.FS_OT = FS_OT: d.FS_SL = FS_SL: d.FS_BC = FS_BC
        d.IsValid = CheckDesignValid
        Exit Function
    End If
'''
s=s.replace(anchor,replacement+anchor,1)
anchor='    Dim result As String\n    Dim wsd As WSDParams\n'
s=s.replace(anchor,'    If ProjectChecksEnabled Then\n        FormatResults = ProjectDesignReport(d, mat, AlgoName)\n        Exit Function\n    End If\n'+anchor,1)
s=s.replace('    Print #f, "AllowableShear=" & AllowableShear', '    Print #f, "ProjectChecksEnabled=" & ProjectChecksEnabled\n    Print #f, "AllowableShear=" & AllowableShear',1)
write('modShared.bas',s)
s=read('Form1.frm')
i=s.index('Private Sub Form_Load()')
s=s[:i]+s[i:].replace('    Call InitializeArrays','    Call InitializeArrays\n    Call EnableProjectChecks',1)
s=s.replace('        lstResults.AddItem "Trial Cost: " & Format(trialCost, "#,##0.00") & " Baht/m"','''        If trialDesign.IsValid Then
            lstResults.AddItem "Trial Cost: " & Format(trialCost, "#,##0.00") & " Baht/m"
        Else
            lstResults.AddItem "Trial: NO_SOLUTION (no admissible price)"
        End If''')
s=s.replace('        lstResults.AddItem "Best So Far: " & Format(globalBestCost, "#,##0.00") & " (Trial " & globalBestTrial & ")"','''        If globalBestTrial > 0 Then
            lstResults.AddItem "Best So Far: " & Format(globalBestCost, "#,##0.00") & " (Trial " & globalBestTrial & ")"
        Else
            lstResults.AddItem "Best So Far: NO_SOLUTION"
        End If''')
write('Form1.frm',s)
for path in [P/'RC_RT_HCA_v2.vbp',*(P/'audit').glob('*.vbp')]:
 s=path.read_bytes().decode('latin1').replace('\r\n','\n')
 relative='' if path.parent==P else '..\\'
 if 'Module=modProjectChecks;' not in s:
  s+='\nModule=modProjectChecks; '+relative+'modProjectChecks.bas\n'
 path.write_bytes(s.replace('\n','\r\n').encode('latin1'))
# Keep historical synthetic regression fixtures explicitly in their old mode.
s=read('audit/GuiRegression.bas').replace('    Load Form1\n','    Load Form1\n    ProjectChecksEnabled = False\n')
write('audit/GuiRegression.bas',s)
