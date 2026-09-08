from pathlib import Path
import re
root=Path(__file__).resolve().parent.parent
def read(n): return (root/n).read_bytes().decode('latin1').replace('\r\n','\n')
def write(n,s): (root/n).write_bytes(s.replace('\n','\r\n').encode('latin1'))
s=read('modWSD.bas')
a=s.index('Public Function WSDCriteriaReady')
b=s.index("' ========================================",a)
helpers=s[a:b]
s=s[:a]+s[b:]+ '\n'+helpers
write('modWSD.bas',s)
s=read('Form1.frm')
s=s.replace('   Begin VB.CommandButton cmdCommand1', '''   Begin VB.TextBox txtSeed
      Height          =   375
      Left            =   19800
      Top             =   8880
      Width           =   1455
      TabIndex        =   45
      Text            =   "12345"
      ToolTipText     =   "Integer seed; trial n uses seed+n-1 (same for BA/HCA)"
   End
   Begin VB.Label lblSeed
      Caption         =   "Seed"
      Height          =   375
      Left            =   19800
      Top             =   8400
      Width           =   1455
      TabIndex        =   46
   End
   Begin VB.CommandButton cmdCommand1''',1)
s=s.replace('BA v9.0 (Bisection Base + HCA Random)','BA (tb, TBase, Base only)')
s=s.replace('BA v9.0 - Trial ', 'BA triple - Trial ')
s=s.replace('    Randomize Timer','    Call SeedSearchRandom(CLng(txtSeed.Text))')
s=s.replace('maxIter, h, H1, gamma_soil, gamma_con, phi, mu, qa, cover, selectedMaterial)', 'maxIter, h, H1, gamma_soil, gamma_con, phi, mu, qa, cover, selectedMaterial, _\n            RandomSeed:=CLng(txtSeed.Text) + CLng(trial) - 1)')
# Global UI best must exclude no-solution returns regardless of sentinel cost.
s=s.replace('        If (trialCost < globalBestCost) Or _','        If trialDesign.IsValid And ((trialCost < globalBestCost) Or _')
s=s.replace('        If (currentCost < globalBestCost) Or _','        If bestDesign.IsValid And ((currentCost < globalBestCost) Or _')
s=s.replace('(globalBestIteration = 0 Or modDataStructures.BestCostIteration < globalBestIteration)) Then', '(globalBestIteration = 0 Or modDataStructures.BestCostIteration < globalBestIteration))) Then')
# Allocate no-solution graph too, avoid uninitialized array on the first failed trial.
s=s.replace('    globalBestCost = 999999999', '    ReDim globalBestCostHistory(1 To maxIter)\n    globalBestCost = NO_SOLUTION_COST')
s=s.replace('    Call DrawCostGraph(picGraph, globalBestCostHistory, globalBestIteration)', '    If globalBestIteration > 0 Then Call DrawCostGraph(picGraph, globalBestCostHistory, globalBestIteration)')
s=s.replace('    BA_HasData = True', '    BA_HasData = (globalBestIteration > 0)').replace('    HCA_HasData = True', '    HCA_HasData = (globalBestIteration > 0)')
s=s.replace('baHasRun = True','baHasRun = (globalBestIteration > 0)').replace('hcaHasRun = True','hcaHasRun = (globalBestIteration > 0)')
# Per-trial authoritative export is FinishSearch, not the obsolete four-column buffer.
s=re.sub(r'^\s*Call (?:modBA\.)?SaveAcceptCSV(?:_BA)?\(h\)\n','\n',s,flags=re.M)
s=re.sub(r'^\s*Call (?:modBA\.)?SaveLoopPriceCSV(?:_BA)?\(h\)\n','\n',s,flags=re.M)
s=re.sub(r'"- D:\\accept-(?:BA|HCA)-H" & Format\(h, "0"\) & "\.csv"', '"Per-trial reports: " & App.Path & "\\\\results"',s)
s=re.sub(r'"- D:\\loopPrice-(?:BA|HCA)-H" & Format\(h, "0"\) & "\.csv"', '"Includes evaluations.csv and run.txt (seed, status, inputs)"',s)
s=s.replace('Private Function ValidateInputs() As Boolean','Private Function ValidateInputs() As Boolean\n    On Error GoTo InvalidInput')
s=s.replace('    ValidateInputs = True\nEnd Function','''    If CDbl(txtH1.Text) > CDbl(txtH.Text) Then Exit Function
    If CDbl(txtH1.Text) < 0.3 Then Exit Function
    If CDbl(txtCover.Text) / 100# >= 0.2 Then Exit Function
    If CDbl(txtMaxIter.Text) <> Fix(CDbl(txtMaxIter.Text)) Then Exit Function
    If CDbl(txtTrials.Text) <> Fix(CDbl(txtTrials.Text)) Then Exit Function
    If Not IsNumeric(txtSeed.Text) Then Exit Function
    If CDbl(txtSeed.Text) <> Fix(CDbl(txtSeed.Text)) Then Exit Function
    If CDbl(txtSeed.Text) < 0 Or CDbl(txtSeed.Text) + CDbl(txtTrials.Text) > 2147483647# Then Exit Function
    ValidateInputs = True
InvalidInput:
End Function''')
# Tooltip makes the height convention explicit without changing the drawing layout.
s=s.replace('Private Sub Form_Load()', '''Private Sub Form_Load()
    txtH.ToolTipText = "Retained soil elevation above BASE UNDERSIDE; stem=H-TBase"
    txtH1.ToolTipText = "FRONT soil elevation above BASE UNDERSIDE (not retained height)"
    txtCover.ToolTipText = "CLEAR cover to bar surface (cm); d=t-cover-db/2"
    txtQa.ToolTipText = "ALLOWABLE bearing pressure, ton-force/m2"
    txtMaxIter.ToolTipText = "Total candidate evaluation budget including initial and reset"
''')
write('Form1.frm',s)
s=read('modShared.bas')+'''
Public Sub SeedSearchRandom(seed As Long)
    Dim dummy As Single
    dummy = Rnd(-1): Randomize seed
End Sub
'''
s=s.replace('    currentMaterial = mat\n    currentWSD', '''    Dim reportOK As Boolean, reportQT As Double, reportQH As Double
    reportOK = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, FS_OT, FS_SL, FS_BC)
    If Not BearingEdges(d, e, reportQT, reportQH) Then
        FormatResults = "Design check: STABILITY_OR_BASE_CONTACT; no supported full-contact pressure diagram."
        Exit Function
    End If
    currentMaterial = mat
    currentWSD''')
s=s.replace('    result = "Design check: " & LastValidationReason', '''    result = result & SectionCheckSummary("Stem", M_stem, d_stem, As_prov_stem, SectionShear(d, 0, d_stem), d.tb, True)
    result = result & SectionCheckSummary("Toe", M_toe, d_toe, As_prov_toe, SectionShear(d, 1, d_toe), d.TBase, False)
    result = result & SectionCheckSummary("Heel", M_heel, d_heel, As_prov_heel, SectionShear(d, 2, d_heel), d.TBase, False)
    result = "Design check: " & LastValidationReason''')
s+='''
Private Function SectionCheckSummary(label As String, M As Double, depth As Double, steel As Double, shear As Double, thickness As Double, stem As Boolean) As String
    Dim c As Double, s As Double, j As Double, minimum As Double
    Call SectionStresses(M, depth, steel, currentWSD.n, c, s, j)
    If stem Then minimum = MinStemRatio * thickness * 10000# Else minimum = MinBaseRatio * thickness * 10000#
    SectionCheckSummary = label & ": d=" & Format(depth, "0.0000") & " m; fc_actual=" & Format(c, "0.00") & _
        "/" & currentWSD.fc & "; fs_actual=" & Format(s, "0.00") & "/" & currentWSD.fs & " kgf/cm2" & vbCrLf & _
        "v=" & Format(shear / (10# * depth), "0.000") & "/" & AllowableShear & " kgf/cm2; As_min=" & minimum & " cm2/m" & vbCrLf
    If Not WSDCriteriaReady() Then SectionCheckSummary = SectionCheckSummary & "Allowables/minimums UNVERIFIED (zero means unset, never a passing criterion)." & vbCrLf
End Function
'''
write('modShared.bas',s)
s=read('modBatch.bas')
for method in ['RunBatchStep3','RunBatchStep3_A3']:
    s=s.replace(f'Public Sub {method}()', f'''Public Sub {method}()
    ' Research batch is blocked until authoritative WSD review is complete.
    If Not WSDCriteriaReady() Then
        MsgBox "Batch not run: WSD criteria remain unverified.", vbExclamation
        Exit Sub
    End If''')
s=s.replace('    Randomize Timer','    Call SeedSearchRandom(12345)')
s=s.replace('allowQa, conCover, mat)', 'allowQa, conCover, mat, RandomSeed:=12345 + CLng(trial) - 1)')
s=s.replace('    IsValid = (iterToBest > 0)', '    IsValid = d.IsValid')
s=s.replace('    isOK = modShared.CheckFS_OT(d, FS_OT)\n    isOK = modShared.CheckFS_SL(d, FS_SL)\n    isOK = modShared.CheckFS_BC(d, FS_BC, e, qMax, qMin)', '    isOK = modShared.CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, FS_OT, FS_SL, FS_BC)')
s=s.replace('"runtime_sec,timestamp"', '"runtime_sec,timestamp,seed,evaluations,status"')
s=s.replace('Format(Now, "yyyy-mm-dd hh:nn:ss")', 'Format(Now, "yyyy-mm-dd hh:nn:ss") & "," & RunSeed & "," & EvaluationCount & "," & RunStatus')
s=s.replace('filePath = "D:\\batch_step3_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"', 'filePath = BatchOutputPath("batch_step3")')
s=s.replace('filePath = "D:\\batch_step3_A3_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"', 'filePath = BatchOutputPath("batch_step3_A3")')
s+='''
Private Function BatchOutputPath(stem As String) As String
    Dim p As String, n As Long, prefix As String
    If Dir$(App.Path & "\\results", vbDirectory) = "" Then MkDir App.Path & "\\results"
    prefix = App.Path & "\\results\\" & stem & "-" & Format$(Now, "yyyymmdd-hhnnss")
    p = prefix & ".csv"
    Do While Len(Dir$(p)) > 0
        n = n + 1: p = prefix & "-" & n & ".csv"
    Loop
    BatchOutputPath = p
End Function
'''
write('modBatch.bas',s)
