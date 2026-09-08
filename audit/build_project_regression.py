from pathlib import Path
P=Path(__file__).resolve().parent
def put(name,s): (P/name).write_bytes(s.replace('\n','\r\n').encode('ascii'))
s=(P.parent/'RC_RT_HCA_v2.vbp').read_bytes().decode('latin1').replace('Startup="Form1"','Startup="Sub Main"')
s=s.replace('Name="Project1"','Name="ProjectChecksRegression"').replace('Form=Form1.frm','Form=..\\Form1.frm').replace('Form=frmBestDesign.frm','Form=..\\frmBestDesign.frm')
import re
s=re.sub(r'(?m)^(Module=[^;]+; )([^\r\n]+)',r'\1..\\\2',s)
s+='\r\nModule=ProjectRegression; ProjectRegression.bas\r\nExeName32="ProjectRegression.exe"\r\n'
(P/'ProjectRegression.vbp').write_bytes(s.encode('latin1'))
put('ProjectRegression.bas','''Attribute VB_Name = "ProjectRegression"
Option Explicit
Private output As Integer, detailFile As Integer, failures As Long, checks As Long
Private Sub Verify(label As String, ok As Boolean)
    checks = checks + 1
    If Not ok Then failures = failures + 1
    Print #output, IIf(ok, "OK: ", "FAIL: ") & label
End Sub
Private Function Fixture(height As Double, top As Double, bottom As Double, base As Double, width As Double, toe As Double) As Design
    Dim d As Design
    H = height: d.tt = top: d.tb = bottom: d.TBase = base: d.Base = width: d.LToe = toe: d.LHeel = width - toe - bottom
    d.ASst_DB = 101: d.ASst_Sp = 110
    d.AStoe_DB = 101: d.AStoe_Sp = 110: d.ASheel_DB = 101: d.ASheel_Sp = 110
    Fixture = d
End Function
Private Sub SaveCase(label As String, d As Design, expected As Boolean)
    Dim r As ProjectDetail, ok As Boolean, ot As Double, sl As Double, bc As Double, i As Integer, f As Integer
    ok = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, ot, sl, bc)
    Print #output, label & " accepted=" & ok & "; reason=" & LastValidationReason
    Call Verify(label & " expected acceptance", ok = expected)
    Print #detailFile, label & "," & H & "," & d.tt & "," & d.tb & "," & d.TBase & "," & d.Base & "," & d.LToe & "," & ok & "," & LastValidationReason
    If ok Then
        Call Verify(label & " deterministic recheck", CheckProjectDesign(d, r))
        Call Verify(label & " price from same detail", Abs(CalculateCost(d) - r.Cost) < 0.000001)
        Call Verify(label & " anchorage and lap allowances excluded", r.Ld(0) = 0 And r.Ld(1) = 0 And r.Ld(2) = 0)
        Call Verify(label & " concrete plus main steel only", Abs(r.Cost - (r.ConcreteVolume * currentMaterial.concretePrice + r.MainWeight * currentMaterial.SteelPrice)) < 0.000001)
        f = FreeFile: Open App.Path & "\\project-" & label & "-report.txt" For Output As #f
        Print #f, ProjectDesignReport(d, currentMaterial, "FIXTURE - not an optimizer minimum")
        Close #f
        f = FreeFile: Open App.Path & "\\project-" & label & "-detail.csv" For Output As #f
        Print #f, "part,DB,SP,depth,As,minimum,M,v,fc_bound,fs_bound,ld,length"
        For i = 0 To 2
            Print #f, i & "," & WP_DB(r.DB(i)) & "," & CsvNumber(WP_SP(r.SP(i))) & "," & CsvNumber(r.Depth(i)) & "," & CsvNumber(r.Steel(i)) & "," & CsvNumber(r.Minimum(i)) & "," & CsvNumber(r.Moment(i)) & "," & CsvNumber(r.Shear(i)) & "," & CsvNumber(r.FcBound(i)) & "," & CsvNumber(r.FsBound(i)) & "," & CsvNumber(r.Ld(i)) & "," & CsvNumber(r.Length(i))
        Next i
        Print #f, "totals," & CsvNumber(r.OT) & "," & CsvNumber(r.SL) & "," & CsvNumber(r.BC) & "," & CsvNumber(r.ConcreteVolume) & "," & CsvNumber(r.MainWeight) & "," & CsvNumber(0) & "," & CsvNumber(r.Cost)
        Close #f
    End If
End Sub
Public Sub Main()
    Dim summaryFile As Integer, summaryHeader As String, summaryRow As String, noSolutionSummary As String
    Dim sr As ProjectDetail
    Dim d As Design, weak As Design, mat As MaterialProperties, ot As Double, sl As Double, bc As Double, i As Integer, text As String
    On Error GoTo Fatal
    output = FreeFile: Open App.Path & "\\project-regression.txt" For Output As #output
    detailFile = FreeFile: Open App.Path & "\\project-fixtures.csv" For Output As #detailFile
    Print #detailFile, "case,H,tt,tb,TBase,B,toe,valid,reason"
    Call InitializeArrays
    Call EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 30: cover = 0.075
    mat = GetSD40Material(320, GetConcretePrice(320), 24)
    currentMaterial = mat: currentWSD = CalculateWSDParameters(4000, 320)
    Call Verify("project basis active", WSDCriteriaReady())
    Call Verify("ACI shear units", Abs(ProjectShearLimit() - 5.2175525032) < 0.000000001)
    d = Fixture(3, 0.25, 0.3, 0.45, 2.5, 0.8): SaveCase "H3_reverse_heel", d, False
    d = Fixture(3, 0.25, 0.3, 0.45, 2, 0.6): SaveCase "H3", d, True
    d = Fixture(4, 0.25, 0.4, 0.45, 2.5, 1): SaveCase "H4", d, True
    d = Fixture(5, 0.25, 0.4, 0.45, 3, 1): SaveCase "H5", d, True
    weak = Fixture(5, 0.2, 0.2, 0.3, 4, 1.2): weak.ASst_DB = 100: weak.ASst_Sp = 113
    Call Verify("reference stem height", Abs(H - weak.TBase - 4.7) < 0.000001)
    Call Verify("reference active moment", Abs(gamma_soil * CalculateKa() * (H - weak.TBase) ^ 3 / 6# - 10.3823) < 0.000001)
    Call Verify("reference full-passive moment", Abs(CalculateMomentStem(weak) - 9.7262) < 0.000001)
    SaveCase "DB12weak", weak, False
    weak = d: weak.tt = 0.2: SaveCase "secondary_excluded", weak, True
    weak = d: weak.TBase = 0.3: SaveCase "anchorage_excluded", weak, True
    weak = d: weak.Base = 1.5: weak.LHeel = weak.Base - weak.LToe - weak.tb: SaveCase "stability", weak, False
    weak = d: weak.tt = 0.075: SaveCase "invalid_top_depth", weak, False
    weak = Fixture(3, 0.1, 0.2, 0.3, 3, 0.6): H1 = 1.5
    weak.ASst_DB = 100: weak.ASst_Sp = 113
    Call Verify("interior fixture base-only passes", CheckSteelOK(CalculateMomentStem(weak), SectionDepth(weak.tb, weak.ASst_DB), weak.ASst_DB, weak.ASst_Sp))
    Call Verify("interior fixture full-height fails", Not ProjectStemStressCheck(weak, sr))
    H1 = 1.2
    ' Equal small budgets; no 30-trial run. Fixed seeds exercise native optimizers.
    d = BisectionOptimization(64, 5, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, mat, RandomSeed:=20260908, TrialNumber:=1)
    Call Verify("BA64 exact budget", EvaluationCount = 64)
    Print #output, "BA64: " & RunStatus & "; folder=" & RunFolder
    d = HillClimbingOptimization(64, 5, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, mat, RandomSeed:=20260908, TrialNumber:=1)
    Call Verify("HCA64 exact budget", EvaluationCount = 64)
    Print #output, "HCA64: " & RunStatus & "; folder=" & RunFolder
    ' Production Form_Load and its actual BA button, one four-evaluation smoke run.
    Load Form1: Form1.Show
    Call Verify("Form_Load enables project basis", ProjectChecksEnabled And WSDCriteriaReady())
    Form1.txtMaxIter.Text = "4": Form1.txtTrials.Text = "1": Form1.txtSeed.Text = "12345"
    Form1.txtH.Text = "5": Form1.txtH1.Text = "1.2"
    ' Intentionally inadequate bearing input: NO_SOLUTION independent of anchorage scope.
    Form1.txtQa.Text = "0.01"
    For i = 0 To Form1.cboConcreteStrength.ListCount - 1
        If Form1.cboConcreteStrength.List(i) = "320" Then Form1.cboConcreteStrength.ListIndex = i
    Next i
    Form1.cmdBA.Value = True
    For i = 0 To Form1.lstResults.ListCount - 1: text = text & Form1.lstResults.List(i) & vbCrLf: Next i
    Call Verify("actual UI uses project report", InStr(text, PROJECT_CHECK_BASIS) > 0)
    Call Verify("actual UI states excluded scope", InStr(text, "anchorage/laps excluded") > 0 And InStr(text, "formwork cost excluded") > 0)
    Call Verify("actual UI restores controls", Form1.cmdBA.Enabled And Form1.cmdRun.Enabled)
    Print #output, "GUI: " & text
    noSolutionSummary = ProjectTrialSummary
    summaryFile = FreeFile: Open ProjectTrialSummary For Input As #summaryFile
    Line Input #summaryFile, summaryHeader: Line Input #summaryFile, summaryRow
    Call Verify("no solution summary has blank price", InStr(summaryRow, ",NO_SOLUTION,4,,0,") > 0)
    Call Verify("summary contains one trial", EOF(summaryFile))
    Close #summaryFile
    Form1.txtMaxIter.Text = "64": Form1.txtSeed.Text = "20260908": Form1.txtQa.Text = "30"
    Form1.cmdBA.Value = True
    text = ""
    For i = 0 To Form1.lstResults.ListCount - 1: text = text & Form1.lstResults.List(i) & vbCrLf: Next i
    Call Verify("actual UI accepted detail", InStr(text, "PASS_IMPLEMENTED_PROJECT_CHECKS") > 0 And InStr(text, "secondary steel excluded") > 0 And InStr(text, "Stem horizontal EACH face:") = 0 And InStr(text, "Stem front vertical:") = 0 And InStr(text, "Base longitudinal EACH face:") = 0)
    Call Verify("new UI session preserves old summary", ProjectTrialSummary <> noSolutionSummary And Len(Dir$(noSolutionSummary)) > 0)
    summaryFile = FreeFile: Open ProjectTrialSummary For Input As #summaryFile
    Line Input #summaryFile, summaryHeader: Line Input #summaryFile, summaryRow
    Call Verify("accepted summary has actual price", InStr(summaryRow, ",SOLUTION_FOUND_CONFIGURED_CHECKS,64," & CsvNumber(RunBestCost) & ",") > 0)
    Close #summaryFile
    Print #output, "GUI64: " & text
    Unload Form1
    Print #output, "PROJECT checks=" & checks & "; failures=" & failures
    Close #detailFile: Close #output
    Exit Sub
Fatal:
    Print #output, "FATAL " & Err.Number & ": " & Err.Description
    Close #detailFile: Close #output
End Sub
''')
