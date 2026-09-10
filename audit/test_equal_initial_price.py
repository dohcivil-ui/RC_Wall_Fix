"""Exercise both real entry points with budget=1: initial evaluation only."""
from pathlib import Path
import importlib.util
import hashlib
import json
import csv
import subprocess

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'audit/equal-initial-price-2026-09-10'

PROBE = '''Attribute VB_Name = "BestTrialProbe"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private logFile As Integer, calls As Long, initialEvents As Long
Private modeName As String

Private Function RawDouble(ByVal value As Double) As String
    Dim bytes(0 To 7) As Byte, i As Integer
    CopyMemory bytes(0), value, 8
    For i = 0 To 7
        RawDouble = RawDouble & Right$("0" & Hex$(bytes(i)), 2)
    Next i
End Function

Public Sub RecordInitial(d As Design, entry As String, ByVal optimizerCost As Double, ByVal exportedPrice As String, ByVal valid As Boolean)
    Dim r As ProjectDetail, reason As String, quantityPrice As Double, signature As String
    If entry <> "initial" Or EvaluationCount <> 1 Or EvaluationBudget <> 1 Then Err.Raise 5, , "Unexpected search step"
    reason = LastValidationReason
    If Not ProjectQuantityCost(d, r) Then Err.Raise 5, , "Cannot price initial quantities"
    quantityPrice = r.Cost
    signature = RawDouble(d.tt) & RawDouble(d.tb) & RawDouble(d.TBase) & RawDouble(d.Base) & RawDouble(d.LToe) & RawDouble(d.LHeel)
    signature = signature & ":" & d.ASst_DB & ":" & d.ASst_Sp & ":" & d.AStoe_DB & ":" & d.AStoe_Sp & ":" & d.ASheel_DB & ":" & d.ASheel_Sp
    initialEvents = initialEvents + 1
    Print #logFile, Join(Array(modeName, RunAlgorithm, H, currentMaterial.fc, EvaluationCount, entry, CsvNumber(d.tt), CsvNumber(d.tb), CsvNumber(d.TBase), CsvNumber(d.Base), CsvNumber(d.LToe), CsvNumber(d.LHeel), WP_DB(d.ASst_DB), CsvNumber(WP_SP(d.ASst_Sp)), WP_DB(d.AStoe_DB), CsvNumber(WP_SP(d.AStoe_Sp)), WP_DB(d.ASheel_DB), CsvNumber(WP_SP(d.ASheel_Sp)), CsvNumber(quantityPrice), exportedPrice, RawDouble(quantityPrice), signature, valid, reason, CsvNumber(optimizerCost)), ",")
End Sub

Private Sub CheckOne(method As String, maximum As Boolean)
    Dim d As Design, savedPath As String
    If maximum Then modeName = "catalog_max" Else modeName = "current"
    initialEvents = 0
    If method = "BA" Then
        d = BisectionOptimization(1, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, currentMaterial, _
            useSharedInit:=maximum, sharedTt:=TT_MAX, sharedTb:=tb_max, sharedTBase:=TBase_max, sharedBase:=BASE_MAX, sharedLToe:=LTOE_MAX, _
            sharedStemDB:=DB_MAX, sharedStemSP:=SP_MIN, sharedToeDB:=DB_MAX, sharedToeSP:=SP_MIN, sharedHeelDB:=DB_MAX, sharedHeelSP:=SP_MIN)
    Else
        d = HillClimbingOptimization(1, H, H1, gamma_soil, gamma_concrete, phi, mu, qa, cover, currentMaterial, _
            useSharedInit:=maximum, sharedTt:=TT_MAX, sharedTb:=tb_max, sharedTBase:=TBase_max, sharedBase:=BASE_MAX, sharedLToe:=LTOE_MAX, _
            sharedStemDB:=DB_MAX, sharedStemSP:=SP_MIN, sharedToeDB:=DB_MAX, sharedToeSP:=SP_MIN, sharedHeelDB:=DB_MAX, sharedHeelSP:=SP_MIN)
    End If
    If initialEvents <> 1 Or EvaluationCount <> 1 Then Err.Raise 5, , "Initial evaluation count mismatch"
    savedPath = App.Path & "\\" & modeName & "-accept-" & method & "-H" & H & "-" & currentMaterial.fc & ".csv"
    FileCopy LastAcceptCSVPath, savedPath
    calls = calls + 1
End Sub

Public Sub Main()
    Dim height As Integer, fcValue As Integer, f As Integer
    On Error GoTo Failed
    If RESULT_CSV_ROOT <> App.Path & "\\output" Then Err.Raise 5, , "Unsafe result path"
    Call InitializeArrays
    Call EnableProjectChecks
    H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4
    phi = 30: mu = 0.6: qa = 30: cover = 0.075
    logFile = FreeFile
    Open App.Path & "\\initial-comparison.csv" For Output As #logFile
    Print #logFile, "mode,method,H,fc,evaluations,entry,tt,tb,TBase,B,toe,heel,stem_DB,stem_SP,toe_DB,toe_SP,heel_DB,heel_SP,quantity_price,exported_price,price_hex,design_signature,valid,reason,optimizer_cost"
    For height = 3 To 5
        H = height: fcValue = 240 + (height - 3) * 40
        currentMaterial = GetSD40Material(fcValue, GetConcretePrice(fcValue), STEEL_PRICE_SD40)
        CheckOne "BA", False
        CheckOne "HCA", False
        CheckOne "BA", True
        CheckOne "HCA", True
    Next height
    Close #logFile
    If calls <> 12 Then Err.Raise 5, , "Expected 12 initial-only calls"
    f = FreeFile
    Open App.Path & "\\runtime.txt" For Output As #f
    Print #f, "INITIAL_ONLY_CALLS=" & calls
    Print #f, "NEIGHBOR_OR_MIDPOINT_EVALUATIONS=0"
    Print #f, "FAILURES=0"
    Close #f
    Exit Sub
Failed:
    Dim errorText As String
    errorText = Err.Number & ": " & Err.Description
    f = FreeFile
    Open App.Path & "\\fatal.txt" For Output As #f
    Print #f, errorText
    Close #f
End Sub
'''

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def inventory(folder):
    return {str(p.relative_to(folder)):sha(p) for p in folder.rglob('*') if p.is_file()}

def main():
    OUT.mkdir(exist_ok=False)
    before={p.name:sha(p) for p in ROOT.iterdir() if p.is_file() and p.suffix.lower() in ('.bas','.frm','.frx','.vbp','.exe')}
    protected=inventory(ROOT/'result_csv')
    (OUT/'sources-before.json').write_text(json.dumps(before,indent=2))
    (OUT/'result-csv-before.json').write_text(json.dumps(protected,indent=2))
    sources=OUT/'sources'
    sources.mkdir()
    for name in before:
        if Path(name).suffix.lower()!='.exe':
            (sources/name).write_bytes((ROOT/name).read_bytes())
    p=sources/'modShared.bas'
    data=p.read_bytes()
    old=('Public Const RESULT_CSV_ROOT As String = "'+str(ROOT/'result_csv')+'"').encode()
    new=('Public Const RESULT_CSV_ROOT As String = "'+str(sources/'result_csv')+'"').encode()
    assert data.count(old)==1
    data=data.replace(old,new)
    anchor=b'    EvaluateCandidate = ok\r\n'
    assert data.count(anchor)==1
    data=data.replace(anchor,b'    Call BestTrialProbe.RecordInitial(d, entry, candidateCost, exportPrice, ok)\r\n'+anchor)
    p.write_bytes(data)
    spec=importlib.util.spec_from_file_location('native_compile',ROOT/'audit/verify_best_trial_export.py')
    native=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(native)
    native.ROOT=sources
    native.PROBE=PROBE
    native.build_project(OUT/'native',True)
    subprocess.run([str(OUT/'native/Check.exe')],timeout=25,check=True,creationflags=subprocess.CREATE_NO_WINDOW)
    assert not (OUT/'native/fatal.txt').exists(), (OUT/'native/fatal.txt').read_text() if (OUT/'native/fatal.txt').exists() else ''
    assert 'FAILURES=0' in (OUT/'native/runtime.txt').read_text()
    rows=list(csv.DictReader((OUT/'native/initial-comparison.csv').open()))
    assert len(rows)==12
    pairs=[]
    for h,fc in ((3,240),(4,280),(5,320)):
        for mode in ('current','catalog_max'):
            matched=[r for r in rows if r['H']==str(h) and r['mode']==mode]
            assert len(matched)==2 and {r['method'] for r in matched}=={'BA','HCA'}
            assert {k:v for k,v in matched[0].items() if k!='method'}=={k:v for k,v in matched[1].items() if k!='method'}
            accept=[]
            for method in ('BA','HCA'):
                p=OUT/'native'/f'{mode}-accept-{method}-H{h}-{fc}.csv'
                accept.append(p.read_bytes())
                parsed=list(csv.DictReader(p.open()))
                assert len(parsed)==1 and parsed[0]['No.']=='0'
                price_column='Passed and Better value' if matched[0]['valid']=='True' else 'Rejected'
                assert parsed[0][price_column]==matched[0]['exported_price']
            assert accept[0]==accept[1]
            pairs.append({'H':h,'fc':fc,'mode':mode,'price':matched[0]['quantity_price'],
                          'exact_price_equal':True,'all_dimensions_and_rebars_equal':True,
                          'initial_csv_bytes_equal':True,'valid':matched[0]['valid'],'reason':matched[0]['reason']})
    assert {name:sha(ROOT/name) for name in before}==before
    assert inventory(ROOT/'result_csv')==protected
    result={'pairs':pairs,'initial_only_entry_point_calls':12,'neighbor_midpoint_evaluations':0,
            'root_unchanged':True,'protected_result_files':len(protected),'protected_results_unchanged':True}
    (OUT/'verification.json').write_text(json.dumps(result,indent=2))
    print(json.dumps(result,indent=2))

if __name__=='__main__':
    main()
