"""Initial-only native checks, CSV replay, graph rendering and GUI compilation."""
from pathlib import Path
import importlib.util
import hashlib
import json
import subprocess
import argparse
from datetime import datetime

ROOT = Path(__file__).resolve().parents[2]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--out', type=Path, default=ROOT / ('audit/initial-graph-check-' + datetime.now().strftime('%Y%m%d-%H%M%S')))
OUT = parser.parse_args().out.resolve()
assert OUT.is_relative_to(ROOT / 'audit'), 'Fixtures must remain in audit.'
OUT.mkdir(exist_ok=False, parents=True)
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
before = {p.name:sha(p) for p in ROOT.iterdir() if p.is_file() and p.suffix.lower() in ('.bas','.frm','.frx','.vbp','.exe')}
protected_before = {str(p.relative_to(ROOT/'result_csv')):sha(p) for p in (ROOT/'result_csv').rglob('*') if p.is_file()}

def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result

initial = module('initial', ROOT / 'audit/test_equal_initial_price.py')
initial.OUT = OUT / 'initial-only'
initial.PROBE = initial.PROBE.replace('    quantityPrice = r.Cost', '''    quantityPrice = r.Cost
    If RawDouble(quantityPrice) <> RawDouble(RunInitialQuantityCost) Then Err.Raise 5, , "Initial display quantity mismatch"''')
initial.main()

native = module('native', ROOT / 'audit/verify_best_trial_export.py')
native.build_project(OUT / 'gui', False)

# Reuse selection/overwrite fixtures without any optimizer call.
native.build_project(OUT / 'csv-fixtures', True)
subprocess.run([str(OUT / 'csv-fixtures/Check.exe')], timeout=30, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
csv_runtime = (OUT / 'csv-fixtures/runtime.txt').read_text()
assert 'FAILURES=0' in csv_runtime, csv_runtime
print(csv_runtime)

# Render the existing confirmed curves, with no newly sampled proposals.
render = OUT / 'render'
render.mkdir(exist_ok=False)
source = OUT / 'initial-only/sources'
form = (ROOT / 'Form1.frm').read_bytes()
wrapper = '''
Public Sub FixtureDrawComparison(hca() As Double, ba() As Double, ByVal cost As Double)
    hcaStoredInitialCost = cost: baStoredInitialCost = cost
    Call DrawDualCostGraph(picGraph, hca, UBound(hca), 1, ba, UBound(ba), 1)
End Sub
'''
(source / 'Form1.frm').write_bytes(form + wrapper.replace('\n','\r\n').encode('ascii'))
shared = (ROOT / 'modShared.bas').read_bytes()
shared = shared.replace(str(ROOT / 'result_csv').encode(), str(source / 'result_csv').encode())
(source / 'modShared.bas').write_bytes(shared)
confirmation = ROOT / 'audit/ba-paired-win-20260910/adaptive-coupling-refinement/confirmation'
for h, fc in [(3,240),(4,280),(5,320)]:
    for method in ['BA','HCA']:
        name = f'accept-{method}-H{h}-{fc}.csv'
        (render / name).write_bytes((confirmation / f'H{h}-{fc}/native/output' / name).read_bytes())

native.ROOT = source
native.PROBE = '''Attribute VB_Name = "BestTrialProbe"
Option Explicit
Private Function LoadHistory(ByVal method As String, ByVal height As Integer, ByVal fc As Integer, history() As Double, ByRef best As Long) As Double
    Dim f As Integer, row As String, parts() As String, i As Long, cost As Double, initial As Double
    ReDim history(1 To 5000)
    cost = NO_SOLUTION_COST
    f = FreeFile
    Open App.Path & "\\accept-" & method & "-H" & height & "-" & fc & ".csv" For Input As #f
    Line Input #f, row
    Do While Not EOF(f)
        Line Input #f, row
        parts = Split(row, ",")
        i = CLng(parts(0)) + 1
        If i = 1 Then
            If Len(parts(1)) > 0 Then initial = Val(parts(1)) Else initial = Val(parts(3))
        End If
        If Len(parts(3)) > 0 Then cost = Val(parts(3)): best = i
        history(i) = cost
    Loop
    Close #f
    LoadHistory = initial
End Function
Public Sub Main()
    Dim height As Integer, fc As Integer, method As Variant, history() As Double, hca() As Double, ba() As Double
    Dim initial As Double, savedInitial As Double, best As Long, f As Integer
    On Error GoTo Failed
    Load Form1
    Form1.picGraph.Width = 14000: Form1.picGraph.Height = 8000
    Form1.picGraph.AutoRedraw = True
    For height = 3 To 5
        fc = 240 + (height - 3) * 40
        savedInitial = 0
        For Each method In Array("BA", "HCA")
            initial = LoadHistory(CStr(method), height, fc, history, best)
            If savedInitial > 0 And savedInitial <> initial Then Err.Raise 5, , "Graph initial prices differ"
            savedInitial = initial
            Call DrawCostGraph(Form1.picGraph, history, best, initial)
            If history(1) <> NO_SOLUTION_COST Then Err.Raise 5, , "Rejected history changed"
            SavePicture Form1.picGraph.Image, App.Path & "\\graph-" & method & "-H" & height & ".bmp"
            If method = "BA" Then ba = history Else hca = history
        Next method
        Call Form1.FixtureDrawComparison(hca, ba, initial)
        SavePicture Form1.picGraph.Image, App.Path & "\\compare-H" & height & ".bmp"
    Next height
    ' Feasible initial and no-feasible-history edge cases.
    ReDim history(1 To 2)
    history(1) = 100: history(2) = 90
    Call DrawCostGraph(Form1.picGraph, history, 2, 100)
    If history(1) <> 100 Or history(2) <> 90 Then Err.Raise 5, , "Feasible history changed"
    history(1) = NO_SOLUTION_COST: history(2) = NO_SOLUTION_COST
    Call DrawCostGraph(Form1.picGraph, history, 0, 100)
    If history(1) <> NO_SOLUTION_COST Then Err.Raise 5, , "No-solution history changed"
    f = FreeFile
    Open App.Path & "\\runtime.txt" For Output As #f
    Print #f, "EXISTING_CURVES=6; COMPARISONS=3; EDGE_CASES=2; OPTIMIZER_RUNS=0; FAILURES=0"
    Close #f
    Unload Form1
    Exit Sub
Failed:
    f = FreeFile
    Open App.Path & "\\fatal.txt" For Output As #f
    Print #f, Err.Number & ": " & Err.Description
    Close #f
    Unload Form1
End Sub
'''
native.build_project(render, True)
subprocess.run([str(render / 'Check.exe')], timeout=30, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
assert not (render / 'fatal.txt').exists(), (render / 'fatal.txt').read_text() if (render / 'fatal.txt').exists() else ''
assert 'FAILURES=0' in (render / 'runtime.txt').read_text()
print((render / 'runtime.txt').read_text())

after = {name: sha(ROOT / name) for name in before}
changed = [name for name in before if before[name] != after[name]]
assert changed == [], changed
assert after['modBA.bas'] == 'e6ab8b41b990fb1f2cb0133adff26c1df027c1e8d0d368671632f2bb5babb6d4'
protected = {str(p.relative_to(ROOT/'result_csv')):sha(p) for p in (ROOT/'result_csv').rglob('*') if p.is_file()}
assert protected == protected_before
for name in ['Form1.frm','modShared.bas','modGraphing.bas']:
    data = (ROOT / name).read_bytes()
    assert not data.startswith(b'\xef\xbb\xbf') and b'\n' not in data.replace(b'\r\n',b'')
result = {'GUI_compile':'passed','CSV_fixtures':csv_runtime.strip(),'initial_only_calls':12,'neighbor_evaluations':0,
          'replayed_confirmed_graphs':9,'graph_edge_cases':2,'optimizer_source_unchanged':True,
          'root_sources_unchanged':True,'source_hashes':after,'protected_result_files':len(protected),'result_csv_unchanged':True}
(OUT / 'verification.json').write_text(json.dumps(result,indent=2))
print('All checks passed; protected files unchanged.')
