"""Native CSV fixture replay only. Never call an optimizer or write result_csv."""
from pathlib import Path
import hashlib
import json
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'audit/best-trial-export-2026-09-10'
VB6 = Path(r'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE')

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inventory(folder):
    return {str(p.relative_to(folder)): digest(p) for p in folder.rglob('*') if p.is_file()}

PROBE = '''Attribute VB_Name = "BestTrialProbe"
Option Explicit
Private logFile As Integer
Private checks As Long, failures As Long

Private Sub Verify(ByVal label As String, ByVal ok As Boolean)
    checks = checks + 1
    Print #logFile, IIf(ok, "PASS ", "FAIL ") & label
    If Not ok Then failures = failures + 1
End Sub

Private Function ReadBytes(ByVal path As String) As String
    Dim f As Integer
    If Len(Dir$(path)) = 0 Then Exit Function
    f = FreeFile
    Open path For Binary Access Read As #f
    ReadBytes = Space$(LOF(f))
    If LOF(f) > 0 Then Get #f, , ReadBytes
    Close #f
End Function

Private Sub ResetBatch(ByVal method As String)
    If method = "BA" Then InitLoopCounter_BA Else InitLoopCounter
End Sub

Private Sub Replay(ByVal method As String, ByVal trial As Long, ByVal cost As Double, ByVal bestLoop As Long, ByVal valid As Boolean)
    Dim i As Long, price As String, better As Boolean
    If method = "BA" Then InitCSVExport_BA Else InitCSVExport
    RunAlgorithm = method: RunTrial = trial: RunFolder = ""
    RunBest.IsValid = valid: RunBestCost = cost: RunBestEvaluation = 0
    If valid Then RunBestEvaluation = bestLoop + 1
    For i = 0 To 5
        better = valid And i = bestLoop
        price = CsvPrice(900 + trial)
        If better Then price = CsvPrice(cost)
        If method = "BA" Then
            LogIteration_BA i, price, better, better
        Else
            LogIteration i, price, better, better
        End If
    Next i
    FinishSearch
    If method = "BA" Then LogLoopResult_BA cost Else LogLoopResult cost
End Sub

Private Sub SaveBatch(ByVal method As String)
    If method = "BA" Then SaveLoopPriceCSV_BA H Else SaveLoopPriceCSV H
End Sub

Private Sub CheckMethod(ByVal method As String)
    Dim chosen As String, lastTrace As String, expected As String, metadata As String
    H = 3: currentMaterial.fc = 240
    ResetBatch method
    Replay method, 1, 120, 3, True
    Replay method, 2, 100, 4, True
    Replay method, 3, 100, 2, True
    chosen = ReadBytes(LastAcceptCSVPath)
    Replay method, 4, 110, 1, True
    Replay method, 5, 100, 2, True
    Replay method, 6, NO_SOLUTION_COST, 0, False
    lastTrace = ReadBytes(LastAcceptCSVPath)
    Verify method & " fixture distinguishes last from best", chosen <> lastTrace
    SaveBatch method
    Verify method & " chosen whole trace is trial 3 (price then loop; stable tie)", ReadBytes(LastAcceptCSVPath) = chosen
    expected = "No.,Loop,BestPrice" & vbCrLf & "1,3,120.00" & vbCrLf & "2,4,100.00" & vbCrLf & "3,2,100.00" & vbCrLf & "4,1,110.00" & vbCrLf & "5,2,100.00" & vbCrLf & "6,0," & vbCrLf
    Verify method & " all summary rows and no-solution retained", ReadBytes(LastLoopCSVPath) = expected
    metadata = RESULT_CSV_ROOT & "\\selectedTrial-" & method & "-H3-240.csv"
    expected = "No.,Loop,BestPrice,Trials,Status" & vbCrLf & "3,2,100.00,6,SELECTED" & vbCrLf
    Verify method & " selected trial is identified", ReadBytes(metadata) = expected

    ResetBatch method
    Replay method, 1, 130, 0, True
    chosen = ReadBytes(LastAcceptCSVPath)
    Replay method, 2, 130, 1, True
    SaveBatch method
    Verify method & " resets between batches and accepts loop zero", ReadBytes(LastAcceptCSVPath) = chosen
    Verify method & " refreshed selection metadata", InStr(ReadBytes(metadata), "1,0,130.00,2,SELECTED") > 0

    H = 4: currentMaterial.fc = 280
    ResetBatch method
    Replay method, 1, NO_SOLUTION_COST, 0, False
    Replay method, 2, NO_SOLUTION_COST, 0, False
    chosen = ReadBytes(LastAcceptCSVPath)
    SaveBatch method
    Verify method & " all-invalid keeps diagnostic last trace", ReadBytes(LastAcceptCSVPath) = chosen
    metadata = RESULT_CSV_ROOT & "\\selectedTrial-" & method & "-H4-280.csv"
    expected = "No.,Loop,BestPrice,Trials,Status" & vbCrLf & ",,,2,NO_SOLUTION" & vbCrLf
    Verify method & " no stale winner for new H/fc with no solution", ReadBytes(metadata) = expected

    ResetBatch method
    Replay method, 1, 140, 5, True
    chosen = ReadBytes(LastAcceptCSVPath)
    SaveBatch method
    Verify method & " single-trial trace preserved", ReadBytes(LastAcceptCSVPath) = chosen
    Verify method & " single-trial metadata", InStr(ReadBytes(metadata), "1,5,140.00,1,SELECTED") > 0
End Sub

Public Sub Main()
    On Error GoTo Failed
    logFile = FreeFile
    Open App.Path & "\\runtime.txt" For Output As #logFile
    If RESULT_CSV_ROOT <> App.Path & "\\output" Then Err.Raise 5, , "Unsafe probe output"
    CheckMethod "BA"
    CheckMethod "HCA"
    Print #logFile, "CHECKS=" & checks
    Print #logFile, "FAILURES=" & failures
    Print #logFile, "OPTIMIZER_RUNS=0"
    Close #logFile
    Exit Sub
Failed:
    Print #logFile, "FATAL=" & Err.Number & ": " & Err.Description
    Close #logFile
End Sub
'''

def prepare():
    OUT.mkdir(exist_ok=False)
    (OUT / 'before').mkdir()
    for name in ('Form1.frm', 'modBA.bas', 'modHillClimbing.bas', 'HANDOFF.md', 'NEXT_CHAT_PROMPT.md'):
        (OUT / 'before' / name).write_bytes((ROOT / name).read_bytes())
    sources = {p.name: digest(p) for p in ROOT.iterdir() if p.suffix.lower() in ('.bas', '.frm', '.frx', '.vbp', '.exe')}
    (OUT / 'sources-before.json').write_text(json.dumps(sources, indent=2))
    (OUT / 'result-csv-before.json').write_text(json.dumps(inventory(ROOT / 'result_csv'), indent=2))
    print('Snapshot saved; result_csv read only.')

def build_project(folder, probe):
    folder.mkdir(exist_ok=True)
    project = (ROOT / 'RC_RT_HCA_v2.vbp').read_bytes().decode('latin1')
    for line in project.splitlines():
        if line.startswith('Form='):
            project = project.replace(line, 'Form=' + str(ROOT / line[5:]))
        elif line.startswith('Module='):
            name, path = line.split('; ', 1)
            project = project.replace(line, name + '; ' + str(ROOT / path))
    if probe:
        shared = (ROOT / 'modShared.bas').read_bytes()
        original = ('Public Const RESULT_CSV_ROOT As String = "' + str(ROOT / 'result_csv') + '"').encode()
        redirected = ('Public Const RESULT_CSV_ROOT As String = "' + str(folder / 'output') + '"').encode()
        assert shared.count(original) == 1
        (folder / 'modShared.bas').write_bytes(shared.replace(original, redirected))
        project = project.replace(str(ROOT / 'modShared.bas'), str(folder / 'modShared.bas'))
        project = project.replace('Startup="Form1"', 'Startup="Sub Main"')
        project += 'Module=BestTrialProbe; ' + str(folder / 'BestTrialProbe.bas') + '\r\n'
        (folder / 'BestTrialProbe.bas').write_bytes(PROBE.replace('\n', '\r\n').encode('ascii'))
    project += 'ExeName32="Check.exe"\r\n'
    (folder / 'Check.vbp').write_bytes(project.encode('latin1'))
    log = folder / 'compile.log'
    if log.exists():
        raise FileExistsError('Use a new phase directory so compiler logs cannot be mixed.')
    def ps(value):
        return "'" + str(value).replace("'", "''") + "'"
    arguments = f'/make "{folder / "Check.vbp"}" /out "{log}" /outdir "{folder}"'
    command = f'$exportCompiler = Start-Process -FilePath {ps(VB6)} -ArgumentList {ps(arguments)} -WindowStyle Hidden -PassThru; if (-not $exportCompiler.WaitForExit(45000)) {{ throw "Compiler still running" }}; exit $exportCompiler.ExitCode'
    subprocess.run(['powershell', '-NoProfile', '-Command', command], timeout=55, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    result = log.read_text(errors='replace')
    print(result.strip())
    assert 'succeeded' in result.lower(), result

def run(phase):
    folder = OUT / phase
    build_project(folder, True)
    subprocess.run([str(folder / 'Check.exe')], timeout=30, creationflags=subprocess.CREATE_NO_WINDOW)
    print((folder / 'runtime.txt').read_text())
    assert inventory(ROOT / 'result_csv') == json.loads((OUT / 'result-csv-before.json').read_text())
    print('Protected result_csv inventory and bytes unchanged.')

def verify():
    checks = []
    def check(name, passed):
        checks.append({'check': name, 'passed': bool(passed)})
        assert passed, name
    before = json.loads((OUT / 'sources-before.json').read_text())
    after = {p.name: digest(p) for p in ROOT.iterdir() if p.suffix.lower() in ('.bas', '.frm', '.frx', '.vbp', '.exe')}
    check('only three intended root sources changed', set(before) == set(after) and {p for p in before if before[p] != after[p]} == {'Form1.frm', 'modBA.bas', 'modHillClimbing.bas'})
    procedure = re.compile(rb'(?m)^(?:Public|Private) (?:Sub|Function) (\w+)\([^\r\n]*(?:.|\r|\n)*?^End (?:Sub|Function)\r?$', re.MULTILINE)
    for name, suffix in [('modBA.bas', '_BA'), ('modHillClimbing.bas', '')]:
        old = (OUT / 'before' / name).read_bytes()
        new = (ROOT / name).read_bytes()
        old_methods = {m.group(1).decode(): m.group() for m in procedure.finditer(old)}
        new_methods = {m.group(1).decode(): m.group() for m in procedure.finditer(new)}
        expected_changed = {'InitLoopCounter' + suffix, 'LogLoopResult' + suffix, 'SaveLoopPriceCSV' + suffix}
        check(name + ' changes only export procedures', {n for n in old_methods if old_methods[n] != new_methods[n]} == expected_changed)
        check(name + ' adds only selection export procedure', set(new_methods) - set(old_methods) == {'SaveSelectedTrialCSV' + suffix})
        check(name + ' all search procedures are checked', any('Optimization' in n for n in old_methods))
        check(name + ' retains CRLF and no BOM', not new.startswith(b'\xef\xbb\xbf') and b'\n' not in new.replace(b'\r\n', b''))
    form = (ROOT / 'Form1.frm').read_bytes()
    for module, suffix in [('modBA', '_BA'), ('modHillClimbing', '')]:
        block = f'''    If globalBestTrial > 0 Then
        AddResultLine "accept CSV (best trial " & globalBestTrial & "): " & LastAcceptCSVPath
    Else
        AddResultLine "accept CSV (last trial; no feasible trial): " & LastAcceptCSVPath
    End If
    AddResultLine "Selected trial details: " & {module}.SelectedTrialCSVPath{suffix}
'''.replace('\n', '\r\n').encode()
        check(module + ' UI label replacement exists once', form.count(block) == 1)
        form = form.replace(block, b'    AddResultLine "accept CSV (last trial): " & LastAcceptCSVPath\r\n')
    check('Form1 only reporting labels changed', form == (OUT / 'before/Form1.frm').read_bytes())
    protected = inventory(ROOT / 'result_csv')
    check('all result_csv paths and bytes preserved', protected == json.loads((OUT / 'result-csv-before.json').read_text()))
    for phase in ('red', 'green', 'production'):
        check(phase + ' native compile passed', 'succeeded' in (OUT / phase / 'compile.log').read_text())
    check('regression reproduced before fix', 'FAILURES=12' in (OUT / 'red/runtime.txt').read_text())
    runtime = (OUT / 'green/runtime.txt').read_text()
    check('20 native replay checks passed after fix', 'CHECKS=20' in runtime and 'FAILURES=0' in runtime and 'FATAL=' not in runtime)
    for phase in ('red', 'green'):
        probe = (OUT / phase / 'BestTrialProbe.bas').read_text()
        check(phase + ' probe has no search/evaluation/random calls', not re.search(r'BisectionOptimization|HillClimbingOptimization|EvaluateCandidate|BeginSearch|\bRnd\b|Randomize', probe))
        original = ('Public Const RESULT_CSV_ROOT As String = "' + str(ROOT / 'result_csv') + '"').encode()
        redirected = ('Public Const RESULT_CSV_ROOT As String = "' + str(OUT / phase / 'output') + '"').encode()
        check(phase + ' probe redirects only output path', (OUT / phase / 'modShared.bas').read_bytes() == (ROOT / 'modShared.bas').read_bytes().replace(original, redirected))
    result = {'checks': checks, 'independent_checks': len(checks), 'native_replay_checks': 20, 'optimizer_runs': 0, 'protected_result_files': len(protected), 'root_after': after}
    (OUT / 'verification.json').write_text(json.dumps(result, indent=2))
    print(f'PASS {len(checks)} independent checks + 20 native replay checks; {len(protected)} protected files unchanged; optimizer runs = 0.')

if __name__ == '__main__':
    action = sys.argv[1]
    if action == 'prepare': prepare()
    elif action == 'verify': verify()
    elif action == 'compile': build_project(OUT / 'production', False)
    else: run(action)
