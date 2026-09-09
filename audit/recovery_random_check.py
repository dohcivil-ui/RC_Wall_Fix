"""Prepare native VB6 checks; only diagnostic copies redirect CSV and log recovery."""
from pathlib import Path
import csv, json, re, sys

ROOT = Path(__file__).resolve().parent.parent

def prepare(folder):
    folder.mkdir(parents=True, exist_ok=True)
    def write(name, text):
        (folder/name).write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('latin1'))
    ba = (ROOT/'modBA.bas').read_bytes().decode('latin1')
    marker = '            RunRecoveryCount = RunRecoveryCount + 1'
    assert ba.count(marker) == 1
    write('modBA.bas', ba.replace(marker, marker + '\n' +
        '            Call RecoveryCheck.RecordRecovery(Currenttb, CurrentTBase, CurrentBase, Midtb, MidTBase, MidBase)'))
    shared = (ROOT/'modShared.bas').read_bytes().decode('latin1')
    shared = shared.replace(str(ROOT/'result_csv'), str(folder/'result_csv'))
    write('modShared.bas', shared)
    vbp = (ROOT/'RC_RT_HCA_v2.vbp').read_bytes().decode('latin1')
    vbp = re.sub(r'^(Module=[^;]+; )([^\r\n]+)', lambda m: m[1]+str((folder if m[2] in ('modBA.bas','modShared.bas') else ROOT)/m[2]), vbp, flags=re.M)
    vbp = re.sub(r'^Form=([^\r\n]+)', lambda m:'Form='+str(ROOT/m[1]), vbp, flags=re.M)
    vbp = vbp.replace('Startup="Form1"','Startup="Sub Main"')
    write('RecoveryCheck.vbp', vbp+'\nModule=RecoveryCheck; RecoveryCheck.bas\nExeName32="RecoveryCheck.exe"\n')
    write('RecoveryCheck.bas', '''Attribute VB_Name = "RecoveryCheck"
Option Explicit
Private logFile As Integer
Public Sub RecordRecovery(tb As Integer, thick As Integer, base As Integer, nextTb As Integer, nextThick As Integer, nextBase As Integer)
    Print #logFile, "RECOVERY=" & tb & "," & thick & "," & base & "," & nextTb & "," & nextThick & "," & nextBase
End Sub
Public Sub Main()
    Dim height As Integer, mat As MaterialProperties, d As Design
    On Error GoTo Failed
    Call InitializeArrays
    Call EnableProjectChecks
    logFile = FreeFile
    Open App.Path & "\\recovery.txt" For Output As #logFile
    For height = 3 To 5
        mat = GetSD40Material(120 + 40 * height, GetConcretePrice(120 + 40 * height), 24)
        d = BisectionOptimization(125, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 0.01, 0.075, mat, RandomSeed:=12345)
        Print #logFile, "NO_SOLUTION=" & RunBest.IsValid
        Print #logFile, "RUN=" & RunFolder
        d = HillClimbingOptimization(125, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 0.01, 0.075, mat, RandomSeed:=12345)
        Print #logFile, "RUN=" & RunFolder
        d = BisectionOptimization(500, CDbl(height), 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12345)
        Print #logFile, "RUN=" & RunFolder
    Next height
    Close #logFile
    Exit Sub
Failed:
    Print #logFile, "FATAL=" & Err.Description
    Close #logFile
End Sub
''')
    print(folder)

def verify(folder):
    log = (folder/'recovery.txt').read_text()
    assert 'FATAL=' not in log
    recoveries = [[int(x) for x in line[9:].split(',')] for line in log.splitlines() if line.startswith('RECOVERY=')]
    assert recoveries
    jumps = sum(row[:3] != row[3:] for row in recoveries)
    paths = [Path(line[4:]) for line in log.splitlines() if line.startswith('RUN=')]
    assert len(paths) == 9 and log.count('NO_SOLUTION=False') == 3
    traces = []
    for path in paths:
        with (path/'evaluations.csv').open(newline='') as f: rows=list(csv.DictReader(f))
        report = (path/'run.txt').read_text()
        assert len(rows) == int(re.search(r'Budget=(\d+)', report)[1])
        best=999999999.0
        for i,row in enumerate(rows,1):
            assert int(row['evaluation']) == i
            if row['valid']=='True': best=min(best,float(row['cost']))
            assert abs(float(row['best_cost'])-best)<1e-7
        traces.append(rows)
    if jumps == 0:
        # In impossible cases current stays at initial. Ignoring counted resets,
        # BA must now consume the exact same neighbor draws as HCA across recoveries.
        for offset in (0,3,6):
            ba=[r for r in traces[offset] if r['entry']=='neighbor']
            hca=traces[offset+1][1:]
            for a,b in zip(ba,hca):
                for key in a.keys()-{'evaluation','entry'}: assert a[key]==b[key],(offset,key,a,b)
    result={'recoveries':len(recoveries),'recovery_position_changes':jumps,'runs':len(paths),'exact_budgets_and_prefix_best':True}
    (folder/'verification.json').write_text(json.dumps(result,indent=2))
    print(json.dumps(result))
    assert jumps == 0, 'Recovery randomly relocates current instead of resuming the HCA neighborhood'

if __name__ == '__main__':
    mode, path = sys.argv[1:]
    (prepare if mode == 'prepare' else verify)(Path(path).resolve())
