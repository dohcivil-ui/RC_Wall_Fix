from pathlib import Path
import hashlib
import json
import re

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'audit/sample-csv-export'
SAMPLES = Path(r'D:\rc-rt-optimize-v2\vb6_samples')
checks = []

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def check(name, ok):
    checks.append({'check': name, 'pass': bool(ok)})
    assert ok, name

before = json.loads((OUT / 'root-before.json').read_text())
after = {p.name: sha(p) for p in ROOT.iterdir() if p.suffix.lower() in ('.bas', '.frm', '.frx', '.vbp', '.exe')}
check('only retired exporter removed; no root files added', set(before) - set(after) == {'modTrialExport.bas'} and not (set(after) - set(before)))
changed = {name for name in after if before[name] != after[name]}
check('only intended root files changed', changed == {'modBA.bas', 'modHillClimbing.bas', 'Form1.frm', 'RC_RT_HCA_v2.vbp'})
for name in ('modBA.bas', 'modHillClimbing.bas'):
    old = (OUT / 'before' / name).read_bytes()
    old_fragment = b'Replace$(CStr(wallHeight), ",", "."), csv'
    new_fragment = b'Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csv'
    check(name + ' only two filename suffixes changed; search code unchanged', old.count(old_fragment) == 2 and (ROOT / name).read_bytes() == old.replace(old_fragment, new_fragment))

old_form = (OUT / 'before/Form1.frm').read_bytes()
for line in (
    b'    Call BeginTrialAcceptExport(CLng(numTrials))\r\n',
    b'        Call AppendTrialAcceptExport\r\n',
    b'    AddResultLine "accept CSV (all trials): " & LastTrialAcceptCSVPath\r\n',
):
    assert old_form.count(line) == 2
    old_form = old_form.replace(line, b'')
old_message = (b'    If Len(LastTrialAcceptCSVPath) > 0 Then\r\n'
               b'        AddResultLine "Saved trial data (may be incomplete): " & LastTrialAcceptCSVPath\r\n'
               b'    End If\r\n')
new_message = (b'    If Len(LastAcceptCSVPath) > 0 Then\r\n'
               b'        AddResultLine "Last saved accept CSV: " & LastAcceptCSVPath\r\n'
               b'    End If\r\n'
               b'    If Len(LastLoopCSVPath) > 0 Then\r\n'
               b'        AddResultLine "Last saved loopPrice CSV: " & LastLoopCSVPath\r\n'
               b'    End If\r\n')
check('Form only removes combined hooks and updates error output paths', (ROOT / 'Form1.frm').read_bytes() == old_form.replace(old_message, new_message))
check('VBP only removes retired exporter reference', (ROOT / 'RC_RT_HCA_v2.vbp').read_bytes() == (OUT / 'before/RC_RT_HCA_v2.vbp').read_bytes().replace(b'Module=modTrialExport; modTrialExport.bas\r\n', b''))
for name in changed:
    data = (ROOT / name).read_bytes()
    check(name + ' retains CRLF and no BOM', not data.startswith(b'\xef\xbb\xbf') and b'\n' not in data.replace(b'\r\n', b''))

shared = (ROOT / 'modShared.bas').read_bytes()
check('budget and zero-based logging unchanged', after['modShared.bas'] == before['modShared.bas'] and b'LogIteration_BA EvaluationCount - 1' in shared and b'LogIteration EvaluationCount - 1' in shared)
probe_shared = (OUT / 'probe/modShared.bas').read_bytes()
old = b'Public Const RESULT_CSV_ROOT As String = "C:\\reserch 69\\RC_Wall_Fix\\result_csv"'
new = ('Public Const RESULT_CSV_ROOT As String = "' + str(OUT / 'probe/output') + '"').encode('ascii')
check('probe redirects only output path', probe_shared == shared.replace(old, new))
probe = (OUT / 'probe/SampleCsvProbe.bas').read_text()
check('probe invokes only export/log APIs, no optimizer or random draws', not re.search(r'BisectionOptimization|HillClimbingOptimization|EvaluateCandidate|BeginSearch|\bRnd\b|Randomize', probe))
runtime = (OUT / 'probe/runtime.txt').read_text()
check('all 27 native export replay checks passed', 'CHECKS=27' in runtime and 'FAILURES=0' in runtime and 'FATAL=' not in runtime and 'OPTIMIZER_RUNS=0' in runtime)

samples = json.loads((OUT / 'samples-before.json').read_text())
for name, item in samples.items():
    check(name + ' original sample unchanged', sha(SAMPLES / name) == item['sha256'])
    check(name + ' native export matches sample byte for byte', (OUT / 'probe/output' / name).read_bytes() == (SAMPLES / name).read_bytes())
check('no per-trial directories or combined output', not any(p.is_dir() for p in (OUT / 'probe/output').iterdir()) and not list((OUT / 'probe/output').glob('acceptRuns*')))
protected = json.loads((OUT / 'result-csv-before.json').read_text())
current = {str(p.relative_to(ROOT / 'result_csv')): sha(p) for p in (ROOT / 'result_csv').rglob('*') if p.is_file()}
check(f'all {len(protected)} user result_csv files and inventory unchanged', protected == current)
for kind in ('production', 'probe'):
    check(kind + ' VB6 compile succeeded', 'succeeded' in (OUT / kind / 'compile.log').read_text())
prod = (OUT / 'production/Production.vbp').read_text()
check('production starts real root Form1 and excludes probes/retired exporter', 'Startup="Form1"' in prod and str(ROOT / 'Form1.frm') in prod and 'SampleCsvProbe' not in prod and 'modTrialExport' not in prod)
result = {'check_count': len(checks), 'failures': 0, 'native_checks': 27, 'optimizer_runs': 0, 'protected_result_files': len(protected), 'checks': checks, 'root_after': after, 'production_exe_sha256': sha(OUT / 'production/RC_RT_HCA_v2.exe')}
(OUT / 'verification.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
print(f'PASS {len(checks)} independent checks + 27 native fixture-replay checks; {len(protected)} user result files unchanged; optimizer runs = 0.')
