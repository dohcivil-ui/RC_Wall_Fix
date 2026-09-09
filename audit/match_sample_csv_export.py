from pathlib import Path
import csv
import hashlib
import json

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'audit/sample-csv-export'
SAMPLES = Path(r'D:\rc-rt-optimize-v2\vb6_samples')
OUT.mkdir(exist_ok=False)
(OUT / 'before').mkdir()

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inventory(folder):
    return {str(p.relative_to(folder)): sha(p) for p in folder.rglob('*') if p.is_file()}

root_before = {p.name: sha(p) for p in ROOT.iterdir() if p.suffix.lower() in ('.bas', '.frm', '.frx', '.vbp', '.exe')}
(OUT / 'root-before.json').write_text(json.dumps(root_before, indent=2), encoding='utf-8')
(OUT / 'result-csv-before.json').write_text(json.dumps(inventory(ROOT / 'result_csv'), indent=2), encoding='utf-8')
samples = {}
for method in ('BA', 'HCA'):
    for kind in ('accept', 'loopPrice'):
        p = SAMPLES / f'{kind}-{method}-H3-240.csv'
        data = p.read_bytes()
        with p.open(encoding='utf-8-sig', newline='') as f:
            rows = list(csv.reader(f))
        samples[p.name] = {'sha256': sha(p), 'header': rows[0], 'row_count': len(rows) - 1, 'first_no': rows[1][0], 'last_no': rows[-1][0], 'crlf_count': data.count(b'\r\n'), 'lf_count': data.count(b'\n')}
(OUT / 'samples-before.json').write_text(json.dumps(samples, indent=2), encoding='utf-8')

names = ('modBA.bas', 'modHillClimbing.bas', 'Form1.frm', 'RC_RT_HCA_v2.vbp', 'modTrialExport.bas')
for name in names:
    (OUT / 'before' / name).write_bytes((ROOT / name).read_bytes())

for name in ('modBA.bas', 'modHillClimbing.bas'):
    path = ROOT / name
    data = path.read_bytes()
    old = b'Replace$(CStr(wallHeight), ",", "."), csv'
    new = b'Replace$(CStr(wallHeight), ",", ".") & "-" & CStr(currentMaterial.fc), csv'
    assert data.count(old) == 2, name
    path.write_bytes(data.replace(old, new))

path = ROOT / 'Form1.frm'
data = path.read_bytes()
for line in (
    b'    Call BeginTrialAcceptExport(CLng(numTrials))\r\n',
    b'        Call AppendTrialAcceptExport\r\n',
    b'    AddResultLine "accept CSV (all trials): " & LastTrialAcceptCSVPath\r\n',
):
    assert data.count(line) == 2, line
    data = data.replace(line, b'')
old = (b'    If Len(LastTrialAcceptCSVPath) > 0 Then\r\n'
       b'        AddResultLine "Saved trial data (may be incomplete): " & LastTrialAcceptCSVPath\r\n'
       b'    End If\r\n')
new = (b'    If Len(LastAcceptCSVPath) > 0 Then\r\n'
       b'        AddResultLine "Last saved accept CSV: " & LastAcceptCSVPath\r\n'
       b'    End If\r\n'
       b'    If Len(LastLoopCSVPath) > 0 Then\r\n'
       b'        AddResultLine "Last saved loopPrice CSV: " & LastLoopCSVPath\r\n'
       b'    End If\r\n')
assert data.count(old) == 1
path.write_bytes(data.replace(old, new))

path = ROOT / 'RC_RT_HCA_v2.vbp'
data = path.read_bytes()
line = b'Module=modTrialExport; modTrialExport.bas\r\n'
assert data.count(line) == 1
path.write_bytes(data.replace(line, b''))

# This module was created by the previous export change and is no longer loaded.
# Preserve its complete bytes under this task's before/ before removing the root copy.
path = (ROOT / 'modTrialExport.bas').resolve()
assert path.parent == ROOT.resolve()
assert path.read_bytes() == (OUT / 'before/modTrialExport.bas').read_bytes()
path.unlink()

for area in ('production', 'probe'):
    (OUT / area).mkdir()
lines = []
for line in (ROOT / 'RC_RT_HCA_v2.vbp').read_text().splitlines():
    if line.startswith('Module='):
        key, value = line.split(';', 1)
        line = key + '; ' + str(ROOT / value.strip())
    elif line.startswith('Form='):
        line = 'Form=' + str(ROOT / line[5:].strip())
    lines.append(line)
(OUT / 'production/Production.vbp').write_bytes(('\r\n'.join(lines) + '\r\nExeName32="RC_RT_HCA_v2.exe"\r\n').encode('ascii'))
print(json.dumps({'sample_structure': samples, 'updated': list(names), 'optimizer_runs': 0}, indent=2))
