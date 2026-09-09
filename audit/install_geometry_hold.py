"""Apply the exact tested candidate only after the user-authorized H5 gate."""
from pathlib import Path
import hashlib
import json

root = Path(__file__).resolve().parent.parent
evidence = root / 'audit/ba-geometry-hold-h5-fc320-20'
batch = root / 'audit/parameter-trials/h5-fc320-20'
candidate = root / 'audit/ba-geometry-hold/candidate-modBA.bas'
out = root / 'audit/ba-geometry-hold-installation'
def read(p): return json.loads(p.read_text(encoding='utf-8'))
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
summary = read(evidence / 'performance-summary.json')
ba, hca = summary['methods']['BA'], summary['methods']['HCA']
assert ba['target_attainment'] >= hca['target_attainment']
assert summary['relative_effort_reduction_percent'] >= 10
assert ba['mean_valid_final_cost'] <= hca['mean_valid_final_cost']
assert read(evidence / 'final-verification.json')['passed']
assert read(batch / 'vector-verification.json')['passed']
assert all(x['passed'] for x in read(batch / 'verification.json'))
before = read(evidence / 'root-sources-before.json')
for name, expected in before.items(): assert sha(root / name) == expected, 'Root changed: ' + name
assert sha(candidate) == read(evidence / 'candidate-hashes.json')['modBA.bas']
assert candidate.read_bytes() == (batch / 'modBA.bas').read_bytes()
out.mkdir(exist_ok=False)
backup = out / 'before'
backup.mkdir()
for name in ('modBA.bas', 'Form1.frm'): (backup / name).write_bytes((root / name).read_bytes())
form = (backup / 'Form1.frm').read_bytes()
for old, new in [(b'AddResultLine "BA (HCA baseline)"', b'AddResultLine "BA"'),
                 (b'AddResultLine "BA HCA baseline - Trial "', b'AddResultLine "BA - Trial "')]:
    assert form.count(old) == 1, old
    form = form.replace(old, new)
(root / 'modBA.bas').write_bytes(candidate.read_bytes())
(root / 'Form1.frm').write_bytes(form)
production = out / 'production'
production.mkdir()
project = (root / 'RC_RT_HCA_v2.vbp').read_bytes().decode('ascii')
for line in project.splitlines():
    if not line.startswith(('Module=', 'Form=')): continue
    name = line.split(';')[-1].strip() if line.startswith('Module=') else line[5:]
    project = project.replace(line, line.replace(name, str(root / name)))
assert 'Startup="Form1"' in project and 'Probe' not in project
project += '\r\nExeName32="RC_RT_HCA_v2.exe"\r\n'
(production / 'Production.vbp').write_bytes(project.encode('ascii'))
decision = {
    'authorization': 'User explicitly authorized applying the code if the H5/fc320 twenty-per-method comparison again supports overall advantage.',
    'observed_reference': summary['observed_target_cost'],
    'BA_attainment': ba['target_attainment'],
    'HCA_attainment': hca['target_attainment'],
    'complete_sample_effort_reduction_percent': summary['relative_effort_reduction_percent'],
    'SD_capped_effort_BA': ba['sd_effort_all_runs'],
    'SD_capped_effort_HCA': hca['sd_effort_all_runs'],
    'SD_caveat': 'BA capped-effort SD is higher; disclosed. User treats SD as supporting evidence, not a mandatory lower-than-HCA gate.',
    'SD_final_price_BA': ba['sd_valid_final_cost'],
    'SD_final_price_HCA': hca['sd_valid_final_cost'],
    'changed_root_files': ['modBA.bas', 'Form1.frm'],
    'source_hashes_after': {name: sha(root / name) for name in before},
    'compile_project': str(production / 'Production.vbp'),
    'root_executable_replaced': False,
    'new_performance_trials_after_install': 0,
}
(out / 'installation-decision.json').write_text(json.dumps(decision, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print('Installed exact tested BA; updated two stale UI labels. Ready to compile actual root sources.')
