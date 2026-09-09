from pathlib import Path
import hashlib
import json
import subprocess

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'audit/github-checkpoint-2026-09-09'
OUT.mkdir(exist_ok=False)
selected = set()

def add(p):
    p = Path(p)
    if not p.is_absolute():
        p = ROOT / p
    assert p.is_file(), p
    rel = p.relative_to(ROOT).as_posix()
    assert not rel.startswith('result_csv/')
    assert p.suffix.lower() not in ('.exe', '.pdf', '.png', '.jpg', '.zip', '.vbw', '.pyc')
    assert p.stat().st_size < 2_000_000, p
    selected.add(rel)

for p in ('Form1.frm', 'modBA.bas', 'modBatch.bas', 'modHillClimbing.bas', 'modProjectChecks.bas', 'modShared.bas', 'HANDOFF.md', 'NEXT_CHAT_PROMPT.md', 'audit/handoff-history-2026-09-09.md'):
    add(p)

for folder in ('ba-geometry-hold', 'ba-geometry-hold-h4-fc280-20', 'ba-geometry-hold-h5-fc320-20', 'fresh-random'):
    for p in (ROOT / 'audit' / folder).iterdir():
        if p.is_file() and p.suffix.lower() in ('.md', '.json', '.bas', '.diff'):
            add(p)

for folder in ('ba-geometry-hold-installation', 'sample-csv-export', 'no-trial-folders', 'center-window'):
    for p in (ROOT / 'audit' / folder).rglob('*'):
        if p.is_file() and p.suffix.lower() in ('.md', '.json', '.bas', '.frm', '.frx', '.vbp', '.py', '.txt', '.log', '.csv'):
            add(p)

for folder in ('ten-trials-h3/geometry-hold-10', 'ten-trials-h3/geometry-hold-replication-10', 'parameter-trials/h4-fc280-20', 'parameter-trials/h5-fc320-20'):
    for p in (ROOT / 'audit' / folder).iterdir():
        if p.is_file() and p.name not in ('BA-trace.csv', 'HCA-trace.csv') and p.suffix.lower() in ('.csv', '.json', '.txt', '.log', '.bas', '.frm', '.frx', '.vbp'):
            add(p)

for p in (
    'audit/manuscript-edit22-review/REVIEW.md',
    'audit/manuscript-edit22-review/review-verification.json',
    'audit/build_geometry_hold_experiment.py', 'audit/geometry_hold_probe.bas',
    'audit/install_geometry_hold.py',
    'audit/h4_280_twenty_probe.bas', 'audit/h5_320_twenty_probe.bas',
    'audit/match_sample_csv_export.py', 'audit/prepare_sample_csv_probe.py',
    'audit/sample_csv_probe.bas', 'audit/verify_sample_csv_export.py',
    'audit/prepare_checkpoint_handoff.py', 'audit/prepare_checkpoint_manifest.py',
):
    add(p)

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

protected = {str(p.relative_to(ROOT / 'result_csv')): sha(p) for p in (ROOT / 'result_csv').rglob('*') if p.is_file()}
(OUT / 'protected-result-csv-before.json').write_text(json.dumps(protected, indent=2), encoding='utf-8')
root_sources = {p.name: sha(p) for p in ROOT.iterdir() if p.suffix.lower() in ('.bas', '.frm', '.frx', '.vbp', '.exe')}
(OUT / 'source-before-git.json').write_text(json.dumps(root_sources, indent=2), encoding='utf-8')

readme = '''# Git checkpoint: version revise BA pass with resersh

User authorized commit, pull and push to the existing GitHub repository on 2026-09-09.
This checkpoint includes current production sources, current handoff and next-chat prompt,
the complete prior handoff, and compact evidence for the installed H3/H4/H5 candidate and
the latest four-file CSV export contract. No optimizer run is part of this checkpoint.

Protected result_csv changes are intentionally NOT staged: user-deleted tracked results
remain deleted locally and existing archive files remain untouched. The checkpoint does
not add, delete, restore or modify those research results in Git. Compiled binaries, large
raw evaluation traces, manuscript renderings and older experimental artifacts remain local.
The trial summary files retained here include all outcomes in the accepted candidate's
reported H3/H4/H5 samples; there is no new 30-trial dataset or p-value.

Source integrity was checked against audit/sample-csv-export/verification.json and the
tested candidate before commit. Publication status and commit hash must be checked using
git log and the remote; this file does not assert that a future push has succeeded.
'''
(OUT / 'README.md').write_text(readme, encoding='utf-8')
for name in ('README.md', 'protected-result-csv-before.json', 'source-before-git.json'):
    add(OUT / name)

records = [{'path': name, 'bytes': (ROOT / name).stat().st_size, 'sha256': sha(ROOT / name)} for name in sorted(selected)]
manifest = OUT / 'commit-files.json'
manifest.write_text(json.dumps(records, indent=2), encoding='utf-8')
paths = sorted(selected | {manifest.relative_to(ROOT).as_posix()})
(OUT / 'stage-paths.txt').write_text('\n'.join(paths) + '\n', encoding='utf-8')
print(json.dumps({'files': len(paths), 'bytes': sum(r['bytes'] for r in records), 'protected_result_files': len(protected), 'manifest': str(manifest)}, indent=2))
