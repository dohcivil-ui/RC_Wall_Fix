"""Check native HCA evidence against the user's existing trials without rerunning 30 trials."""
import csv
import json
import re
import sys
from pathlib import Path

folder = Path(sys.argv[1])
analysis = json.loads((folder / 'user-trace-analysis.json').read_text())
old = {item['method']: {int(t['seed']): t for t in item['trials']} for item in analysis}


def rows(path):
    with path.open(newline='') as stream:
        return list(csv.DictReader(stream))


log = (folder / 'final' / 'probe.txt').read_text()
assert 'FATAL' not in log and 'FAIL:' not in log
assert 'HCA_BA_SEED valid=True; cost=8540.19588' in log
results = []
for line in log.splitlines():
    if '; folder=' not in line:
        continue
    run = Path(line.split('; folder=')[1])
    trace = rows(run / 'evaluations.csv')
    report = (run / 'run.txt').read_text()
    method = re.search(r'Algorithm=(\w+);', report)[1]
    budget = int(re.search(r'Budget=(\d+)', report)[1])
    seed = int(re.search(r'Seed=(\d+);', report)[1])
    assert len(trace) == budget
    best = 999999999.0
    best_eval = 0
    accepts = rows(run / f'accept-{method}-H{re.search(r"H=(\d+);", report)[1]}.csv')
    assert len(accepts) == budget
    assert list(accepts[0]) == ['No.', 'Rejected', 'Passed', 'Passed and Better value']
    for number, (row, accept) in enumerate(zip(trace, accepts), 1):
        assert int(row['evaluation']) == number and int(accept['No.']) == number - 1
        if row['valid'] == 'True' and float(row['cost']) < best:
            best = float(row['cost'])
            best_eval = number
            column = 'Passed and Better value'
        else:
            column = 'Passed' if row['valid'] == 'True' else 'Rejected'
        assert abs(float(row['best_cost']) - best) < 1e-7
        if row['valid'] == 'True':
            assert abs(float(accept[column]) - float(row['cost'])) <= 0.00500001
        assert all(accept[c] == '' for c in ['Rejected', 'Passed', 'Passed and Better value'] if c != column)
    assert int(re.search(r'BestEvaluation=(\d+);', report)[1]) == best_eval
    if budget == 5000:
        previous = rows(Path(old[method][seed]['folder']) / 'evaluations.csv')
        if method == 'BA':
            assert trace == previous, 'BA search changed'
        else:
            assert trace[:4940] == previous[:4940], 'Original random prefix changed'
            sweep = trace[4940:]
            assert all(row['entry'] == 'steel_refine' for row in sweep)
            geometry = ['tt', 'tb', 'TBase', 'Base', 'LToe']
            assert len({tuple(row[k] for k in geometry) for row in sweep}) == 1
            for part, chunk in zip(['Stem', 'Toe', 'Heel'], [sweep[:20], sweep[20:40], sweep[40:]]):
                assert {(int(r[part + 'DB']), int(r[part + 'SP'])) for r in chunk} == {
                    (db, sp) for db in range(100, 105) for sp in range(110, 114)}
            assert best <= float(trace[4939]['best_cost']), 'Refinement lost a previous best'
    results.append({'label': line.split('; folder=')[0], 'budget': budget, 'seed': seed,
                    'best_cost': best if best_eval else None, 'run': str(run)})

local = next(r for r in results if r['label'].startswith('HCA_LOCAL125'))
assert local['best_cost'] < 8573.88408, 'Known cheaper feasible main-bar pair missed'
new_best = next(r for r in results if r['label'].startswith('HCA5000 seed12365'))
assert abs(new_best['best_cost'] - 8538.41892) < 1e-7
assert any(r['best_cost'] is None for r in results), 'Missing no-solution case'
result = {'checked_native_runs': len(results), 'results': results,
          'scope': 'Three H5 HCA seeds, bounded H3/H4/local/no-solution tests, and BA control. No 30-trial rerun.'}
(folder / 'verification.json').write_text(json.dumps(result, indent=2) + '\n')
print('Native evidence verified: original HCA 4940-evaluation prefixes, 60 counted steel candidates, global best, CSV schema and unchanged BA control.')
