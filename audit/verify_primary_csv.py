"""Compare user-facing accept/loopPrice exports to native evaluation and trial logs."""
import csv
import json
import sys
from collections import Counter
from pathlib import Path

evidence = Path(sys.argv[1])
records = []
for line in (evidence/'native-csv-paths.txt').read_text().splitlines():
    if line.startswith('ACCEPT='):
        records.append({'accept': Path(line.split('=', 1)[1])})
    elif line.startswith('LOOP='):
        records[-1]['loop'] = Path(line.split('=', 1)[1])
    elif line.startswith('EVALUATIONS='):
        records[-1]['evaluations'] = Path(line.split('=', 1)[1])
    elif ' summary: ' in line:
        records[-1]['summary'] = Path(line.split(' summary: ', 1)[1])

def rows(path):
    with path.open(newline='') as f:
        return list(csv.DictReader(f))

categories = Counter()
trial_counts = []
for record in records:
    exported, trace = rows(record['accept']), rows(record['evaluations'])
    assert len(exported) == len(trace) > 0
    best = float('inf')
    for out, native in zip(exported, trace):
        assert int(out['No.']) == int(native['evaluation'])
        if native['valid'] != 'True':
            assert out['Rejected'] == 'INVALID'
            assert out['Passed'] == out['Passed and Better value'] == ''
            categories['Rejected'] += 1
        else:
            cost = float(native['cost'])
            column = 'Passed and Better value' if cost < best else 'Passed'
            assert abs(float(out[column]) - cost) < 1e-7
            assert out['Rejected'] == ''
            other = 'Passed' if column == 'Passed and Better value' else 'Passed and Better value'
            assert out[other] == ''
            categories[column] += 1
            best = min(best, cost)
    loop, summary = rows(record['loop']), rows(record['summary'])
    assert len(loop) == len(summary) > 0
    trial_counts.append(len(loop))
    for out, native in zip(loop, summary):
        assert out['No.'] == native['trial']
        assert out['Loop'] == native['best_evaluation']
        assert out['Status'] == native['status']
        if native['best_cost']:
            assert abs(float(out['BestPrice']) - float(native['best_cost'])) < 1e-7
        else:
            assert out['BestPrice'] == '' and out['Loop'] == '0'
    root = record['evaluations'].parent.parent
    assert record['accept'].parent == root and record['loop'].parent == root

assert len(records) == 4 and set(trial_counts) == {1, 2}
assert all(categories[k] > 0 for k in ('Rejected', 'Passed', 'Passed and Better value'))
result = {'sessions': len(records), 'accept_categories': dict(categories), 'loop_rows_per_session': trial_counts,
          'outcome': 'All export rows match actual native evaluation/trial logs; no 30-trial run.',
          'files': [{k: str(v) for k,v in r.items()} for r in records]}
(evidence/'primary-csv-verification.json').write_text(json.dumps(result, indent=2)+'\n')
print(json.dumps({k:v for k,v in result.items() if k != 'files'}))
