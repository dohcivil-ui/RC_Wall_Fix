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
    elif line.startswith('ACCEPT_SNAPSHOT='):
        records[-1]['accept_snapshot'] = Path(line.split('=', 1)[1])
    elif line.startswith('LOOP_SNAPSHOT='):
        records[-1]['loop_snapshot'] = Path(line.split('=', 1)[1])
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
    summary = rows(record['summary'])
    exported = rows(record['accept_snapshot'])
    trace = []
    for trial in summary:
        run_folder = Path(trial['run_folder'])
        native_rows = rows(run_folder/'evaluations.csv')
        per_trial = rows(run_folder/record['accept'].name)
        assert len(per_trial) == len(native_rows) > 0
        trace.extend((trial, ev) for ev in native_rows)
    assert len(exported) == len(trace) > 0
    best = float('inf')
    previous_trial = None
    for out, (trial, native) in zip(exported, trace):
        assert out['Trial'] == trial['trial'] and out['Seed'] == trial['seed']
        if out['Trial'] != previous_trial:
            best = float('inf')
            previous_trial = out['Trial']
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
    loop = rows(record['loop_snapshot'])
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
    for kind in ('accept', 'loop'):
        archived = list((root/'archive').glob(record[kind].stem+'-*.csv'))
        assert any(p.read_bytes() == record[kind+'_snapshot'].read_bytes() for p in archived), kind


assert len(records) == 4 and set(trial_counts) == {1, 2}
assert all(categories[k] > 0 for k in ('Rejected', 'Passed', 'Passed and Better value'))
result = {'sessions': len(records), 'accept_categories': dict(categories), 'loop_rows_per_session': trial_counts,
          'outcome': 'All export rows match actual native evaluation/trial logs; no 30-trial run.',
          'files': [{k: str(v) for k,v in r.items()} for r in records]}
(evidence/'primary-csv-verification.json').write_text(json.dumps(result, indent=2)+'\n')
print(json.dumps({k:v for k,v in result.items() if k != 'files'}))
