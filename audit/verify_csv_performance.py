"""Verify the native VB6 CSV buffering experiment; does not run research trials."""
import csv
import json
import re
import sys
from pathlib import Path

folder = Path(sys.argv[1])


def sessions(path):
    result = []
    for line in path.read_text().splitlines():
        match = re.fullmatch(r'method=(\w+); evaluations=(\d+); elapsed=(.+)', line)
        if match:
            result.append(dict(method=match[1], budget=int(match[2]), seconds=float(match[3])))
        elif line.startswith('run='):
            result[-1]['folder'] = Path(line[4:])
    return result


before = sessions(folder / 'timing-before-array.txt')
after = sessions(folder / 'production' / 'timing.txt')
assert len(before) == 2 and len(after) == 4
for old, new in zip(before, after[:2]):
    assert old['method'] == new['method'] and old['budget'] == new['budget'] == 1024
    for name in ('evaluations.csv', f"accept-{old['method']}-H5.csv"):
        assert (old['folder'] / name).read_bytes() == (new['folder'] / name).read_bytes(), name
    # Detect the measured serialization regression with a wide timing margin.
    assert new['seconds'] < old['seconds'] * 0.6, (old, new)

for run in after:
    with (run['folder'] / 'evaluations.csv').open(newline='') as stream:
        rows = list(csv.DictReader(stream))
    assert len(rows) == run['budget']
    best = 999999999.0
    for number, row in enumerate(rows, 1):
        assert int(row['evaluation']) == number
        if row['valid'] == 'True':
            best = min(best, float(row['cost']))
        assert abs(float(row['best_cost']) - best) < 1e-7
    with (run['folder'] / f"accept-{run['method']}-H5.csv").open(newline='') as stream:
        accept = list(csv.DictReader(stream))
    assert len(accept) == run['budget']
    assert list(accept[0]) == ['No.', 'Rejected', 'Passed', 'Passed and Better value']
    assert [int(row['No.']) for row in accept] == list(range(run['budget']))

result = {'before': before, 'after': after, 'same_1024_evaluation_and_accept_bytes': True,
          'budgets_verified': [run['budget'] for run in after],
          'scope': 'Isolated compiled VB6, H5/fc320/seed12345. No 30-trial batch; IDE timing not measured.'}
(folder / 'verified-performance.json').write_text(json.dumps(result, default=str, indent=2) + '\n')
print('Native CSV performance verified: identical 1024-row outputs; BA/HCA 5000-row budgets preserved.')
