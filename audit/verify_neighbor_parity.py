"""Validate recorded native generator parity and search traces for the aligned HCA."""
import csv
import json
import re
import sys
from pathlib import Path

folder = Path(sys.argv[1])
before = (folder/'before/parity.txt').read_text()
final = (folder/'final/parity.txt').read_text()
assert 'GENERATOR checks=5400; failures=530' in before
assert 'GENERATOR checks=5400; failures=0' in final
assert 'FATAL' not in final and 'FAIL:' not in final
paths = [Path(line[4:]) for line in final.splitlines() if line.startswith('RUN=')]
assert len(paths) == 8
traces = []
results = []
for path in paths:
    with (path/'evaluations.csv').open(newline='') as stream:
        trace = list(csv.DictReader(stream))
    report = (path/'run.txt').read_text()
    budget = int(re.search(r'Budget=(\d+)', report)[1])
    method = re.search(r'Algorithm=(\w+);', report)[1]
    height = re.search(r'H=(\d+);', report)[1]
    assert len(trace) == budget
    if method == 'HCA':
        assert 'SearchPolicy=HCA_BA_NEIGHBOR_V1' in report
        assert trace[0]['entry'] == 'initial'
        assert all(row['entry'] == 'neighbor' for row in trace[1:])
    best = 999999999.0
    first_best = 0
    with (path/f'accept-{method}-H{height}.csv').open(newline='') as stream:
        accepts = list(csv.DictReader(stream))
    assert len(accepts) == budget
    assert list(accepts[0]) == ['No.', 'Rejected', 'Passed', 'Passed and Better value']
    for number, (row, accept) in enumerate(zip(trace, accepts), 1):
        assert int(row['evaluation']) == number and int(accept['No.']) == number-1
        if row['valid'] == 'True' and float(row['cost']) < best:
            best = float(row['cost'])
            first_best = number
        assert abs(float(row['best_cost'])-best) < 1e-7
    assert int(re.search(r'BestEvaluation=(\d+);', report)[1]) == first_best
    traces.append(trace)
    results.append({'method':method, 'H':int(height), 'budget':budget,
                    'best':best if first_best else None, 'folder':str(path)})
for i in (0,2,4):
    assert traces[i][:21] == traces[i+1][:21], 'Search differs before first BA range update'
assert results[-1]['best'] is None
(folder/'verification.json').write_text(json.dumps({'generator_checks':5400, 'failures':0,
    'initial_21_evaluations_equal_at_H':[3,4,5], 'runs':results,
    'scope':'Native VB6; one H5 5000-evaluation run, short H3/H4/H5 and no-solution runs; no 30-trial batch.'},indent=2)+'\n')
print('Native parity verified: 5400 comparisons, identical first 21 BA/HCA evaluations at H3/H4/H5, exact budgets and no steel sweep.')
