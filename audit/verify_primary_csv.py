"""Compare user-facing accept/loopPrice exports to native evaluation and trial logs."""
import csv
import json
import re
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
    assert list(exported[0]) == ['No.', 'Rejected', 'Passed', 'Passed and Better value']
    # The primary accept file is exactly the final trial; other trials stay in their folders.
    for trial in summary:
        run_folder = Path(trial['run_folder'])
        native_rows = rows(run_folder/'evaluations.csv')
        per_trial = rows(run_folder/record['accept'].name)
        assert len(per_trial) == len(native_rows) > 0
    trace = rows(record['evaluations'])
    assert exported == per_trial and len(exported) == len(trace) > 0
    best = float('inf')
    db = {100:12, 101:16, 102:20, 103:25, 104:28}
    sp = {110:0.10, 111:0.15, 112:0.20, 113:0.25}
    for out, native in zip(exported, trace):
        assert int(out['No.']) == int(native['evaluation']) - 1
        if native['valid'] != 'True':
            # Independent quantity takeoff at H5/fc320/cover75mm; no strength acceptance.
            tt, tb, base, width, toe = (float(native[k]) for k in ('tt','tb','TBase','Base','LToe'))
            heel = width-toe-tb
            lengths = [5-base-0.075, toe-0.075, heel-0.075]
            volume = (tt+tb)*(5-base)/2 + width*base
            mass = sum(0.00617*db[int(native[part+'DB'])]**2/sp[int(native[part+'SP'])]*length
                       for part,length in zip(('Stem','Toe','Heel'),lengths))
            price = volume*2617 + mass*24
            assert min(lengths)>0
            assert abs(float(out['Rejected'])-price) <= 0.00500001
            assert out['Passed'] == out['Passed and Better value'] == ''
            categories['Rejected'] += 1
        else:
            cost = float(native['cost'])
            column = 'Passed and Better value' if cost < best else 'Passed'
            assert abs(float(out[column]) - cost) <= 0.00500001
            assert out['Rejected'] == ''
            other = 'Passed' if column == 'Passed and Better value' else 'Passed and Better value'
            assert out[other] == ''
            categories[column] += 1
            best = min(best, cost)
        populated = [out[k] for k in ('Rejected','Passed','Passed and Better value') if out[k]]
        assert len(populated)==1 and all(re.fullmatch(r'\d+\.\d{2}',v) for v in populated)
    loop = rows(record['loop_snapshot'])
    assert list(loop[0]) == ['No.', 'Loop', 'BestPrice']
    assert len(loop) == len(summary) > 0
    trial_counts.append(len(loop))
    for out, native in zip(loop, summary):
        assert out['No.'] == native['trial']
        if native['best_cost']:
            assert int(out['Loop']) == int(native['best_evaluation'])-1
            assert abs(float(out['BestPrice']) - float(native['best_cost'])) <= 0.00500001
            assert re.fullmatch(r'\d+\.\d{2}',out['BestPrice'])
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
          'outcome': 'Reference headers, last-trial accept, zero-based loops, two decimals and independent rejected quantities verified; no 30-trial run.',
          'files': [{k: str(v) for k,v in r.items()} for r in records]}
(evidence/'primary-csv-verification.json').write_text(json.dumps(result, indent=2)+'\n')
print(json.dumps({k:v for k,v in result.items() if k != 'files'}))
