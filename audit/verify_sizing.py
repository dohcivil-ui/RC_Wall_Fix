"""Independent numerical verification of the native sizing output, not a code check."""
from pathlib import Path
import csv
import math
import independent_checks as independent

P = Path(__file__).resolve().parent
rows = list(csv.DictReader((P / 'sizing-options.csv').open()))
assert len(rows) == 120
comparisons = 0
qualifying = 0
for row in rows:
    f = {k: float(row[k]) for k in ('tt', 'tb', 'TBase', 'Base', 'toe', 'DB', 'spacing')}
    ref = independent.reference(5, f['tb'], f['TBase'], f['Base'], f['toe'], tt=f['tt'], db=f['DB'], sp=f['spacing'])
    for column, expected in [('stem_height', ref['hs']), ('d', ref['depth']), ('Mstem', ref['net']), ('Mtoe', ref['mt']), ('Mheel', ref['mh']), ('FSot', ref['fsot']), ('FSsl', ref['fssl']), ('FSbc', 20 / max(ref['qt'], ref['qh']))]:
        assert abs(float(row[column]) - expected) < 1e-8, (row['option'], column)
        comparisons += 1
    area = math.pi * (f['DB'] / 10) ** 2 / 4 / f['spacing']
    dc = ref['depth'] * 100
    neutral = (-9 * area + math.sqrt((9 * area) ** 2 + 200 * 9 * area * dc)) / 100
    lever = dc - neutral / 3
    stresses = [(abs(m) * 1e5 / (50 * neutral * lever), abs(m) * 1e5 / (area * lever)) for m in (ref['net'], ref['mt'], ref['mh'])]
    ok = (ref['fsot'] >= 2 and ref['fssl'] >= 1.5 and min(ref['qt'], ref['qh']) >= 0 and max(ref['qt'], ref['qh']) <= 20 and min(ref['net'], ref['mt'], ref['mh']) >= 0 and all(c <= 144 and s <= 1700 for c, s in stresses))
    assert (row['known_checks_only'] == 'True') == ok
    assert row['complete_WSD'] == 'UNVERIFIED'
    qualifying += ok
text = (P / 'sizing-vb6-report.md').read_text()
assert 'FATAL' not in text
assert f'layouts=120; within known checks={qualifying}; accepted by full validator=0; WSDReviewed=False' in text
assert qualifying == 57
result = f'Independent sizing verification: {comparisons} numeric comparisons and 120 screening classifications agree.\n57 within known checks; 0 accepted as a complete WSD design.\n'
(P / 'sizing-independent-checks.txt').write_text(result, encoding='ascii')
print(result)
for option in range(1, 7):
    choices = [r for r in rows if int(r['option']) == option and r['known_checks_only'] == 'True']
    best = min(choices, key=lambda r: float(r['DB']) ** 2 / float(r['spacing']))
    print(f"Option {option}: tb/TBase={best['tb']}, B={best['Base']}, least-area common layout in this limited comparison DB{best['DB']}@{best['spacing']}; UNVERIFIED")
