"""Compare actual VB6 report numbers with independent integration/neutral-axis math.
This verifies numerical agreement only, not EIT requirements or VB6 execution.
"""
from pathlib import Path
import math
import independent_checks as independent
from independent_stem_profile import stem_profile

P = Path(__file__).resolve().parent
n = 2040000 / (15100 * math.sqrt(320))  # Reference material model; retain precision.
reports = (P / 'h5-vb6-checks.md').read_text().split('## ')[1:]
assert len(reports) == 3
cases = [
    independent.reference(5, .2, .3, 3.5, .5, db=12, sp=.25),
    independent.reference(5, .2, .3, 3.5, .5, p=1, db=12, sp=.25),
    independent.reference(5, .45, .45, 3, .6, tt=.25, db=25, sp=.20),
]
count = 0
messages = []


def near(actual, expected, label):
    global count
    assert abs(actual - expected) < .000051, (label, actual, expected)
    count += 1


for case_index, (text, case) in enumerate(zip(reports, cases)):
    rows = {}
    for line in text.splitlines():
        if line.startswith('| '):
            columns = [c.strip() for c in line.split('|')[1:-1]]
            rows[columns[0]] = columns[1:]

    def actual(label):
        return float(rows[label][0].split()[0])

    near(actual('Stem height'), case['hs'], 'height')
    bars = [('Stem', 12, .25)] if case_index < 2 else [('Stem', 25, .20), ('Toe', 25, .25), ('Heel', 20, .15)]
    for member, db, spacing in bars:
        depth = (case['tb'] if member == 'Stem' else case['tbase']) - .075 - db / 2000
        steel = math.pi * (db / 10) ** 2 / 4 / spacing
        moment = case[{'Stem': 'net', 'Toe': 'mt', 'Heel': 'mh'}[member]]
        # Solve compression/tension equilibrium in cm units, independently
        # from the production non-dimensional k/rho implementation.
        dc = depth * 100
        neutral = (-n * steel + math.sqrt((n * steel) ** 2 + 200 * n * steel * dc)) / 100
        lever = dc - neutral / 3
        steel_stress = abs(moment) * 1e5 / (steel * lever)
        concrete_stress = abs(moment) * 1e5 / (50 * neutral * lever)
        near(actual(member + ' effective depth'), depth, member + ' depth')
        near(actual(member + ' As'), steel, member + ' area')
        near(actual(member + ' signed M'), moment, member + ' M')
        near(actual(member + ' concrete stress'), concrete_stress, member + ' fc')
        near(actual(member + ' steel stress'), steel_stress, member + ' fs')
        near(float(rows[member + ' necessary yield bound'][1].split()[1]), steel * 4000 * depth / 1000, member + ' bound')
        if member == 'Stem':
            profile = stem_profile(case['h'],case['h1'],case['tb'],case['tt'],case['tbase'],.075,db,spacing,case['p'])
            force = profile['max_v'] * (10 * depth)
            near(actual('Stem governing shear height'), profile['shear_height'], 'stem shear height')
        else:
            q = lambda x: case['qt'] + (case['qh'] - case['qt']) * x / case['b']
            if member == 'Toe':
                force = abs(independent.integrate(lambda x: q(x) - 2.4 * case['tbase'] - 1.8 * (1.2 - case['tbase']), 0, case['toe'] - depth))
            else:
                force = abs(independent.integrate(lambda x: 2.4 * case['tbase'] + 1.8 * case['hs'] - q(x), case['toe'] + case['tb'] + depth, case['b']))
        near(actual(member + ' nominal shear V/bd'), force / (10 * depth), member + ' shear')
        assert rows[member + ' nominal shear V/bd'][1:3] == ['UNSET', 'UNVERIFIED']
        assert rows[member + ' minimum steel'][1:3] == ['UNSET', 'UNVERIFIED']
    if case_index == 2:
        near(actual('Overturning FS'), case['fsot'], 'OT')
        near(actual('Sliding FS'), case['fssl'], 'SL')
        near(actual('Full contact abs(e)'), abs(case['e']), 'e')
        near(actual('Bearing qa/qmax'), 20 / max(case['qt'], case['qh']), 'BC')
        near(float(rows['q_toe / q_heel'][0].split()[0]), case['qt'], 'q_toe')
        near(float(rows['q_toe / q_heel'][0].split()[2]), case['qh'], 'q_heel')
        assert 'AUDIT RESULT: INDETERMINATE_WSD' in text
    else:
        assert 'AUDIT RESULT: FAIL_NECESSARY_YIELD_BOUND' in text
    messages.append(f'Case {case_index + 1}: numerical agreement at four-decimal report precision; missing criteria remain UNVERIFIED.')

messages.append(f'Independent comparisons: {count}; mismatches=0. Actual VB6 execution evidence is separate in native-regression.txt.')
(P / 'report-independent-checks.txt').write_text('\n'.join(messages) + '\n', encoding='utf-8')
print('\n'.join(messages))
