"""Check the native source supplement against independent arithmetic and force audit."""
from pathlib import Path
import csv, json, math
P = Path(__file__).resolve().parent
rows = list(csv.DictReader((P/'aci-supplement-native.csv').open()))
actual = {row['item']:float(row['value']) for row in rows}
psi_ksc = 0.45359237 / (2.54**2)
expected = dict(vc_aci99=1.1*math.sqrt(320/psi_ksc)*psi_ksc,
                vc_textbook=.29*math.sqrt(320),fs_grade40_50=20000*psi_ksc,
                fs_grade60_plus=24000*psi_ksc,fy_sd40_psi=4000/psi_ksc,
                stem_horizontal_min_gross=.0025*100*40,
                production_criteria_ready=0,NATIVE_COMPLETE=1)
prior = json.loads((P/'recalculation-details.json').read_text())
assert len(prior['cases']) == 1
for i, member in enumerate(prior['cases'][0]['members']):
    expected[f'member_{i}_shear'] = member['nominal_v']
assert actual.keys() == expected.keys()
for key,value in expected.items():
    assert abs(actual[key]-value) < 1e-7, (key,actual[key],value)
print('Native supplementary ACI arithmetic: 11 values agree with independent units and existing independent force reconciliation.')
print('Production acceptance remains disabled; no ACI-compliant design or new optimum is asserted.')
