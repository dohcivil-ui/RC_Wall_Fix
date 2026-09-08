# Actual VB6 H5 recalculation: independent stem/toe/heel reinforcement
Other inputs read from Form1.frm defaults; H=5 and fc=320 retained study. See full-sizing-inputs.json. One trial per passive case.
Main case eta=1: full Rankine passive using H1. Eta=0 is the comparison. Heights alone do not establish field mobilization.
Production BA requires all criteria. Separate relaxed grid: stability + supported bending + legacy stress only; NOT an accepted EIT design.
Missing minimum/shear/detailing are NOT silently passed. No synthetic limits. Concrete unit price=2617; steel=24
Cost model includes concrete and listed main bars, with original +0.4 m bar length allowances; not a complete construction BOQ.

## Passive fraction=0
Actual production BA: NO_SOLUTION; evaluations=5000; best evaluation=0; folder=C:\reserch 69\RC_Wall_Fix\audit\results\BA-25690908-211730-seed12345
Relaxed grid: rows=520200; all 3 stability=292021; screened=288781; steel checks=17391660
Lowest relaxed grid estimate=9144.0755000000 baht/m; complete validator=WSD_CRITERIA_UNVERIFIED
AUDIT RESULT: INDETERMINATE_WSD
PER-CHECK AUDIT; no EIT compliance claim
H=5; H1=1.2; gamma_soil=1.8; gamma_concrete=2.4; phi=30; mu=0.6; qa=30
tt=0.2; tb=0.45; TBase=0.3; Base=3; LToe=1.1; LHeel=1.45
fc_prime=320; fy=4000; cover=0.075; passive fraction=0

| Check | Actual | Limit / rule | Comparison | Basis / verification |
| --- | --- | --- | --- | --- |
| Stem height | 4.7000 m | H-TBase | CALCULATED | H measured from base underside |
| Stem bars | DB25 @ 0.25 m | specified candidate | INPUT | Design |
| Stem effective depth | 0.3625 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Stem As | 19.6350 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Stem signed M | 10.3823 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Stem necessary yield bound | 10.3823 tf.m/m | <= 28.4707 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Stem concrete stress | 64.9034 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Stem steel stress | 1601.3334 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Stem governing shear height | 0.0000 m | above base top | CALCULATED | Maximum nominal abs(V)/(b*d) along tapered stem; EIT critical-section rule UNVERIFIED |
| Stem nominal shear V/bd | 1.8281 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Stem minimum steel | 19.6350 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel length | 1.4500 m | >= 0.3000 | WITHIN_LISTED_LIMIT | Existing project geometry rule |
| Overturning FS | 2.9782  | >= 2.0000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Sliding FS | 1.5931  | >= 1.5000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Full contact abs(e) | 0.2582 m | <= 0.5000 | WITHIN_LISTED_LIMIT | Linear full-compression model; signed e=0.2582 positive to toe |
| q_toe / q_heel | 10.0664 / 3.2095 tf/m2 | signed pressure diagram | CALCULATED | BearingEdges |
| Bearing qa/qmax | 2.9802  | >= 1.0000 | WITHIN_LISTED_LIMIT | qa INPUT interpreted as ALLOWABLE |
| Toe bars | DB20 @ 0.25 m | specified candidate | INPUT | Design |
| Toe effective depth | 0.2150 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Toe As | 12.5664 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Toe signed M | 4.1674 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Toe necessary yield bound | 4.1674 tf.m/m | <= 10.8071 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Toe concrete stress | 71.9514 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Toe steel stress | 1698.7603 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Toe nominal shear V/bd | 2.7641 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Toe minimum steel | 12.5664 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel bars | DB20 @ 0.20 m | specified candidate | INPUT | Design |
| Heel effective depth | 0.2150 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Heel As | 15.7080 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Heel signed M | 5.1152 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Heel necessary yield bound | 5.1152 tf.m/m | <= 13.5088 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Heel concrete stress | 81.2953 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Heel steel stress | 1684.6506 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Heel nominal shear V/bd | 2.6189 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Heel minimum steel | 15.7080 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Complete EIT 011007-19 compliance | Not established | Verified applicable clauses and complete detailing | UNVERIFIED | See audit/WSD_PRIMARY_SOURCE_SEARCH.md |
| Cover, bar order, spacing, anchorage, distribution steel | Not fully checked | Applicable detailing provisions | UNVERIFIED | Current d assumes main bar outermost; no interposed transverse bar |


## Passive fraction=1
Actual production BA: NO_SOLUTION; evaluations=5000; best evaluation=0; folder=C:\reserch 69\RC_Wall_Fix\audit\results\BA-25690908-211841-seed12345
Relaxed grid: rows=520200; all 3 stability=309671; screened=306427; steel checks=18450500
Lowest relaxed grid estimate=8273.1436800000 baht/m; complete validator=WSD_CRITERIA_UNVERIFIED
AUDIT RESULT: INDETERMINATE_WSD
PER-CHECK AUDIT; no EIT compliance claim
H=5; H1=1.2; gamma_soil=1.8; gamma_concrete=2.4; phi=30; mu=0.6; qa=30
tt=0.2; tb=0.4; TBase=0.3; Base=2.5; LToe=1; LHeel=1.1
fc_prime=320; fy=4000; cover=0.075; passive fraction=1

| Check | Actual | Limit / rule | Comparison | Basis / verification |
| --- | --- | --- | --- | --- |
| Stem height | 4.7000 m | H-TBase | CALCULATED | H measured from base underside |
| Stem bars | DB16 @ 0.10 m | specified candidate | INPUT | Design |
| Stem effective depth | 0.3170 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Stem As | 20.1062 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Stem signed M | 9.7262 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Stem necessary yield bound | 9.7262 tf.m/m | <= 25.4947 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Stem concrete stress | 74.9128 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Stem steel stress | 1686.5534 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Stem governing shear height | 0.5711 m | above base top | CALCULATED | Maximum nominal abs(V)/(b*d) along tapered stem; EIT critical-section rule UNVERIFIED |
| Stem nominal shear V/bd | 1.6475 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Stem minimum steel | 20.1062 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel length | 1.1000 m | >= 0.3000 | WITHIN_LISTED_LIMIT | Existing project geometry rule |
| Overturning FS | 2.1604  | >= 2.0000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Sliding FS | 1.8097  | >= 1.5000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Full contact abs(e) | 0.3514 m | <= 0.4167 | WITHIN_LISTED_LIMIT | Linear full-compression model; signed e=0.3514 positive to toe |
| q_toe / q_heel | 11.9013 / 1.0115 tf/m2 | signed pressure diagram | CALCULATED | BearingEdges |
| Bearing qa/qmax | 2.5207  | >= 1.0000 | WITHIN_LISTED_LIMIT | qa INPUT interpreted as ALLOWABLE |
| Toe bars | DB20 @ 0.25 m | specified candidate | INPUT | Design |
| Toe effective depth | 0.2150 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Toe As | 12.5664 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Toe signed M | 4.0547 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Toe necessary yield bound | 4.0547 tf.m/m | <= 10.8071 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Toe concrete stress | 70.0046 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Toe steel stress | 1652.7955 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Toe nominal shear V/bd | 2.8667 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Toe minimum steel | 12.5664 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel bars | DB20 @ 0.25 m | specified candidate | INPUT | Design |
| Heel effective depth | 0.2150 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Heel As | 12.5664 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Heel signed M | 3.9756 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Heel necessary yield bound | 3.9756 tf.m/m | <= 10.8071 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Heel concrete stress | 68.6403 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Heel steel stress | 1620.5851 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | User-selected project n=9; EIT clause UNVERIFIED |
| Heel nominal shear V/bd | 2.5690 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Heel minimum steel | 12.5664 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Complete EIT 011007-19 compliance | Not established | Verified applicable clauses and complete detailing | UNVERIFIED | See audit/WSD_PRIMARY_SOURCE_SEARCH.md |
| Cover, bar order, spacing, anchorage, distribution steel | Not fully checked | Applicable detailing provisions | UNVERIFIED | Current d assumes main bar outermost; no interposed transverse bar |

NATIVE COMPLETE: two grids completed; production accepted designs=0; WSDReviewed=False
