# Actual VB6 main-steel sizing comparison
Six explicit geometry alternatives x 20 common main-bar layouts. Not a BA run or full WSD optimization.
Common bars are a practical comparison choice, not a restriction in production BA. No synthetic criteria.
Legacy flexural limits: n=9, fc=144, fs=1700 kgf/cm2. Code verification remains outstanding.

## Option 4: common DB20 @ 0.15 m
AUDIT RESULT: INDETERMINATE_WSD
PER-CHECK AUDIT; no EIT compliance claim
H=5; H1=1.2; gamma_soil=1.8; gamma_concrete=2.4; phi=30; mu=0.6; qa=20
tt=0.25; tb=0.45; TBase=0.45; Base=3; LToe=0.6; LHeel=1.95
fc_prime=320; fy=4000; cover=0.075; passive fraction=0

| Check | Actual | Limit / rule | Comparison | Basis / verification |
| --- | --- | --- | --- | --- |
| Stem height | 4.5500 m | H-TBase | CALCULATED | H measured from base underside |
| Stem bars | DB20 @ 0.15 m | specified candidate | INPUT | Design |
| Stem effective depth | 0.3650 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Stem As | 20.9440 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Stem signed M | 9.4196 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Stem necessary yield bound | 9.4196 tf.m/m | <= 30.5782 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Stem concrete stress | 56.8228 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Stem steel stress | 1355.9881 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Stem nominal shear V/bd | 1.7016 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Stem minimum steel | 20.9440 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel length | 1.9500 m | >= 0.3000 | WITHIN_LISTED_LIMIT | Existing project geometry rule |
| Overturning FS | 3.2626  | >= 2.0000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Sliding FS | 1.9092  | >= 1.5000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Full contact abs(e) | 0.3149 m | <= 0.5000 | WITHIN_LISTED_LIMIT | Linear full-compression model; signed e=0.3149 positive to toe |
| q_toe / q_heel | 12.9644 / 2.9454 tf/m2 | signed pressure diagram | CALCULATED | BearingEdges |
| Bearing qa/qmax | 1.5427  | >= 1.0000 | WITHIN_LISTED_LIMIT | qa INPUT interpreted as ALLOWABLE |
| Toe bars | DB20 @ 0.15 m | specified candidate | INPUT | Design |
| Toe effective depth | 0.3650 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Toe As | 20.9440 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Toe signed M | 1.7760 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Toe necessary yield bound | 1.7760 tf.m/m | <= 30.5782 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Toe concrete stress | 10.7133 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Toe steel stress | 255.6559 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Toe nominal shear V/bd | 0.6530 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Toe minimum steel | 20.9440 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel bars | DB20 @ 0.15 m | specified candidate | INPUT | Design |
| Heel effective depth | 0.3650 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Heel As | 20.9440 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Heel signed M | 7.8974 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Heel necessary yield bound | 7.8974 tf.m/m | <= 30.5782 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Heel concrete stress | 47.6400 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Heel steel stress | 1136.8554 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Heel nominal shear V/bd | 1.5971 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Heel minimum steel | 20.9440 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Complete EIT 011007-19 compliance | Not established | Verified applicable clauses and complete detailing | UNVERIFIED | See audit/WSD_PRIMARY_SOURCE_SEARCH.md |
| Cover, bar order, spacing, anchorage, distribution steel | Not fully checked | Applicable detailing provisions | UNVERIFIED | Current d assumes main bar outermost; no interposed transverse bar |


## Option 4: common DB20 @ 0.2 m
AUDIT RESULT: FAIL_LISTED_CHECKS; EIT compliance still UNVERIFIED
PER-CHECK AUDIT; no EIT compliance claim
H=5; H1=1.2; gamma_soil=1.8; gamma_concrete=2.4; phi=30; mu=0.6; qa=20
tt=0.25; tb=0.45; TBase=0.45; Base=3; LToe=0.6; LHeel=1.95
fc_prime=320; fy=4000; cover=0.075; passive fraction=0

| Check | Actual | Limit / rule | Comparison | Basis / verification |
| --- | --- | --- | --- | --- |
| Stem height | 4.5500 m | H-TBase | CALCULATED | H measured from base underside |
| Stem bars | DB20 @ 0.20 m | specified candidate | INPUT | Design |
| Stem effective depth | 0.3650 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Stem As | 15.7080 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Stem signed M | 9.4196 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Stem necessary yield bound | 9.4196 tf.m/m | <= 22.9336 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Stem concrete stress | 63.4954 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Stem steel stress | 1787.2756 kgf/cm2 | <= 1700.0000 | FAIL_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Stem nominal shear V/bd | 1.7016 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Stem minimum steel | 15.7080 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel length | 1.9500 m | >= 0.3000 | WITHIN_LISTED_LIMIT | Existing project geometry rule |
| Overturning FS | 3.2626  | >= 2.0000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Sliding FS | 1.9092  | >= 1.5000 | WITHIN_LISTED_LIMIT | Existing project stability criterion |
| Full contact abs(e) | 0.3149 m | <= 0.5000 | WITHIN_LISTED_LIMIT | Linear full-compression model; signed e=0.3149 positive to toe |
| q_toe / q_heel | 12.9644 / 2.9454 tf/m2 | signed pressure diagram | CALCULATED | BearingEdges |
| Bearing qa/qmax | 1.5427  | >= 1.0000 | WITHIN_LISTED_LIMIT | qa INPUT interpreted as ALLOWABLE |
| Toe bars | DB20 @ 0.20 m | specified candidate | INPUT | Design |
| Toe effective depth | 0.3650 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Toe As | 15.7080 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Toe signed M | 1.7760 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Toe necessary yield bound | 1.7760 tf.m/m | <= 22.9336 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Toe concrete stress | 11.9713 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Toe steel stress | 336.9702 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Toe nominal shear V/bd | 0.6530 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Toe minimum steel | 15.7080 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Heel bars | DB20 @ 0.20 m | specified candidate | INPUT | Design |
| Heel effective depth | 0.3650 m | t-cover-db/2 > 0 | CALCULATED | Main bar outermost; single layer assumption |
| Heel As | 15.7080 cm2/m | area per metre | CALCULATED | CalculateAsProv |
| Heel signed M | 7.8974 tf.m/m | >= 0.0000 | WITHIN_LISTED_LIMIT | Supported tension face; modShared |
| Heel necessary yield bound | 7.8974 tf.m/m | <= 22.9336 | WITHIN_LISTED_LIMIT | M<=As*fy*d; no modular ratio or code allowable used |
| Heel concrete stress | 53.2343 kgf/cm2 | <= 144.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Heel steel stress | 1498.4452 kgf/cm2 | <= 1700.0000 | WITHIN_LISTED_LIMIT | Legacy n=9; EIT clause UNVERIFIED |
| Heel nominal shear V/bd | 1.5971 kgf/cm2 | UNSET | UNVERIFIED | EIT shear limit/definition/critical section missing |
| Heel minimum steel | 15.7080 cm2/m | UNSET | UNVERIFIED | EIT applicable minimum ratio missing |
| Complete EIT 011007-19 compliance | Not established | Verified applicable clauses and complete detailing | UNVERIFIED | See audit/WSD_PRIMARY_SOURCE_SEARCH.md |
| Cover, bar order, spacing, anchorage, distribution steel | Not fully checked | Applicable detailing provisions | UNVERIFIED | Current d assumes main bar outermost; no interposed transverse bar |

Native completion: layouts=120; within known checks=57; accepted by full validator=0; WSDReviewed=False
