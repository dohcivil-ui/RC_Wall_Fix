# H5 HCA/BA search gap: native diagnosis, not a shipped algorithm change

This document records the initial diagnosis. After the user authorized a changed shared random neighborhood, MEMBER_WEIGHTED_V1 was implemented and tested; see [release/README.md](release/README.md) for the shipped code and the complete reported benchmark set. The results below remain historical evidence of the prior policy. The directory named final records an earlier, unshipped uniform-group candidate.

Production at diagnosis: commit 5eabaa5. The active RC_RT_HCA_v2.vbp references the root BA/HCA/shared modules. There are no new production changes in this investigation. User result_csv files are read only; every new native test redirects output to its own folder here.

## Latest recorded experiments

The latest user sessions contain 30 trials of 5000 evaluations for each method and each height. `latest-user-analysis.json` records the exact session and winning candidate paths. All 900,000 recorded evaluations have sequential counts, correct feasible prefix-best values and summaries matching the first occurrence of their best price.

| Height | BA best (Baht/m) | HCA best (Baht/m) |
| --- | ---: | ---: |
| 3 | 2942.34192 | 2942.34192 |
| 4 | 4985.56464 | 4985.56464 |
| 5 | 8538.41892 | 8646.93688 |

The H5 winners have the same geometry: tt=0.20, tb=0.40, TBase=0.30, Base=3.00, LToe=1.00, LHeel=1.60 m. Their main bars differ:

| Member | HCA winner | BA winner |
| --- | --- | --- |
| Stem | DB20 @ 0.15 m | DB16 @ 0.10 m |
| Toe | DB12 @ 0.10 m | DB16 @ 0.20 m |
| Heel | DB16 @ 0.10 m | DB25 @ 0.25 m |

## Native evidence and the specific mechanism

`before/probe.txt` reproduces HCA seed12360 at exactly 8646.93688 and accepts the BA winner through HCA's shared-initialization entry at 8538.41892. Thus this price is not being excluded by HCA's domain or engineering checker. The known BA answer is supplied only to a diagnostic test, never to a production search.

Holding other members fixed, actual VB6 checks find three cheaper single-member substitutions: stem gives 8573.88408, toe 8625.02104, heel 8633.38756. Changing all eleven variables every time makes it unlikely to preserve all the already-good choices while making one such substitution.

There is also a reachability restriction. The heel spacing index must move from 110 to 113, exceeding the allowed +/-2 spacing jump. `transitions/probe.txt` enumerates the 20 DB/SP combinations for each member using the native checker. With geometry and other members held fixed, every feasible heel state reachable in one +/-2 DB/SP jump from DB16@0.10 is more expensive. For example DB20@0.15 costs 8671.02456; the cheaper target DB25@0.25 costs 8633.38756 but is not a direct neighbor. A strictly improving single-member walk cannot cross that intermediate state. Simultaneous savings in other members or a different geometry trajectory can still permit a route in the full search; no claim of global unreachability is made.

BA changes its bounds and deterministically moves to bisection states, which can change the current cost and subsequent trajectory. HCA accepts only strictly cheaper feasible neighbors. Identical random step sizes therefore do not imply equal reachable trajectories or equal best prices within a finite budget. The H3/H4 observed agreement is empirical, not an equality guarantee.

## Rejected experiments, kept separate

No alternative was promoted to production:

- `grouped`: same policy applied to BA and HCA generators; draw the original eleven steps then uniformly select all variables, geometry, stem bars, toe bars or heel bars. Seed12360 from the standard start worsened to 9541.55067; from the recorded HCA winner it improved to 8551.96824 but still could not directly cross the heel-spacing gap.
- `spacing3`: both generators use +/-3 for spacing instead of +/-2. Standard-start seed12360 gave 8617.12344, but starting at the recorded HCA winner remained 8646.93688 within 5000 evaluations. Enlarging spacing alone is not demonstrated to resolve the task.

These are bounded diagnostic runs, not a new 30-trial study. No engineering criteria, prices, random seeds, production jump sizes, BA contraction variables or production CSV schemas were changed. No minimum from an experimental policy should be reported as a result of the current production algorithm.

The diagnosis does not identify a lost-best or feasibility bug to fix. Improving reliability requires an explicitly chosen change to search policy (and matching BA/HCA random movement rules), with fresh experiments. Guaranteeing an identical global minimum would require a complete/certifying search, which is outside the stated HCA/BA heuristic methods.

## Reproduce

`python audit/h5_search_analysis.py` analyzes the latest available user sessions.
`python audit/h5_search_probe.py <new-folder>` generates a redirected native H5Probe.vbp; compile it with VB6 and execute H5Probe.exe. `grouped_neighbor_experiment.py` modifies only a supplied diagnostic folder.
`python audit/verify_h5_search_diagnosis.py` verifies these saved native findings, including the improving-move barrier for heel reinforcement.
