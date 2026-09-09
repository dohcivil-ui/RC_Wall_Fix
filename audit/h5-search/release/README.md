# Released shared search policy: MEMBER_WEIGHTED_V1

The user authorized changing the random neighborhood, requiring HCA and every random movement in BA to use the same rules. The implementation makes the H5 reference price discoverable through normal stochastic search. It does not artificially delay HCA to produce a preferred comparison, inject the known design, stop at a target price or add an H5/seed-specific rule.

## Changes

- `modShared.DrawSearchMove` contains all twelve neighbor draws. Geometry index jumps are tt/tb/LToe +/-2, TBase +/-5 and Base +/-1. Every main-bar pair is drawn uniformly from the existing five DB and four spacing choices; this removes the old inability to propose spacing 0.10 to 0.25 m directly.
- The whole design is selected with probability 7/12. Each of five component moves has probability 1/12: stem geometry plus stem bars; base geometry plus toe/heel bars; stem bars alone; toe bars alone; heel bars alone. `SearchMoveIncludes` is shared by BA/HCA. Unselected components remain at current.
- `modBA` and `modHillClimbing` call this common draw routine and apply equivalent bounds/repair logic. Their initialization, acceptance, global-best handling and evaluation accounting are unchanged. BA retains only tb/TBase/Base contraction, deterministic midpoint transitions and no-random-jump recovery. HCA remains strict stochastic hill climbing within full bounds.
- Both run.txt files record MEMBER_WEIGHTED_V1. No final steel sweep, extra candidate evaluation, artificial delay or engineering/cost change is introduced. CSV columns, popup behavior and user output paths remain unchanged.

## Why this policy was selected

The original 11-variable-only neighborhood missed cheaper bar substitutions and had a +/-2 spacing barrier. Uniformly weighting all six move groups let HCA reach the reference but BA did not reach it in the bounded development cases. Increasing whole-design proposals to 7/12 retains more coupled exploration while keeping direct member improvements available. Both methods then reached the reference within the same 5000-evaluation budget. This selection used the H5 benchmark and is algorithm development, not an independent research comparison. All tested alternatives and failures are preserved in adjacent folders.

## All final benchmark results (Baht/m)

All runs start from the normal conservative design and have exactly 5000 evaluations. No BA result is supplied to HCA in these runs. Evaluations are 1-based; the primary CSV Loop convention is one less.

| H | Seed | HCA best | First evaluation | BA best | First evaluation |
| --- | --- | ---: | ---: | ---: | ---: |
| 3 | 12345 | 2942.34192 | 299 | 2942.34192 | 283 |
| 4 | 12345 | 4985.56464 | 279 | 5034.13488 | 2695 |
| 5 | 12345 | 8556.48468 | 4403 | **8538.41892** | **4196** |
| 5 | 12360 | 8540.19588 | 3793 | 8540.19588 | 3930 |
| 5 | 12374 | **8538.41892** | **3162** | 8551.96824 | 2537 |
| 5 | 24680 | 8540.19588 | 1292 | **8538.41892** | **3904** |

Additional seed24681 was not used to choose the policy:

| H | HCA best | First evaluation | BA best | First evaluation |
| --- | ---: | ---: | ---: | ---: |
| 3 | 2942.34192 | 236 | 2942.34192 | 240 |
| 4 | 4985.56464 | 3046 | 4985.56464 | 3379 |
| 5 | 8540.19588 | 2133 | 8540.19588 | 1804 |

Both algorithms reach 8538.42 in the tested set, but not in every seed and not in the same tested seed. The data do NOT establish that HCA is slower than BA at the target price. At equal prices, either method can be faster. Comparing first-best evaluations at different final prices is not a valid speed comparison. The H4 seed12345 BA run misses the prior best within this budget; H4 seed24681 reaches it. A finite-budget heuristic has no promise of equal minima for every trial, and this work does not certify the global optimum.

The target geometry is tt=.20, tb=.40, TBase=.30, Base=3.00, LToe=1.00 and LHeel=1.60 m; stem DB16@.10, toe DB16@.20, heel DB25@.25. Existing configured project checks accept it. Engineering scope remains the project's main-steel-only model, not full standards certification.

## Evidence

- `compile-MainCompile.log`: native VB6 production compile; diagnostic executables were also compiled and run with VB6.
- `parity.txt`: 5400 comparisons, zero differences for equal states/bounds, including the next PRNG state. Test wrappers are absent from production.
- `verification.json`: all 18 final normal-start runs, exact budgets, first 21 BA/HCA candidates identical before contraction, feasible prefix-best, first-best evaluation, unchanged accept CSV schema and HCA acceptance flow. `bench.txt` and `holdout.txt` preserve complete outputs, including worse results.
- `../release-recovery`: impossible cases, recovery/no-solution and shared random-stream checks.
- `audit/source_checks.py`: shared RNG entry point, equivalent neighbor logic, only three BA contraction bounds, unchanged optimizer flow/engineering/cost routines and preserved VB6 byte encoding/CRLF.
- User research CSV files are read only and excluded from the commit. No new full 30-trial batch was run. The user can reopen the project and run their 30 trials under the new policy; prior-policy research data must not be pooled with these results.

Verify with `python audit/verify_shared_search.py`, `python audit/recovery_random_check.py verify audit/h5-search/release-recovery`, and `python audit/source_checks.py`. Earlier `final` and `final-recovery` folders refer to the unshipped uniform-group candidate, not this released weighted policy.
