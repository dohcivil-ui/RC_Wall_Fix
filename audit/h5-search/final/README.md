# MEMBER_RANDOM_V1: shared BA/HCA neighborhood

UNSHIPPED CANDIDATE: this uniform-group policy was superseded by MEMBER_WEIGHTED_V1 in [../release/README.md](../release/README.md). Its results are retained as exploratory evidence and must not be presented as current production results.

The user explicitly allowed a different random neighborhood, provided every random movement in BA follows the same rule as HCA. The objective is to make the known 8538.42 Baht/m H5 solution discoverable without planting the answer or relaxing engineering checks. Requested relative convergence speed is measured honestly; the implementation does not slow HCA or give BA a hidden answer/budget advantage.

## Code change

`modShared.DrawSearchMove` is the sole source of random neighbor draws for both algorithms. It makes twelve draws in a fixed order: tb, tt, TBase, LToe, Base, the three DB/SP pairs, then a uniformly selected move group (0 through 5). The common `SearchMoveIncludes` selects which components to keep from the proposal:

0. Whole design.
1. Stem tt/tb and stem main bars.
2. TBase/Base/LToe and toe/heel main bars.
3. Stem main bars only.
4. Toe main bars only.
5. Heel main bars only.

The five geometry jump ranges remain unchanged: tt/tb/LToe +/-2, TBase +/-5, Base +/-1 index levels. Selected bar pairs draw independently and uniformly from the existing five DB and four spacing choices (20 pairs), so spacing 0.10 to 0.25 m is a legal single proposal. Unselected components remain at current. All eleven candidate values are drawn before group masking so both callers consume the same random stream. This is a changed stochastic hill-climbing neighborhood, not the previous eleven-variable-only policy.

`modBA` and `modHillClimbing` apply the same proposal and repair logic. BA alone retains active contraction bounds for tb/TBase/Base; its deterministic midpoint and recovery flow are unchanged. HCA retains full bounds and strictly improving feasible acceptance. The HCA-only final steel sweep remains absent. There is no stopping test keyed to the reference price, no special H=5 branch, no seed-specific branch, no artificial delay, and no supplied BA winner in the normal-start benchmarks. Both run.txt files now record `SearchPolicy=MEMBER_RANDOM_V1`.

## Native results, all benchmark cases reported

Every row below has a 5000-evaluation budget and normal conservative initialization. The prototype and final common-routine implementation reproduce these results exactly. This is a bounded diagnostic set, not a 30-trial research batch.

| H | Seed | HCA best | First evaluation | BA best | First evaluation |
| --- | --- | ---: | ---: | ---: | ---: |
| 3 | 12345 | 2942.34192 | 329 | 2942.34192 | 334 |
| 4 | 12345 | 4985.56464 | 360 | 4985.56464 | 1707 |
| 5 | 12345 | 8638.02148 | 1037 | 8654.44873 | 4974 |
| 5 | 12360 | 8540.19588 | 2734 | 8540.19588 | 2713 |
| 5 | 12374 | **8538.41892** | **4866** | 8856.10753 | 1567 |
| 5 | 24680 | 8757.44800 | 3399 | 8653.93045 | 1734 |
| 5 | 12355 | 8540.19588 | 2150 | 8837.54965 | 2046 |

The additional seed12355 pair was tested afterward because it was the former BA winning seed. It is preserved in reference-seed.txt and was not favorable to this uniform-group candidate.

The target was reached by HCA seed12374, with no seeded-in solution. Seed24680 was an additional check outside the user's 12345-12374 sequence. These results do not establish that BA finds 8538.42 under the new policy in the limited tested seeds, or that HCA converges more slowly. Comparing first-best evaluations at unequal final prices would not establish a speed advantage. No global optimum guarantee is made. The search policy was developed against the H5 reference; independent future experiments are necessary before making research performance claims.

The HCA target has tt=.20, tb=.40, TBase=.30, Base=3.00, LToe=1.00, LHeel=1.60 m; stem DB16@.10, toe DB16@.20, heel DB25@.25. It passes the project's existing configured engineering checks. This investigation does not extend that scope to a full standards certification.

## Verification

- `parity.txt`: 5400 comparisons of generated indices and next PRNG state across H3/H4/H5, zero differences at identical states/bounds. Wrappers are diagnostic only.
- `verification.json`: exact budgets, first 21 BA/HCA candidates identical before contraction, feasible prefix-best and first-best evaluation, unchanged four-column accept CSV and 0-based CSV numbering; HCA has only initial/neighbor entries.
- `../final-recovery`: nine additional bounded checks including impossible qa=.01 cases; 20 recovery events with no full-domain random jump; matching BA/HCA neighbor streams across recovery events when state/bounds match.
- `compile-main.log`: successful VB6 compilation of the current root production sources. Other logs compile and run native test executables.
- `../before`, `../grouped`, `../spacing3`, and `../member` preserve diagnosis and exploratory outcomes, including unsuccessful alternatives. The older README describes the pre-change diagnosis; this document describes the shipped policy.
- Production VB6 encoding/CRLF, engineering and pricing routines, evaluation counting, acceptance rules and three-variable BA contraction are checked by `audit/source_checks.py`. User result_csv files are not modified or staged.

Run `python audit/verify_member_neighborhood.py`, `python audit/recovery_random_check.py verify audit/h5-search/final-recovery`, and `python audit/source_checks.py` to verify saved evidence/source invariants. The original 30-trial user data belong to the prior policy and must not be pooled with new-policy results.
