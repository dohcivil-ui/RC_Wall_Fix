# Align HCA random neighborhood with the original BA

The user clarified that different design variables may have different step sizes, but each variable must use the same random movement rule in HCA and BA. Therefore the final 60-evaluation HCA steel sweep from 3743ba6 has been removed. Its previously reported 8538.41892 result belongs to that superseded search policy, not this aligned HCA.

## Production change

Only modHillClimbing.bas and the existing HCA policy line in modShared.bas change. BA source and all engineering checks, prices, output columns, random seed initialization, global-best handling and evaluation budgets are unchanged.

HCA uses the same eleven draws in the same order as GenerateNeighbor_BA: tb ±2, tt ±2, TBase ±5, LToe ±2, Base ±1, then the six main-bar DB/SP indices ±2. These are index steps, including zero; the physical increment follows each existing design array.

HCA now applies the same geometry repair as BA: draw tb first, draw tt second, then lower tt if it exceeds tb. Previously HCA drew tt first and raised tb before applying the height limit. The numerical step ranges were already equal, but the mapping of draws and geometry repair were different.

HCA caches its full-domain tb/TBase/Base limits at initialization and never contracts them. BA retains contraction of only those three variables. The generator bodies are checked for normalized equivalence after substituting HCA's full-domain bounds for BA's active bounds. Equality of generator rules does not guarantee equal final prices once BA changes its bounds and visits different states.

## Native evidence

- before/parity.txt: 530 differences in 5400 comparisons using the prior HCA.
- final/parity.txt: zero differences in 5400 comparisons after alignment. 450 starting states across H3/H4/H5 include lower/upper limits and cases requiring tt/tb repair. Each comparison checks all eleven generated indices and the next random number.
- The wrappers in diagnostic copies only expose the real private generators; they are not part of the production modules.
- verification.json checks the actual BA/HCA searches at H3/fc240, H4/fc280 and H5/fc320: the first 21 evaluations are identical before the first BA range update. It also verifies exact budgets, CSV schema, best bookkeeping, one 5000-evaluation HCA run and a no-solution case. H5 seed12365 returned 8773.87525 in this single run; no equality with BA or global optimum is claimed.
- Run `python audit/verify_neighbor_parity.py audit/neighbor-parity/20260909-093716` to verify the saved evidence. `audit/source_checks.py` also checks that BA remains original and the HCA generator body matches BA under the bound substitution.
- No new 30-trial research batch was run. User result_csv files were read/preserved, not overwritten or staged. All diagnostic results are isolated here.

The current run.txt policy is HCA_BA_NEIGHBOR_V1. Reports using results from older HCA policies must identify those versions separately.
