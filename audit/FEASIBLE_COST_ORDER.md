# Safety and structural feasibility before cost minimization

The search objective is the lowest cost **among candidates satisfying all required checks**, not the cheapest geometry followed by a safety check.

After input/geometry and supported reinforcement validation, `CheckDesignValid` requires:

1. Overturning FS >= `FS_OT_MIN` (current project value 2.0).
2. Sliding FS >= `FS_SL_MIN` (current project value 1.5).
3. Supported full-compression base contact and `qa/qmax >= FS_BC_MIN` (current value 1.0 because the existing input is **allowable** bearing pressure).
4. Supported bending/tension faces, concrete and steel stresses, sourced WSD criteria, minimum steel and shear checks.

`EvaluateCandidate` now calls `CalculateCost` only when the complete validator returns True. A rejected candidate retains `NO_SOLUTION_COST=999999999` as a sentinel, not its estimated price. It still consumes one evaluation and records its rejection reason. Global best can only be updated inside the valid-candidate branch. Older retained trial logs may contain diagnostic prices for rejected candidates from before this change; those candidates were already excluded from best selection.

The earlier implementation already restricted best selection to valid candidates but calculated diagnostic costs for valid geometry even when safety or WSD failed. This change removes that unnecessary pricing and makes the sequence explicit. Original material pricing, cost formulas, BA three-variable contraction, other searches and evaluation budgets remain unchanged.

## Actual VB6 verification

The new `FEASIBLE-COST-TEST` uses synthetic structural criteria **only to exercise program branches**, not to establish EIT compliance. A known valid baseline is followed by physically cheaper rejected candidates:

- A candidate failing overturning is not priced or selected.
- A candidate failing sliding is not priced or selected.
- A candidate failing allowable bearing pressure is not priced or selected.
- After restoring the valid inputs, the cheaper valid candidate replaces the baseline.

The five requests consume exactly five evaluations. See [native regression](native-regression.txt), [test definitions](feasible_cost_checks.inc), and the separately retained `results/FEASIBLE-COST-TEST-*` trace.

Latest actual VB6 run: **150 checks, failures=0**; **GUI failures=0**. Independent report comparison: **44 values, mismatches=0**. No research batch was run.

The production WSD verification remains unset. Until applicable criteria are verified, a candidate passing only the three wall stability checks is not promoted to a fully accepted structural design. BA reports the lowest feasible cost it finds within its evaluation budget; it does not guarantee a global optimum.
