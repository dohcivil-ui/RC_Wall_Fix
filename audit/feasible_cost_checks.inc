    ' Cost selection is downstream of every safety/structural check.
    ' Synthetic structural thresholds here are branch fixtures only.
    H = 5: H1 = 1.2: mu = 0.6: qa = 20
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    Call BeginSearch(5, 123, "FEASIBLE-COST-TEST")
    ok = EvaluateCandidate(d, "initial", val)
    Call AssertTrue("feasible-cost baseline accepted in fixture", ok And RunBestEvaluation = 1)
    d = Fixture(0.6, 0.7, 2, 0.5)
    Call AssertTrue("OT rejection fixture is cheaper", CalculateCost(d) < RunBestCost)
    Call AssertTrue("OT rejection fixture violates overturning", Not CheckFS_OT(d, x))
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("OT failure not priced or selected", Not ok And val = NO_SOLUTION_COST And RunBestEvaluation = 1)
    d = Fixture(0.6, 0.7, 3.5, 0.5)
    d.ASst_DB = 102: d.AStoe_DB = 102: d.ASheel_DB = 102
    Call AssertTrue("SL/BC rejection fixture is cheaper", CalculateCost(d) < RunBestCost)
    mu = 0.01
    Call AssertTrue("SL rejection fixture violates sliding", Not CheckFS_SL(d, x))
    ok = EvaluateCandidate(d, "reset", val)
    Call AssertTrue("SL failure not priced or selected", Not ok And val = NO_SOLUTION_COST And RunBestEvaluation = 1)
    mu = 0.6: qa = 1
    Call AssertTrue("BC rejection fixture violates bearing", Not CheckFS_BC(d, x, e, qt, qh))
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("BC failure not priced or selected", Not ok And val = NO_SOLUTION_COST And RunBestEvaluation = 1)
    qa = 20
    ok = EvaluateCandidate(d, "neighbor", val)
    Call AssertTrue("cheaper candidate selected after all checks satisfied", ok And RunBestEvaluation = 5 And val = RunBestCost)
    Call FinishSearch
