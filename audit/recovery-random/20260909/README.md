# BA recovery must use the HCA random neighborhood

The previous check covered GenerateNeighbor_BA but missed the full-domain Rand calls in the recovery branch. This contradicted the user's requirement that EVERY random movement in BA use the same per-variable jump rule as HCA.

Production change: modBA.bas now reopens the original three bounds and keeps the current tb, TBase and Base as the next reset position. It consumes no random draws during recovery. Subsequent random candidates come only from the unchanged GenerateNeighbor_BA. The deterministic bisection midpoint transition remains part of BA. HCA, engineering checks, costs, acceptance rules and evaluation accounting are unchanged.

Native regression: diagnostic copies insert one RecordRecovery call after the real recovery branch. CSV output is redirected to this audit directory. Other modules are compiled from the active root project. Production code has no diagnostic call.

- Before: 9 native runs at H3/H4/H5, 17 recovery events, all 17 changed position through full-domain random draws. The regression assertion failed as expected.
- After: the same 9 runs, 17 recovery events, zero recovery position changes. For the deliberately impossible qa=0.01 cases, BA neighbor candidates exactly equal the corresponding HCA candidates across recovery boundaries (excluding BA's counted deterministic reset evaluations).
- Each height uses BA/HCA 125-evaluation impossible cases and one BA 500-evaluation normal-input case. Exact budgets and prefix-best bookkeeping are checked from every trace. No 30-trial batch was run and no root result_csv data was changed.
- This verifies search behavior, not structural-code certification or global optimality. Reopening at current removes long-distance random restarts and can change search performance and the final price compared with 8c2a468.

Reproduce with `python audit/recovery_random_check.py prepare <new-folder>`, compile its RecoveryCheck.vbp with VB6, run RecoveryCheck.exe, then `python audit/recovery_random_check.py verify <new-folder>`.
The before/after module snapshots differ from production only by test instrumentation and the isolated CSV path. The before snapshot represents 8c2a468; the after snapshot contains this correction.
