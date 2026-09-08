# VB6 CSV performance investigation

Production baseline: ce14469. User reported that 30 trials of 5000 evaluations were slower than the old program. The user was already running VB6, so all measurement runs and CSV outputs were isolated under this audit folder. User research output folders were not modified or staged.

The files here are diagnostic copies, not active project modules. Root modShared.bas is the only production source changed. The initial compile log records a test-project/module name collision; compile-2.log and subsequent logs show successful builds. Direct compilation of the user's open VBP was refused by VB6; production/MainCompile.vbp references the exact root sources and successfully compiled via a distinct VBP path without closing the user's IDE.

## Measurement

The native harness uses QueryPerformanceCounter. Profile slots: 0 full project validation, 1 base envelope, 2 stem stress envelope, 3 quantity pricing, 4 EvaluateCandidate, 5 CSV row construction/accumulation, 6 DoEvents, 7 FinishSearch. Nested slot times must not be added together. Original profile showed extra full-validation calls, but the measured CSV accumulation dominated elapsed time.

At 1024 evaluations, before the change: BA 1.679 s (CSV 1.561 s), HCA 0.619 s (CSV 0.512 s). After changing only EvaluationCSV from a growing string to an array of rows, the instrumented run took BA 0.148 s / HCA 0.072 s, with identical evaluation/accept bytes.

An independent harness using copies of final production sources (only RESULT_CSV_ROOT redirected for isolation, no production instrumentation) took BA 0.128 s / HCA 0.068 s at 1024 evaluations, and BA 0.647 s / HCA 0.556 s at 5000 evaluations. H=5, fc=320, seed=12345. These are short compiled-EXE observations on this machine, not a prediction of 30-trial duration or IDE speed. Other workloads were running concurrently.

## Cause and fix

`EvaluationCSV = EvaluationCSV & ...` copied an ever-growing detailed trace inside each evaluation. Total copying grew roughly with the square of the row count. BeginSearch now allocates one slot per evaluation plus the header, EvaluateCandidate stores its row once, and FinishSearch joins all rows once. Acceptance, seed, random calls, engineering checks, pricing and the three BA contraction variables are unchanged. The smaller accept/loop buffers are unchanged.

Run `python audit/verify_csv_performance.py audit/performance-checks/20260909-010639` from the repository root to check recorded timing margins, byte equality before/after at 1024 evaluations, and exact row counts, ordering and feasible prefix-best at 1024/5000 evaluations. See verified-performance.json and production-hashes.txt. No 30-trial or research batch was run by the agent.

The running user IDE has the prior code loaded. Reload the updated source after the current experiment finishes to use the fix; do not overwrite the new disk source with an older IDE copy.
