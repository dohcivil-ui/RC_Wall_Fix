# Remove seed controls and fixed replay

The user explicitly requested removing Seed after repeated clicks replayed the same30 trials. This change removes fixed replay throughout the active root VB6 project, including batch call sites, while preserving the agreed search/engineering rules.

## Behavior

- No seed field, label, input validation, status line, comparison-key component or optimizer seed argument remains in active VBP sources.
- A single guarded initializer calls parameterless Randomize once per process. The form initializes before selecting random material, and direct optimizer/batch entry points also initialize when necessary.
- The existing VB6 random stream continues across trials and subsequent clicks. There is no per-trial reset or incrementing seed schedule. A newly opened program initializes automatically from the system timer.
- Same-input runs can still reach identical prices or best-found iterations; the change removes forced replay, not the possibility of coincident results.
- Main accept/loopPrice schemas, filenames and archive preservation remain intact. Seed fields were removed only from optional project-summary/batch reports. No old output file was rewritten.

## Narrow source scope

The BA/HCA module differences relative to the task's starting snapshot are only removal of the optional RandomSeed argument and the matching BeginSearch argument. DrawSearchMove, SearchMoveIncludes, Rand, EvaluateCandidate, FinishSearch, screen formatting and the CSV writer are unchanged. The previously fixed unchanged-midpoint guard remains. Structural calculations, units and price scope remain unchanged.

## Native verification

`python audit/run_fresh_random_probe.py before` reproduced five failures: seed controls still present, repeated BA/HCA button clicks using identical candidate sequences, and repeated direct BA/HCA calls using identical sequences.

`python audit/run_fresh_random_probe.py final` passed31 native checks. The harness loads the real Form1 and invokes its actual click handlers: two clicks per method, two trials per click,64 evaluations per trial. It also calls each optimizer directly twice with64 evaluations. All output goes to an isolated audit directory. The diagnostic copy suppresses showing the final sketch window but keeps its rendering logic; production UI was not altered for testing.

`python audit/verify_fresh_random.py` passes72 checks, covering changed candidate sequences between clicks and trials, initial candidate preservation, exact evaluation budgets, unchanged primary CSV schemas/indices, seed removal, once-only initialization, unchanged search/formula routines, preserved non-ASCII source bytes/CRLF/no-BOM, and absence of production diagnostic hooks.

`production/compile.log` confirms successful compilation against root sources using an audit VBP with absolute paths. The root executable was not replaced. `final` is the completed snapshot; `before` is the baseline. The intermediate `after` snapshot was created after a preparation script failed before making changes and therefore still records baseline behavior.

No30-trial research batch was run, and no result_csv file was modified. These are short behavioral diagnostics, not new comparative BA/HCA research results. Historical seed-based evidence remains unchanged and cannot be treated as tests of the new automatic-random behavior.
