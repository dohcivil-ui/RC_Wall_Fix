# Validated BA checkpoint and shared initial-price display

The installed `modBA.bas` is the successful adaptive-coupling candidate, with observer hooks removed. Its SHA256 is `e6ab8b41b990fb1f2cb0133adff26c1df027c1e8d0d368671632f2bb5babb6d4`. The final confirmation used the same BA source and seeds 142345–142374, paired separately for H3/fc240, H4/fc280 and H5/fc320. Every trial used 5,000 candidate evaluations, including initialization.

| Case | BA joint wins / 30 | Both methods valid | Equal initial quantity price (Baht/m) |
| --- | ---: | ---: | ---: |
| H3 / 240 | 26 | 30/30 | 29,924.364480 |
| H4 / 280 | 24 | 30/30 | 33,912.911680 |
| H5 / 320 | 21 | 30/30 | 38,041.658880 |

A joint win requires both final designs valid, unrounded BA final price <= HCA, and BA first occurrence of its own final best strictly earlier. These are results of the fixed confirmation set; they do not guarantee the same counts in later runs. GUI sessions seed from the timer once and do not automatically reproduce paired seeds.

The checkpoint includes prior authorized changes to the full catalogue initialization, local DB/spacing moves, recovery and book-based project checks. During the final save/publish task, only reporting code changed: `modShared.bas` captures initial quantity cost; `Form1.frm` retains that value with the selected trial; `modGraphing.bas` draws a labelled initial point and dashed reference until the first feasible point. BA/HCA search modules, checks, material formulas, bounds, RNG, feasible cost histories, acceptance, evaluation counts and selection rules were unchanged in this final task.

Both methods initialize tt=0.60, tb=1.00, TBase=1.00, B=7.00, toe=1.20, heel=4.80 m, and all three reinforcement sets at DB28@0.10 m. The initial design is rejected in these cases. Its quantity cost is a plotting reference, never a feasible best. CSV No.0 contains that price in `Rejected`. Subsequent convergence uses `Passed and Better value`, never a running minimum of rejected prices. CSV Loop is zero based; first-best evaluation is Loop+1.

Native verification: 12 initial-only entry calls (default/max override, both methods, three heights), zero neighbor evaluations; 20 selected-trial CSV fixtures; six existing confirmed curves and three comparison graphs replayed; two graph edge cases; full GUI compile passed. The 890 existing result files and their names were unchanged. The root EXE was not replaced. Rendering fixtures load a hidden form, not an optimization run.

Compact evidence is under `../ba-paired-win-20260910/adaptive-coupling-refinement/confirmation`: protocol, all-height summary, all 90 pairs, per-case native summaries and six selected-trial traces. Detailed development/failed-variant traces remain local; they were not merged into the confirmation.

To repeat only these checks on Windows with VB6 installed:

```powershell
C:\Python314\python.exe audit/commit-ready-20260911/verify.py
```

Each check creates a fresh directory under `audit`. CSV/BMP writes are redirected there. Production writes overwrite fixed filenames in `C:\reserch 69\RC_Wall_Fix\result_csv` only when the user runs VB6. Source encoding is preserved as legacy bytes, CRLF, no BOM.
