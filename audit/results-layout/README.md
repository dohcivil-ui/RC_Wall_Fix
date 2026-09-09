# Restore the original sectioned VB6 results panel

The user supplied screenshots of the older list layout and requested that readable section order. The active project uses Form1.lstResults (a VB6 ListBox). It clips long lines and has no automatic wrapping; the project report had combined many fields and explanatory notes into single long lines.

Production changes are limited to Form1.frm and modShared.bas:

- FormatScreenResults uses the existing CheckProjectDesign result and existing pressure/weight/bearing functions. It restores the old headings in order: MATERIAL PROPERTIES, OPTIMIZATION RESULTS, TOTAL COST, DIMENSIONS, EARTH PRESSURES, WEIGHTS, STEEL REINFORCEMENT, SAFETY FACTORS and BEARING CAPACITY, followed by member stresses/notes.
- The original CordiaUPC font, control size and overall form layout remain unchanged. AddResultLine wraps rows at a word boundary within 48 characters, splitting long unbroken file paths without losing characters. All list output, including progress/CSV paths/comparisons, uses this helper. The two run handlers reset TopIndex to zero after adding final output.
- Display precision is generally two decimals; dimensionless k/j/Ka/Kp, clear cover/effective depth and eccentricity retain three where appropriate. Internal values and full reports are unchanged.
- As_min and As_prov come directly from the actual member checker. As_min is explicitly not mislabeled as the former flexural As_req. Actual bar stresses remain visible.
- Safety-factor minima and the maximum bearing pressure use the existing 16-case envelope. The full-weight eccentricity/qmax/qmin are separately labeled. Allowable-bearing threshold comes from FS_BC_MIN (=1), not the obsolete screenshot's factor 2. No old engineering values or old constant n=9 are restored.
- The full FormatResults/ProjectDesignReport text in run.txt remains untouched. No CSV schema/path, optimizer, price, popup or engineering-acceptance change is included.

## Native evidence

`before/layout.txt` uses the former long-line report; VB6 measured 38 rows wider than the actual list width across the H4/H5 fixtures. `after/layout.txt` uses the sectioned report; zero rows exceed that width, with zero fixture failures. The width check uses VB6 TextWidth with the real control font, accounting for borders and vertical scrollbar. Long-path wrapping is checked for exact character preservation, and no-solution/rejected designs cannot produce a PASS summary.

The full-report-H4.txt and full-report-H5.txt files before/after are byte-identical. compile-main.log confirms native compilation of the current root production project through a separate VBP. The form was also inspected with Computer Use after native compilation; its section headings and short rows fit the original panel. Lower sections were read through the window's accessibility tree after user input interrupted an attempted scroll.

The fixtures are saved accepted H4/H5 designs, not new optimization trials. Algorithm/trial/evaluation text in this GUI harness is illustrative test metadata (the display was exercised with the BA label); it must not be cited as an actual BA timing result. No optimization or research CSV writing is performed by the harness. The diagnostic form copies set ShowInTaskbar=True and move the window on-screen so it can be captured; production Form1 keeps its original window properties.

Run `python audit/verify_results_layout.py` to check the saved native evidence and prove that removing only the presentation additions restores the original source. `audit/source_checks.py` also checks existing engineering/search invariants and VB6 encoding. The earlier results_layout_edit.py records the superseded compact-layout draft; results_screen_sections.source is the final screen formatter.

The 30-trial research batch was not run for this presentation change. Existing user result_csv files were not modified or staged.
