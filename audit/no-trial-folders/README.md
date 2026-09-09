# Stop automatic per-trial folders

User request: stop creating BA/HCA timestamp-and-seed directories shown in the screenshot. Preserve primary accept/loopPrice CSVs and archive; do not touch existing experiments or run 30 trials.

Production changes remove automatic run-directory creation and detailed evaluation/run reports. RecordProjectTrial skips implicit output without a folder or an explicitly supplied summary path. Both methods save each trial's accept CSV using the existing fixed-name writer, which archives the previous primary file. UI references to removed reports are omitted.

## Verification

- `MainCompile.vbp` references current production sources by absolute path. `compile-main.log`: build succeeded. No root executable was replaced.
- `ExportProbe.vbp` uses those sources except a local `modShared.bas` copy with only RESULT_CSV_ROOT redirected to this directory's `output` folder. `compile-probe.log`: build succeeded.
- `ExportProbe.bas` exercises two export cycles for each method, one invalid candidate per cycle, synthetic accepted/improved CSV rows, summary suppression, fixed filenames and accumulated loop rows. It does not run the BA/HCA optimizer loops or a research batch.
- `runtime.txt`: FAILURES=0.
- `verification.json`: output contains exactly four primary files and the archive directory, both trial copies remain archived for each method, no detailed reports exist, result_csv Git status is unchanged, and production VB6 non-ASCII bytes/CRLF/no-BOM are preserved. BA/HCA module edits only add the fixed-name argument to accept export.
- `git -c core.whitespace=cr-at-eol diff --check` passed, honoring the required VB6 CRLF.

The output CSVs here contain diagnostic data, not experimental findings or structural verification. Historical evidence in other audit directories remains unchanged.
