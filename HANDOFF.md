# RC_Wall_Fix handoff — 2026-09-09

## Workspace and checkpoint

- Work directly in `C:\reserch 69\RC_Wall_Fix` only; remote `https://github.com/dohcivil-ui/RC_Wall_Fix`, branch `main`.
- Latest code checkpoint: `d0f399a`, pushed and pulled successfully. This is an empty documentation checkpoint following `ee09e26` (kg/cm^2 labels), `030f4fa` (original ton labels), and `254cbba` (multiline TextBox).
- Checkpoint message: แก้ไขช่องผลลัพธ์เป็น TextBox หลายบรรทัด ตัดบรรทัดอัตโนมัติ เลื่อนอ่านและคัดลอกได้ แต่แก้ข้อความไม่ได้ ใช้หน่วยตามโค้ดคำนวณมาแสดงผล
- No pending code change at handoff. The user has NOT specified a new implementation task; read this, inspect status, then await the next instruction.
- Read AGENTS.md and JIT_code_audit.md if present (not found in earlier checks), then the root `RC_RT_HCA_v2.vbp` to identify active files.
- User-owned `result_csv` files have many deletions and untracked experiment archives. Do not stage, restore, delete, overwrite, or include them in commits. Stage only task-owned files explicitly.
- Preserve original VB6 byte encoding, CRLF and no BOM. Use byte-preserving patches, not UTF-8 rewriting of .frm/.bas. No large refactoring or unsolicited scope expansion.

## Completed UI and unit work

- `Form1.frm`: `txtResults` is a native VB.TextBox replacing lstResults; MultiLine=True, ScrollBars=2 (vertical), Locked=True, Enabled=True. Original size, position and CordiaUPC font retained.
- `AddResultLine` appends original text plus vbCrLf using SelStart/SelLength/SelText. No manual 48-character wrapping. Clear, comparison and both optimizer result paths use txtResults. BA/HCA completion sets SelStart=0.
- `modShared.FormatScreenResults` presents old-style sections: materials, optimization, price, dimensions, pressures, weights, reinforcement, safety factors, bearing, member stresses and notes. Detailed run.txt remains available.
- User explicitly requires simple original display units. Current labels: `ton`, `ton-m`, `ton/m2`, **`kg/cm^2`**. Do not restore `tf`, `tf/m`, `tf.m/m` or `ksc` display labels. The calculation is per 1 m of wall length, and kg in stress labels means kilogram-force. These were STRING-ONLY edits, with no numerical conversion.
- Input concrete-strength label also says kg/cm^2. SectionStresses uses depth*100 (cm), 100 cm section width, and moment*100000 (kg-force.cm).
- As_min is actual minimum steel, not flexural As_req; do not mislabel it. Envelope quantities and full-weight nominal bearing case are distinguished.

## Engineering and pricing scope already agreed

- H and H1 measured from base underside. Full-wall active/passive use H and H1, resultants at H/3 and H1/3. Stem calculations use H-TBase and H1-TBase of each candidate. Full active AND passive are required for this project.
- Reference H=5, TBase=.30, H1=1.20, soil unit weight1.8, phi30: stem4.70m, active moment10.3823, net with full passive9.7262 ton-m. Stem .20m, clear cover .075m, DB12@.25 must be rejected.
- SD40, fy4000, fs1700; fc allowable .45fc'; n=2040000/(15100*sqrt(fc')). Actual selected DB and spacing determine steel area and effective depth.
- Current basis PROJECT_WSD_ACI99_V3_MAIN_ONLY (provided reference plus ACI99 supplement). This is NOT certification of all EIT2562 requirements. Consult existing source notes for criteria; never invent standards.
- Stability thresholds in code: overturning >=2, sliding >=1.5, bearing qa_allowable/qmax >=1. Use actual VB6 inputs, not potentially incorrect diagram labels. 16 vertical dead-load factor .85/1.0 combinations are checked with full active/passive.
- Price includes all structural concrete (stem+base) and main stem/toe/heel steel only. Excludes anchorage, laps, secondary/horizontal steel and formwork, per explicit user direction. Do not add these or tune formulas to retain a price.

## Search, CSV and final picture constraints

- BA contracts ONLY tb, TBase, Base, never five variables. tt, LToe and reinforcement remain in shared random search.
- BA/HCA use the same MEMBER_WEIGHTED_V1 random movement helper DrawSearchMove/SearchMoveIncludes. BA-specific contraction bounds and deterministic midpoint steps remain. Do not reintroduce HCA-only steel sweeps, hardcoded target designs/prices or artificial delays.
- Evaluation budgets include initial/reset/neighbor checks; global best updates from all feasible entry paths. Seeds repeat runs. No-solution is explicit.
- Both have found H5 cost8538.41892 in bounded native evidence, with different seeds. BA is not guaranteed faster every trial; neither heuristic guarantees global optimum. Do not rig results to make HCA slower.
- Native current-policy results: `audit/h5-search/release/README.md`, `bench.txt`, `holdout.txt`, `verification.json`. H5 HCA seed12374 found8538.41892 at evaluation3162; BA seed12345 found it at4196. These DIFFERENT seeds are not a fair direct convergence comparison. `audit/h5-search/final` is an UNSHIPPED older policy; do not treat it as release evidence.
- CSV root `C:\reserch 69\RC_Wall_Fix\result_csv`. Primary names accept-BA/HCA-H3/H4/H5.csv and loopPrice-BA/HCA-H3/H4/H5.csv.
- accept schema: No.,Rejected,Passed,Passed and Better value (last trial in primary; individual trial copies preserved). loopPrice schema: No.,Loop,BestPrice, one row per trial. CSV row/evaluation indices are zero-based; internal evaluations one-based. Keep buffering/performance fixes and archives.
- Final picture appears once after all trials. Choose feasible lowest cost, tie-break on fewer evaluations to first find it. Native frmBestDesign popup with OK/X, SVG10x8cm, outlined geometry, dimensions and steel circles with DB/spacing leaders, method and price; no PNG, earth fill, force arrows or footer.

## Verification evidence and tooling

- VB6 compiler installed: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`.
- Latest compile: `audit/unit-labels/compile-kgcm2.log` succeeded against root sources via `audit/unit-labels/MainCompile.vbp`. This compile includes TextBox and final kg/cm^2 labels.
- Actual VB6 TextBox runtime test: `audit/results-textbox/runtime.txt`, FAILURES=0. H4/H5 report text preserved; native wrapped line count7 increased to10 on narrowing; locked/enabled and selected-text checks succeeded. Test loaded real Form1 invisibly, not a screenshot test or optimizer run.
- This runtime evidence predates unit string changes; after those, only compilation and source comparisons were rerun. Do not claim a new runtime or fresh 30-trial test.
- `python audit/verify_results_layout.py` verifies historical native evidence and current presentation-only scope, normalizing authorized unit label changes. It is NOT a new native runtime test.
- Before TextBox: historical `audit/results-layout` native ListBox evidence had38 clipped rows before manual wrapping,0 after. Keep historical reports intact; do not overwrite them to match new labels.
- No fresh 30-trial research batch is authorized; user prefers running those themselves. Bounded diagnostic tests only when justified. Keep outputs isolated under audit, never result_csv.
- Do not run old change_results_textbox.py or results_layout_edit.py blindly: they are one-time migration/prototype scripts, not current application entry points.
- If VB6 IDE holds stale project files open, reload them; do not save stale editor state over current disk source. Compile separate audit projects with absolute source paths; don't interrupt user IDE.
- For future work: state what actually compiled/ran, give evidence, and list limits. Do not claim PASS without evidence or imply that UI checks certify structural safety.

## Suggested first prompt in the new chat

ทำงานต่อในโปรเจกต์ C:\reserch 69\RC_Wall_Fix โดยใช้โฟลเดอร์นี้โดยตรง อ่าน HANDOFF.md และ AGENTS.md ถ้ามีก่อน แล้วตรวจ git status และไฟล์ .vbp ยืนยันสถานะล่าสุด ห้ามแตะผลทดลอง result_csv ห้ามเปลี่ยนสูตร อัลกอริทึม หน่วย หรือขอบเขตงานเพิ่มเติมเอง และยังไม่รัน 30 trials เมื่ออ่านครบแล้วสรุปสถานะสั้น ๆ และรอคำสั่งงานถัดไปจากผม
