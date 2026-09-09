# Native VB6 multiline results control

Replaces Form1's result ListBox with the built-in TextBox, keeping its font,
position and size. MultiLine=True, ScrollBars=Vertical, Locked=True and
Enabled=True. AddResultLine appends the original text and CRLF through SelText;
there is no manual 48-character line splitting. BA/HCA completion selects the
start of the report. Clear and comparison paths use the same control.

Validation performed against the active root sources:
- MainCompile.vbp compiled successfully with the installed VB6 compiler.
- TextBoxProbe.vbp compiled and ran in VB6, without running the optimizers.
- H4/H5 fixture report text is preserved exactly; full reports match the prior
  native fixture outputs byte for byte.
- Native edit line count was 7 for a long unbroken string plus report lines;
  narrowing the control gave 10 lines, proving automatic wrapping/reflow.
- Native horizontal autoscroll is disabled; multiline, vertical scrollbar,
  locked/enabled and selected-text preservation checks passed (FAILURES=0).
- Existing scope verification confirms calculation/search modules are unchanged.

No 30-trial research batch was run. No new interactive visual screenshot was
taken. The test loads the real Form1 invisibly and inspects its native control.
Compiled executables are ignored by Git. User result_csv files are untouched.
