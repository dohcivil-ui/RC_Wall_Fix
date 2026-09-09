Restored the user's original output unit labels in modShared.bas and
modProjectChecks.bas: tf/m -> ton, tf.m/m -> ton-m, tf/m2 -> ton/m2,
kgf/cm2 -> ksc. These reports use the existing one-metre wall strip convention.

Only VB6 string literals changed; removing string literals from the before/after
files gives identical source. Encoding and CRLF are preserved. No numerical
conversion, calculation, search, input, CSV schema or price change was made.

The active source compiled successfully with VB6 (compile-main.log). This edit
did not run the optimizer or a new VB6 runtime test. Earlier layout/runtime
evidence is retained without rewriting historical report files.

Follow-up: the user requested the explicit spelling kg/cm^2 instead of ksc.
Updated the input label and report string literals in Form1.frm, modShared.bas
and modProjectChecks.bas. The formulas use kg-force and centimetres (SectionStresses
converts metres to centimetres and ton-metres to kg-force-centimetres).
Only labels changed; the numerical values are unchanged. The native VB6 build
succeeded again (compile-kgcm2.log); no new calculation run was performed.
