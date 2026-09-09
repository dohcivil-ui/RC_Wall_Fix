Restored the user's original output unit labels in modShared.bas and
modProjectChecks.bas: tf/m -> ton, tf.m/m -> ton-m, tf/m2 -> ton/m2,
kgf/cm2 -> ksc. These reports use the existing one-metre wall strip convention.

Only VB6 string literals changed; removing string literals from the before/after
files gives identical source. Encoding and CRLF are preserved. No numerical
conversion, calculation, search, input, CSV schema or price change was made.

The active source compiled successfully with VB6 (compile-main.log). This edit
did not run the optimizer or a new VB6 runtime test. Earlier layout/runtime
evidence is retained without rewriting historical report files.
