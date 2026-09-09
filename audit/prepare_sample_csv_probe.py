from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'audit/sample-csv-export'
PROBE = OUT / 'probe'
data = (ROOT / 'modShared.bas').read_bytes()
old = b'Public Const RESULT_CSV_ROOT As String = "C:\\reserch 69\\RC_Wall_Fix\\result_csv"'
new = ('Public Const RESULT_CSV_ROOT As String = "' + str(PROBE / 'output') + '"').encode('ascii')
assert data.count(old) == 1
(PROBE / 'modShared.bas').write_bytes(data.replace(old, new))
probe = (ROOT / 'audit/sample_csv_probe.bas').read_text(encoding='ascii')
(PROBE / 'SampleCsvProbe.bas').write_bytes(probe.replace('\n', '\r\n').encode('ascii'))
data = (OUT / 'production/Production.vbp').read_text(encoding='ascii')
data = data.replace('Startup="Form1"', 'Startup="Sub Main"')
data = data.replace('ExeName32="RC_RT_HCA_v2.exe"', 'ExeName32="SampleCsvProbe.exe"')
data = data.replace(str(ROOT / 'modShared.bas'), str(PROBE / 'modShared.bas'))
data += 'Module=SampleCsvProbe; ' + str(PROBE / 'SampleCsvProbe.bas') + '\n'
(PROBE / 'Probe.vbp').write_bytes(data.replace('\n', '\r\n').encode('ascii'))
print('Prepared isolated export-only sample replay; no optimizer calls.')
