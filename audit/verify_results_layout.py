"""Check native result-panel layout evidence and presentation-only source scope."""
from pathlib import Path
import re,json,subprocess,hashlib
P=Path(__file__).resolve().parent.parent
D=P/'audit/results-layout'
before=(D/'before/layout.txt').read_text(encoding='latin1')
after=(D/'after/layout.txt').read_text(encoding='latin1')
assert 'CLIPPED_ROWS=38; FAILURES=0' in before
assert 'CLIPPED_ROWS=0; FAILURES=0' in after and 'FATAL=' not in after
for height in (4,5):
    assert (D/f'before/full-report-H{height}.txt').read_bytes()==(D/f'after/full-report-H{height}.txt').read_bytes()
for section in ('MATERIAL PROPERTIES','OPTIMIZATION RESULTS','TOTAL COST','DIMENSIONS','EARTH PRESSURES','WEIGHTS','STEEL REINFORCEMENT','SAFETY FACTORS','BEARING CAPACITY'):
    assert after.count(f'=== {section} ===')==2
assert 'As_min=' in after and 'not flexural As_req' in after
assert 'q_max (16-case envelope)' in after and '--- Full-weight case ---' in after
b=(P/'modShared.bas').read_bytes()
start=b.index(b'Private Function ScreenResultRow(');end=b.index(b'Public Function FormatResults(',start)
assert b[:start]+b[end:]==(D/'before/modShared.bas').read_bytes()
assert b[start:end].replace(b'\r\n',b'\n')==(P/'audit/results_screen_sections.source').read_bytes().replace(b'\r\n',b'\n')
form=(P/'Form1.frm').read_bytes()
if b'Begin VB.TextBox txtResults' in form:
    probe=(P/'audit/results-textbox/runtime.txt').read_text()
    assert 'FAILURES=0' in probe and 'FATAL=' not in probe
    assert probe.count('TEXT_PRESERVED=True')==2
    for height in (4,5):
        assert (P/f'audit/results-textbox/full-report-H{height}.txt').read_bytes()==(D/f'after/full-report-H{height}.txt').read_bytes()
    form=form.replace(b'Begin VB.TextBox txtResults',b'Begin VB.ListBox lstResults')
    for prop in (b"         MultiLine       =   -1  'True\r\n",b"         ScrollBars      =   2  'Vertical\r\n",b"         Locked          =   -1  'True\r\n",b"         Enabled         =   -1  'True\r\n"):
        assert form.count(prop)==1
        form=form.replace(prop,b'')
    form=form.replace(b'txtResults.Text = vbNullString',b'lstResults.Clear')
    form=form.replace(b'txtResults.SelStart = 0',b'lstResults.TopIndex = 0')
assert form.count(b'lstResults.TopIndex = 0')==2
original=subprocess.check_output(['git','show','e4c2e79:Form1.frm'],cwd=P)
# Restore just the UI substitutions to establish that optimizer/event flow is unchanged.
clean=re.sub(rb'Public Sub AddResultLine\([\s\S]*?End Sub\r\n\r\n',b'',form)
clean=clean.replace(b'AddResultLine ',b'lstResults.AddItem ')
clean=clean.replace(b'modShared.FormatScreenResults(bestDesign, selectedMaterial, "BA")',b'modShared.FormatResults(bestDesign, selectedMaterial, "Bisection Algorithm v1.0")')
clean=clean.replace(b'FormatScreenResults(bestDesign, selectedMaterial, "HCA")',b'FormatResults(bestDesign, selectedMaterial)')
clean=clean.replace(b'    lstResults.TopIndex = 0\r\n',b'')
assert clean==original
for name in ('modBA.bas','modHillClimbing.bas','modProjectChecks.bas','modWSD.bas'):
    assert (P/name).read_bytes()==subprocess.check_output(['git','show','e4c2e79:'+name],cwd=P)
for name in ('Form1.frm','modShared.bas'):
    data=(P/name).read_bytes();assert data.count(b'\n')==data.count(b'\r\n') and not data.startswith(b'\xef\xbb\xbf')
assert 'succeeded' in (D/'compile-main.log').read_text().lower()
result={'native_clipped_rows_before':38,'native_clipped_rows_after':0,'native_layout_failures':0,
        'full_reports_H4_H5_unchanged':True,'optimizer_and_engineering_code_unchanged':True,
        'hashes':{name:hashlib.sha256((P/name).read_bytes()).hexdigest() for name in ('Form1.frm','modShared.bas')}}
(D/'verification.json').write_text(json.dumps(result,indent=2))
print('Native VB6 layout verified: 38 clipped rows before, 0 after; H4/H5 full reports byte-identical; calculations/search unchanged.')
