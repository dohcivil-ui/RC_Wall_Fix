"""Replay exports and render selected-design fixtures; no optimizer calls."""
from pathlib import Path
import hashlib, importlib.util, subprocess, sys

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('native', ROOT/'audit/verify_best_trial_export.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
native.PROBE = (OUT/'fixtures.bas').read_text(encoding='ascii')
def inventory(folder):
    return {str(p.relative_to(folder)):hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.rglob('*') if p.is_file()}
protected = inventory(ROOT/'result_csv')
folder = OUT/sys.argv[1]
native.build_project(folder,True)
subprocess.run([str(folder/'Check.exe')],timeout=40,check=True,creationflags=subprocess.CREATE_NO_WINDOW)
runtime = (folder/'runtime.txt').read_text()
print(runtime)
assert 'FAILURES=0' in runtime and 'FATAL' not in runtime
expected = {f'{prefix}-{method}-H{height}-{240+(height-3)*40}.{extension}'
            for height in (3,4,5) for method in ('BA','HCA')
            for prefix,extension in [('accept','csv'),('loopPrice','csv'),('design','bmp')]}
assert set(inventory(folder/'output')) == expected
form = (ROOT/'Form1.frm').read_bytes()
assert form.count(b'frmBestDesign.SaveResultImage(') == 2
assert b'SaveResultPicture(' not in form and b'SelectedTrialCSV' not in form
assert b'"Best cost first found at Loop " & (bestEvaluation - 1)' in (ROOT/'frmBestDesign.frm').read_bytes()
assert inventory(ROOT/'result_csv') == protected
native.build_project(OUT/(sys.argv[1]+'-gui'),False)
print('PASS: exactly 4 CSV + 2 BMP per case, fixed names overwritten; user result files unchanged.')
