"""Check real initializer/CSV paths with budget 1: no neighbor or midpoint runs."""
from pathlib import Path
import hashlib, importlib.util, json, subprocess, sys

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('native', ROOT/'audit/verify_best_trial_export.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
native.PROBE = (OUT/'fixtures.bas').read_text(encoding='ascii')
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def inventory():
    return {str(p.relative_to(ROOT/'result_csv')):sha(p) for p in (ROOT/'result_csv').rglob('*') if p.is_file()}
protected = inventory()
folder = OUT/sys.argv[1]
native.build_project(folder,True)
subprocess.run([str(folder/'Check.exe')],timeout=40,check=True,creationflags=subprocess.CREATE_NO_WINDOW)
runtime=(folder/'runtime.txt').read_text()
print(runtime)
assert 'FAILURES=0' in runtime and 'FATAL' not in runtime
assert 'INITIAL_ONLY_CALLS=12; NEIGHBOR_OR_MIDPOINT_EVALUATIONS=0' in runtime
native.build_project(OUT/(sys.argv[1]+'-gui'),False)
assert inventory()==protected
print('Protected CSV/image files unchanged:',len(protected))
