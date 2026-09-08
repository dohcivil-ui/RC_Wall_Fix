"""One-time insertion of the VB6 best-geometry popup; do not rerun."""
from pathlib import Path
P=Path(__file__).resolve().parent.parent
(P/'frmBestDesign.frm').write_bytes((P/'audit/best_design_form.source').read_text(encoding='ascii').replace('\n','\r\n').encode('ascii'))
p=P/'RC_RT_HCA_v2.vbp';s=p.read_bytes();s=s.replace(b'Form=Form1.frm',b'Form=Form1.frm\r\nForm=frmBestDesign.frm');p.write_bytes(s)
p=P/'Form1.frm';s=p.read_bytes()
for prefix,algo in [('ba','BA'),('hca','HCA')]:
    line=f'{prefix}HasRun = (globalBestIteration > 0)'.encode('ascii')
    assert s.count(line)==1
    s=s.replace(line,line+f'\r\n    Call frmBestDesign.ShowBest(bestDesign, h, H1, "{algo}", CLng(globalBestTrial))'.encode('ascii'))
s+=b'\r\nPrivate Sub Form_Unload(Cancel As Integer)\r\n    Unload frmBestDesign\r\nEnd Sub\r\n'
p.write_bytes(s)
# Test projects copied from the real VBP must resolve both forms within this copy.
for name in ['audit/build_project_regression.py','audit/build_trial30.py']:
    p=P/name;s=p.read_text(encoding='ascii')
    needle=".replace('Form=Form1.frm','Form=..\\\\Form1.frm')"
    assert needle in s
    s=s.replace(needle,needle+".replace('Form=frmBestDesign.frm','Form=..\\\\frmBestDesign.frm')")
    p.write_bytes(s.replace('\n','\r\n').encode('ascii') if b'\r\n' in p.read_bytes() else s.encode('ascii'))
p=P/'audit/run_checks.ps1';s=p.read_bytes();line=b"$guiProject = $guiProject.Replace('Form=Form1.frm', 'Form=..\\Form1.frm')";assert line in s
p.write_bytes(s.replace(line,line+b"\r\n$guiProject = $guiProject.Replace('Form=frmBestDesign.frm', 'Form=..\\frmBestDesign.frm')"))
