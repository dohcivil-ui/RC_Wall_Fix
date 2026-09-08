"""One-time scope correction requested by user: main stem/toe/heel steel only."""
from pathlib import Path
import re
P=Path(__file__).resolve().parent.parent
def update(path,s):
    p=P/path
    old=p.read_bytes()
    p.write_bytes(s.replace('\n','\r\n').encode('ascii') if b'\r\n' in old else s.encode('ascii'))
p=P/'audit/modProjectChecks.source'
s=p.read_text(encoding='ascii').replace('PROJECT_WSD_ACI99_V2_NOANCHORAGE','PROJECT_WSD_ACI99_V3_MAIN_ONLY')
s=s.replace("' Geometric member lengths only; no compression-steel credit.","' Main stem/toe/heel reinforcement only; geometric lengths; no other steel.")
s=re.sub(r'^    (?:FrontDB|FrontSP|HorizontalDB|HorizontalSP|BaseDB|BaseSP|FrontLd|HorizontalLap|BaseLap|ExtraWeight) As \w+\n','',s,flags=re.M)
s=re.sub(r'Private Function SelectProjectDetail\([\s\S]*?End Function\n\n','',s,count=1)
s=s.replace('    Dim topDepth As Double, beta As Double, rhoMax As Double, y As Double, maxSpacingStress As Double','    Dim topDepth As Double, beta As Double, rhoMax As Double, y As Double')
s=s.replace('    Dim frontArea As Double, frontBar As Double, horBar As Double, baseBar As Double, gap As Double\n','')
s=s.replace('    If hs <= 2# * cover Or d.Base <= 2# * cover Then Exit Function\n','')
s=s.replace('        If Not SpacingOK(r.DB(i), r.SP(i), thickness, r.FsBound(i)) Then', '        If thickness < 2# * cover + WP_DB(r.DB(i)) / 1000# Then LastValidationReason = "MAIN_BAR_COVER": Exit Function\n        If Not SpacingOK(r.DB(i), r.SP(i), thickness, r.FsBound(i)) Then')
s=s.replace("    ' Derive the least-weight feasible detail within the existing discrete bar catalog.\n    ' Both faces are provided; front bars are not credited as compression reinforcement.\n    If Not SelectProjectDetail(d, r) Then Exit Function\n", "    ' Only the three main-bar selections are included in the research model.\n")
s=s.replace('(r.MainWeight + r.ExtraWeight)', 'r.MainWeight')
s=s.replace('Research scope: anchorage/laps excluded', 'Research scope: main stem/toe/heel steel only; secondary steel excluded; anchorage/laps excluded')
s='\n'.join(line for line in s.split('\n') if not any(x in line for x in ('"Stem front vertical:', '"Stem horizontal EACH face:', '"Base longitudinal EACH face:')))
s=s.replace(' & "; extra steel=" & Format$(r.ExtraWeight, "0.0000")', '')
s=s.replace('Concrete+provided-steel estimate=', 'Concrete+main-steel estimate=')
s=s.replace('main bars outermost, distribution inside; two curtains/mats;', 'single main layer per member;')
s=s.replace(',14.3.3,7.6/7.12', ',7.6/7.12')
update('audit/modProjectChecks.source',s)
(P/'modProjectChecks.bas').write_bytes(s.replace('\n','\r\n').encode('ascii'))
p=P/'modWSD.bas';p.write_bytes(p.read_bytes().replace(b'PROJECT_WSD_ACI99_V2_NOANCHORAGE',b'PROJECT_WSD_ACI99_V3_MAIN_ONLY'))
p=P/'audit/build_project_regression.py';s=p.read_text(encoding='ascii')
s=s.replace(' And r.FrontLd = 0 And r.HorizontalLap = 0 And r.BaseLap = 0','')
s=s.replace('Call Verify(label & " extra steel included", r.ExtraWeight > 0 And r.Cost > r.ConcreteVolume * currentMaterial.concretePrice + r.MainWeight * currentMaterial.SteelPrice)', 'Call Verify(label & " concrete plus main steel only", Abs(r.Cost - (r.ConcreteVolume * currentMaterial.concretePrice + r.MainWeight * currentMaterial.SteelPrice)) < 0.000001)')
s='\n'.join(line for line in s.split('\n') if not any(x in line for x in ('Print #f, "front,', 'Print #f, "horizontal,', 'Print #f, "base,')))
s=s.replace('CsvNumber(r.ExtraWeight)', 'CsvNumber(0)') # preserve totals schema; excluded quantity is zero
s=s.replace('SaveCase "curtain_clearance", weak, False', 'SaveCase "secondary_excluded", weak, True')
s=s.replace('InStr(text, "Stem horizontal EACH face") > 0', 'InStr(text, "secondary steel excluded") > 0 And InStr(text, "Stem horizontal EACH face:") = 0 And InStr(text, "Stem front vertical:") = 0 And InStr(text, "Base longitudinal EACH face:") = 0')
update('audit/build_project_regression.py',s)
p=P/'audit/verify_project_checks.py';s=p.read_text(encoding='ascii')
s=s.replace("'H5','anchorage_excluded'", "'H5','anchorage_excluded','secondary_excluded'")
s=re.sub(r'    front_db,front_sp,front_ld=[\s\S]*?(?=    concrete=)', '    assert set(extra)=={\'totals\'} # no secondary steel in current detail export\n    extra_weight=0\n',s,count=1)
s=s.replace('assert len(summaries)==4','assert len(summaries)==5')
s=s.replace("assert reason['anchorage_excluded']=='PASS_IMPLEMENTED_PROJECT_CHECKS'", "assert reason['anchorage_excluded']=='PASS_IMPLEMENTED_PROJECT_CHECKS'\nassert reason['secondary_excluded']=='PASS_IMPLEMENTED_PROJECT_CHECKS'")
s=s.replace('and anchorage-excluded case verified', 'and excluded-detail cases verified')
update('audit/verify_project_checks.py',s)
p=P/'audit/run_project_checks.ps1';s=p.read_text(encoding='ascii').replace('H[345]|anchorage_excluded','H[345]|anchorage_excluded|secondary_excluded');update('audit/run_project_checks.ps1',s)
