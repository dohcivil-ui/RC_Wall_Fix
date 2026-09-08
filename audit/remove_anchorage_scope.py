"""One-time user-requested scope change; do not rerun on patched sources."""
from pathlib import Path
import re
P=Path(__file__).resolve().parent.parent
p=P/'audit/modProjectChecks.source'
s=p.read_text(encoding='ascii')
s=s.replace('PROJECT_WSD_ACI99_V1','PROJECT_WSD_ACI99_V2_NOANCHORAGE')
s=s.replace("' Normal-weight concrete, uncoated deformed bars, aggregate <=20 mm, stock 12 m.\n' Straight main-bar anchorage; no bundles, hooks, cutoffs or compression credit.", "' User research scope: no anchorage/lap checks or allowances; no formwork price.\n' Geometric member lengths only; no compression-steel credit.")
s=s.replace('    Ld(0 To 2) As Double','    Ld(0 To 2) As Double  \' Reserved CSV field: zero means excluded, not checked')
s=re.sub(r'Public Function ProjectDevelopment\([\s\S]*?End Function\n\n','',s,count=1)
selector='''Private Function SelectProjectDetail(d As Design, r As ProjectDetail) As Boolean
    Dim fb As Integer, fs As Integer, hb As Integer, hs As Integer, bb As Integer, bs As Integer
    Dim frontArea As Double, horArea As Double, score As Double, best As Double
    Dim bar As Double, gap As Double, baseBest As Double, length As Double
    best = 1E+30: baseBest = 1E+30
    length = H - d.TBase
    For fb = DB_MIN To DB_MAX
    For fs = SP_MIN To SP_MAX
        frontArea = CalculateAsProv(fb, fs)
        If frontArea >= MaxV(0.0015 * d.tb * 10000#, r.Steel(0) / 2#) And frontArea <= r.Steel(0) And SpacingOK(fb, fs, d.tt, 0) Then
            For hb = DB_MIN To DB_MAX
            For hs = SP_MIN To SP_MAX
                horArea = CalculateAsProv(hb, hs)
                bar = WP_DB(hb) / 1000#
                gap = d.tt - 2# * cover - (WP_DB(r.DB(0)) + WP_DB(fb)) / 1000# - 2# * bar
                If horArea >= 0.00125 * d.tb * 10000# And gap >= CLEAR_GAP And SpacingOK(hb, hs, d.tt, 0) Then
                    score = ProvidedSteelWeight(fb, fs, length - cover) + 2# * ProvidedSteelWeight(hb, hs, length - 2# * cover)
                    If score < best Then
                        best = score: r.FrontDB = fb: r.FrontSP = fs
                        r.HorizontalDB = hb: r.HorizontalSP = hs
                    End If
                End If
            Next hs
            Next hb
        End If
    Next fs
    Next fb
    If best = 1E+30 Then LastValidationReason = "STEM_DETAIL_LAYOUT": Exit Function
    For bb = DB_MIN To DB_MAX
    For bs = SP_MIN To SP_MAX
        bar = WP_DB(bb) / 1000#
        gap = d.TBase - 2# * cover - (WP_DB(r.DB(1)) + WP_DB(r.DB(2))) / 1000# - 2# * bar
        If CalculateAsProv(bb, bs) >= 0.002 * d.TBase * 10000# And gap >= CLEAR_GAP And SpacingOK(bb, bs, d.TBase, 0) Then
            score = 2# * ProvidedSteelWeight(bb, bs, d.Base - 2# * cover)
            If score < baseBest Then baseBest = score: r.BaseDB = bb: r.BaseSP = bs
        End If
    Next bs
    Next bb
    If baseBest = 1E+30 Then LastValidationReason = "BASE_DISTRIBUTION_LAYOUT": Exit Function
    r.ExtraWeight = best + baseBest
    SelectProjectDetail = True
End Function'''
s=re.sub(r'Private Function SelectProjectDetail\([\s\S]*?End Function',lambda m:selector,s,count=1)
(P/'audit/project_detail_selection.source').write_text(selector+'\n',encoding='ascii')
s='\n'.join(line for line in s.split('\n') if not any(k in line for k in ('r.Ld(i) = ProjectDevelopment','Then LastValidationReason = "STEM_STRAIGHT_ANCHORAGE"','Then LastValidationReason = "BASE_STRAIGHT_ANCHORAGE"')))
for i in range(3):s=s.replace(f' + r.Ld({i})','')
s=s.replace('If r.Length(i) <= 0 Or r.Length(i) > 12# Then LastValidationReason = "MAIN_STOCK_LENGTH"','If r.Length(i) <= 0 Then LastValidationReason = "INVALID_MAIN_LENGTH"')
s=s.replace("' Horizontal/longitudinal quantities are per metre, including Class B laps/12m stock.","' Geometric quantities per metre; user excludes anchorage, laps and formwork.")
s=s.replace(' kgf/cm2; straight ld=" & Format$(r.Ld(i), "0.0000") & "; main length=', ' kgf/cm2; modeled main length=')
s=s.replace(' & "; straight ld=" & r.FrontLd & " m"','')
s=s.replace(' & "; Class B lap=" & r.HorizontalLap & " m"','')
s=s.replace(' & "; Class B lap=" & r.BaseLap & " m"','')
s=s.replace('; stock=12m; no hooks/cutoffs.', '; geometric quantities only.')
s=s.replace(',7.6/7.12,12.2/12.15.', ',7.6/7.12.')
needle='    If Not d.IsValid Then\n'
s=s.replace(needle,'    s = s & "Research scope: anchorage/laps excluded from checks and steel quantities; no 0.40m allowance; formwork cost excluded." & vbCrLf\n'+needle)
p.write_text(s,encoding='ascii')
(P/'modProjectChecks.bas').write_bytes(s.replace('\n','\r\n').encode('ascii'))
p=P/'modWSD.bas'
p.write_bytes(p.read_bytes().replace(b'PROJECT_WSD_ACI99_V1',b'PROJECT_WSD_ACI99_V2_NOANCHORAGE'))
p=P/'audit/build_project_regression.py'
s=p.read_text(encoding='utf8').replace('SaveCase "short_anchorage", weak, False','SaveCase "anchorage_excluded", weak, True')
s=s.replace('    Call Verify("actual UI restores controls",', '    Call Verify("actual UI states excluded scope", InStr(text, "anchorage/laps excluded") > 0 And InStr(text, "formwork cost excluded") > 0)\n    Call Verify("actual UI restores controls",')
s=s.replace('        Call Verify(label & " extra steel included",','        Call Verify(label & " anchorage and lap allowances excluded", r.Ld(0) = 0 And r.Ld(1) = 0 And r.Ld(2) = 0 And r.FrontLd = 0 And r.HorizontalLap = 0 And r.BaseLap = 0)\n        Call Verify(label & " extra steel included",')
p.write_text(s,encoding='utf8')
p=P/'audit/verify_project_checks.py'
s=p.read_text(encoding='utf8')
s=re.sub(r'def ld\([\s\S]*?(?=def mass)','',s,count=1)
s=s.replace('development=ld(db,sp,i==2 and depth>12*.0254)','development=0 # User excludes development length from the model')
s=s.replace('near(front_ld,ld(front_db,front_sp,False));near(hor_lap,1.3*ld(hor_db,hor_sp,True));near(base_lap,1.3*ld(base_db,base_sp,True))','near(front_ld,0);near(hor_lap,0);near(base_lap,0)')
s=s.replace("assert reason['short_anchorage']=='STEM_STRAIGHT_ANCHORAGE'", "assert reason['anchorage_excluded']=='PASS_IMPLEMENTED_PROJECT_CHECKS'")
p.write_text(s,encoding='utf8')
