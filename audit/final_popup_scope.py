"""One-time final-session caption update; preserve the existing cost-first winner rule."""
from pathlib import Path
P=Path(__file__).resolve().parent.parent
p=P/'audit/best_design_form.source';s=p.read_text(encoding='ascii')
s=s.replace('Optional ByVal clearCover As Double = 0.075)', 'Optional ByVal clearCover As Double = 0.075, Optional ByVal totalTrials As Long = 1, Optional ByVal bestEvaluation As Long = 0)',1)
s=s.replace('Me.Tag = algorithm & "; trial=" & trial', 'Me.Tag = algorithm & "; trial=" & trial & "; trials=" & totalTrials & "; evaluation=" & bestEvaluation',1)
s=s.replace('"BEST DESIGN  |  " & algorithm & "  |  Trial " & trial', '"FINAL RESULT  |  " & algorithm & "  |  Selected trial " & trial & " / " & totalTrials')
s=s.replace('"RESULT  |  " & algorithm', '"FINAL RESULT  |  " & algorithm & "  |  " & totalTrials & " trials complete"')
s=s.replace('        DrawWall d, wallHeight, frontHeight, clearCover', '''        If bestEvaluation > 0 Then
            CenterText picSketch.ScaleWidth / 2, 69, "Best cost first found at evaluation " & bestEvaluation & "  |  Dimensions in metres"
        Else
            CenterText picSketch.ScaleWidth / 2, 69, "Dimensions in metres  |  Same scale in both directions"
        End If
        DrawWall d, wallHeight, frontHeight, clearCover''')
s=s.replace('    CenterText picSketch.ScaleWidth / 2, 69, "Dimensions in metres  |  Same scale in both directions"\n    DrawSection', '    DrawSection')
p.write_bytes(s.encode('ascii'));(P/'frmBestDesign.frm').write_bytes(s.replace('\n','\r\n').encode('ascii'))
p=P/'Form1.frm';s=p.read_bytes()
s=s.replace(b'globalBestCost, cover)', b'globalBestCost, cover, CLng(numTrials), globalBestIteration)')
s=s.replace(b'    ProjectTrialSummary = ""', b'    Unload frmBestDesign\r\n    ProjectTrialSummary = ""')
p.write_bytes(s)
