"""One-time popup update: outline, actual main bars, winning price; no formula changes."""
from pathlib import Path
import re
P=Path(__file__).resolve().parent.parent
p=P/'audit/best_design_form.source';s=p.read_text(encoding='ascii')
s=s.replace('ByVal trial As Long)', 'ByVal trial As Long, Optional ByVal bestCost As Double = 0, Optional ByVal clearCover As Double = 0.075)',1)
s=s.replace('    DrawWall d, wallHeight, frontHeight','    DrawWall d, wallHeight, frontHeight, clearCover',1)
s=s.replace('        DrawWall d, wallHeight, frontHeight, clearCover', '''        If ValidBars(d) Then
            picSketch.Font.Bold = True
            If bestCost > 0 Then
                CenterText picSketch.ScaleWidth / 2, 44, "COST  " & Format$(bestCost, "#,##0.00") & " Baht/m  (concrete + main steel)"
            Else
                CenterText picSketch.ScaleWidth / 2, 44, "Drawing fixture - price not supplied"
            End If
            picSketch.Font.Bold = False
            Me.Tag = Me.Tag & "; price=" & bestCost & "; Stem: " & RebarText(d.ASst_DB, d.ASst_Sp) & "; Toe: " & RebarText(d.AStoe_DB, d.AStoe_Sp) & "; Heel: " & RebarText(d.ASheel_DB, d.ASheel_Sp)
        End If
        DrawWall d, wallHeight, frontHeight, clearCover''')
s=s.replace('Private Sub DrawWall(d As Design, wallHeight As Double, frontHeight As Double)', 'Private Sub DrawWall(d As Design, wallHeight As Double, frontHeight As Double, clearCover As Double)')
s=s.replace('picSketch.ScaleHeight - 235', 'picSketch.ScaleHeight - 275')
s=s.replace(' / 2, 44, "Dimensions', ' / 2, 69, "Dimensions')
s=s.replace('    DrawHorizontal top, back', '    If ValidBars(d) Then DrawMainBars d, pixelsPerMetre, clearCover, x0, x1, toe, back, y0, yTop, yBase\n    DrawHorizontal top, back',1)
s=s.replace('H and H1 are measured from the base underside.', 'Main bars schematic; spacing is along the wall. H/H1 measured from base underside.')
s=re.sub(r'    Dim soil As Long[\s\S]*?(?=    picSketch.DrawWidth = 2)', '    Dim edge As Long\n    edge = RGB(35, 35, 35)\n',s,count=1)
insert='''Private Function ValidBars(d As Design) As Boolean
    ValidBars = d.ASst_DB >= DB_MIN And d.ASst_DB <= DB_MAX And d.ASst_Sp >= SP_MIN And d.ASst_Sp <= SP_MAX And d.AStoe_DB >= DB_MIN And d.AStoe_DB <= DB_MAX And d.AStoe_Sp >= SP_MIN And d.AStoe_Sp <= SP_MAX And d.ASheel_DB >= DB_MIN And d.ASheel_DB <= DB_MAX And d.ASheel_Sp >= SP_MIN And d.ASheel_Sp <= SP_MAX
End Function

Private Function RebarText(DB As Integer, SP As Integer) As String
    RebarText = "DB" & WP_DB(DB) & " @ " & Format$(WP_SP(SP), "0.000") & " m"
End Function

Private Sub DrawMainBars(d As Design, pixelsPerMetre As Double, clearCover As Double, x0 As Double, x1 As Double, toe As Double, back As Double, y0 As Double, yTop As Double, yBase As Double)
    Dim stemX As Double, toeY As Double, heelY As Double, endCover As Double
    endCover = clearCover * pixelsPerMetre
    stemX = back - (clearCover + WP_DB(d.ASst_DB) / 2000#) * pixelsPerMetre
    toeY = y0 - (clearCover + WP_DB(d.AStoe_DB) / 2000#) * pixelsPerMetre
    heelY = yBase + (clearCover + WP_DB(d.ASheel_DB) / 2000#) * pixelsPerMetre
    picSketch.DrawWidth = 3
    picSketch.Line (stemX, yTop + endCover)-(stemX, yBase)
    picSketch.Line (x0 + endCover, toeY)-(toe, toeY)
    picSketch.Line (back, heelY)-(x1 - endCover, heelY)
    picSketch.DrawWidth = 1
    RebarCallout x1 + 65, yTop + 75, "Stem - back face", RebarText(d.ASst_DB, d.ASst_Sp), stemX, yTop + (yBase - yTop) * 0.42, False
    RebarCallout x0 - 110, yBase - 130, "Toe - bottom", RebarText(d.AStoe_DB, d.AStoe_Sp), (x0 + endCover + toe) / 2, toeY, True
    RebarCallout x1 + 65, yBase - 80, "Heel - top", RebarText(d.ASheel_DB, d.ASheel_Sp), (back + x1 - endCover) / 2, heelY, False
End Sub

Private Sub RebarCallout(x As Double, y As Double, title As String, bars As String, targetX As Double, targetY As Double, pointRight As Boolean)
    Dim startX As Double
    picSketch.Font.Bold = True
    picSketch.CurrentX = x: picSketch.CurrentY = y: picSketch.Print title
    picSketch.Font.Bold = False
    picSketch.CurrentX = x: picSketch.CurrentY = y + 18: picSketch.Print bars
    startX = x - 8: If pointRight Then startX = x + picSketch.TextWidth(bars) + 8
    DrawLeader startX, y + 25, targetX, targetY
End Sub

Private Sub DrawLeader(x As Double, y As Double, targetX As Double, targetY As Double)
    Dim length As Double, ux As Double, uy As Double
    length = Sqr((targetX - x) ^ 2 + (targetY - y) ^ 2)
    If length = 0 Then Exit Sub
    ux = (targetX - x) / length: uy = (targetY - y) / length
    picSketch.Line (x, y)-(targetX, targetY)
    picSketch.Line (targetX - 8 * ux + 3 * uy, targetY - 8 * uy - 3 * ux)-(targetX, targetY)
    picSketch.Line -(targetX - 8 * ux - 3 * uy, targetY - 8 * uy + 3 * ux)
End Sub

'''
s=s.replace('Private Sub DrawHorizontal(',insert+'Private Sub DrawHorizontal(',1)
p.write_bytes(s.encode('ascii'));(P/'frmBestDesign.frm').write_bytes(s.replace('\n','\r\n').encode('ascii'))
p=P/'Form1.frm';s=p.read_bytes();s=s.replace(b'CLng(globalBestTrial))',b'CLng(globalBestTrial), globalBestCost, cover)');p.write_bytes(s)
p=P/'audit/SketchRegression.source';s=p.read_text(encoding='ascii')
s=s.replace('frmBestDesign.Tag = algorithm & "; trial=" & bestTrial', 'InStr(frmBestDesign.Tag, algorithm & "; trial=" & bestTrial & ";") = 1')
s=s.replace('        Print #f, algorithm & " winning trial="', '''        Check algorithm & " price equals session best", InStr(frmBestDesign.Tag, "; price=" & bestCost & ";") > 0
        summary = FreeFile: Open bestFolder & "\\run.txt" For Input As #summary
        Do Until EOF(summary)
            Line Input #summary, line
            If Left$(line, 5) = "Stem:" Or Left$(line, 4) = "Toe:" Or Left$(line, 5) = "Heel:" Then
                Check algorithm & " bar label matches winning report", InStr(frmBestDesign.Tag, Split(line, ";")(0)) > 0
            End If
        Loop
        Close #summary
        Print #f, algorithm & " winning trial="''')
s=s.replace('    For height = 3 To 5', '    d.ASst_DB = 101: d.ASst_Sp = 110: d.AStoe_DB = 101: d.AStoe_Sp = 110: d.ASheel_DB = 101: d.ASheel_Sp = 110\n    For height = 3 To 5')
p.write_bytes(s.encode('ascii'))
