VERSION 5.00
Begin VB.Form frmBestDesign
   BorderStyle     =   3
   Caption         =   "Best retaining wall - dimensions"
   ClientHeight    =   10500
   ClientWidth     =   13500
   MaxButton       =   0
   MinButton       =   0
   ShowInTaskbar   =   0
   StartUpPosition =   1
   Begin VB.PictureBox picSketch
      AutoRedraw      =   -1
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0
      Height          =   9600
      Left            =   180
      ScaleMode       =   3
      TabStop         =   0
      Top             =   180
      Width           =   13140
   End
   Begin VB.CommandButton cmdClose
      Cancel          =   -1
      Caption         =   "OK"
      Default         =   -1
      Height          =   420
      Left            =   11760
      TabIndex        =   0
      Top             =   9960
      Width           =   1560
   End
End
Attribute VB_Name = "frmBestDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Display only: receives the completed session best, never searches or changes it.
Friend Sub ShowBest(ByRef d As Design, ByVal wallHeight As Double, ByVal frontHeight As Double, ByVal algorithm As String, ByVal trial As Long, Optional ByVal bestCost As Double = 0, Optional ByVal clearCover As Double = 0.075, Optional ByVal totalTrials As Long = 1, Optional ByVal bestEvaluation As Long = 0)
    Load Me
    Me.Tag = algorithm & "; trial=" & trial & "; trials=" & totalTrials & "; evaluation=" & bestEvaluation
    picSketch.Cls
    picSketch.Tag = "NO_SOLUTION"
    picSketch.Font.Name = "Tahoma"
    picSketch.ForeColor = RGB(35, 39, 40)
    picSketch.Font.Size = 15
    picSketch.Font.Bold = True
    If trial > 0 Then
        CenterText picSketch.ScaleWidth / 2, 14, "FINAL RESULT  |  " & algorithm & "  |  Selected trial " & trial & " / " & totalTrials
    Else
        CenterText picSketch.ScaleWidth / 2, 14, "FINAL RESULT  |  " & algorithm & "  |  " & totalTrials & " trials complete"
    End If
    picSketch.Font.Size = 10
    picSketch.Font.Bold = False
    If d.IsValid And ValidGeometry(d, wallHeight, frontHeight) Then
        If ValidBars(d) Then
            picSketch.Font.Bold = True
            If bestCost > 0 Then
                CenterText picSketch.ScaleWidth / 2, 44, "COST  " & Format$(bestCost, "#,##0.00") & " Baht/m  (concrete + main steel)"
            Else
                CenterText picSketch.ScaleWidth / 2, 44, "Drawing fixture - price not supplied"
            End If
            picSketch.Font.Bold = False
            Me.Tag = Me.Tag & "; price=" & bestCost & "; Stem: " & RebarText(d.ASst_DB, d.ASst_Sp) & "; Toe: " & RebarText(d.AStoe_DB, d.AStoe_Sp) & "; Heel: " & RebarText(d.ASheel_DB, d.ASheel_Sp)
        End If
        If bestEvaluation > 0 Then
            CenterText picSketch.ScaleWidth / 2, 69, "Best cost first found at evaluation " & bestEvaluation & "  |  Dimensions in metres"
        Else
            CenterText picSketch.ScaleWidth / 2, 69, "Dimensions in metres  |  Same scale in both directions"
        End If
        DrawWall d, wallHeight, frontHeight, clearCover
        picSketch.Tag = "H=" & wallHeight & "; H1=" & frontHeight & "; tt=" & d.tt & "; tb=" & d.tb & "; TBase=" & d.TBase & "; B=" & d.Base & "; toe=" & d.LToe & "; heel=" & d.LHeel
    Else
        CenterText picSketch.ScaleWidth / 2, 150, "NO_SOLUTION - no valid best geometry to display."
    End If
End Sub

Private Function ValidGeometry(d As Design, wallHeight As Double, frontHeight As Double) As Boolean
    If d.tt <= 0 Or d.tb < d.tt Or d.TBase <= 0 Or d.Base <= 0 Or d.LToe <= 0 Or d.LHeel <= 0 Then Exit Function
    If wallHeight <= d.TBase Or frontHeight < d.TBase Or frontHeight > wallHeight Then Exit Function
    ValidGeometry = Abs(d.Base - d.LToe - d.tb - d.LHeel) < 0.000001
End Function

Private Sub DrawWall(d As Design, wallHeight As Double, frontHeight As Double, clearCover As Double)
    Dim pixelsPerMetre As Double, x0 As Double, x1 As Double, toe As Double, back As Double, top As Double
    Dim y0 As Double, yTop As Double, yBase As Double, yFront As Double
    pixelsPerMetre = (picSketch.ScaleHeight - 275) / wallHeight
    If pixelsPerMetre * d.Base > picSketch.ScaleWidth - 440 Then pixelsPerMetre = (picSketch.ScaleWidth - 440) / d.Base
    x0 = (picSketch.ScaleWidth - d.Base * pixelsPerMetre) / 2: x1 = x0 + d.Base * pixelsPerMetre
    toe = x0 + d.LToe * pixelsPerMetre: back = toe + d.tb * pixelsPerMetre: top = back - d.tt * pixelsPerMetre
    y0 = picSketch.ScaleHeight - 135: yTop = y0 - wallHeight * pixelsPerMetre
    yBase = y0 - d.TBase * pixelsPerMetre: yFront = y0 - frontHeight * pixelsPerMetre
    DrawSection x0, x1, toe, back, top, y0, yTop, yBase, yFront
    If ValidBars(d) Then DrawMainBars d, pixelsPerMetre, clearCover, x0, x1, toe, back, y0, yTop, yBase
    DrawHorizontal top, back, yTop, yTop - 22, "tt = " & Format$(d.tt, "0.00")
    DrawHorizontal x0, toe, y0, y0 + 28, "LToe = " & Format$(d.LToe, "0.00")
    DrawHorizontal toe, back, y0, y0 + 60, "tb = " & Format$(d.tb, "0.00")
    DrawHorizontal back, x1, y0, y0 + 28, "LHeel = " & Format$(d.LHeel, "0.00")
    DrawHorizontal x0, x1, y0, y0 + 92, "Base = " & Format$(d.Base, "0.00")
    DrawVertical yTop, y0, x0, x0 - 140, "H = " & Format$(wallHeight, "0.00"), False
    DrawVertical yFront, y0, x0 - 30, x0 - 48, "H1 = " & Format$(frontHeight, "0.00"), False
    DrawVertical yBase, y0, x1, x1 + 26, "TBase = " & Format$(d.TBase, "0.00"), True
End Sub

Private Sub DrawSection(x0 As Double, x1 As Double, toe As Double, back As Double, top As Double, y0 As Double, yTop As Double, yBase As Double, yFront As Double)
    Dim edge As Long
    edge = RGB(35, 35, 35)
    picSketch.DrawWidth = 2
    picSketch.Line (x0 - 30, yFront)-(top + (toe - top) * (yFront - yTop) / (yBase - yTop), yFront), edge
    picSketch.Line (back, yTop)-(x1 + 20, yTop), edge
    picSketch.Line (x0, y0)-(x0, yBase), edge
    picSketch.Line -(toe, yBase), edge
    picSketch.Line -(top, yTop), edge
    picSketch.Line -(back, yTop), edge
    picSketch.Line -(back, yBase), edge
    picSketch.Line -(x1, yBase), edge
    picSketch.Line -(x1, y0), edge
    picSketch.Line -(x0, y0), edge
    picSketch.DrawWidth = 1
End Sub

Private Function ValidBars(d As Design) As Boolean
    ValidBars = d.ASst_DB >= DB_MIN And d.ASst_DB <= DB_MAX And d.ASst_Sp >= SP_MIN And d.ASst_Sp <= SP_MAX And d.AStoe_DB >= DB_MIN And d.AStoe_DB <= DB_MAX And d.AStoe_Sp >= SP_MIN And d.AStoe_Sp <= SP_MAX And d.ASheel_DB >= DB_MIN And d.ASheel_DB <= DB_MAX And d.ASheel_Sp >= SP_MIN And d.ASheel_Sp <= SP_MAX
End Function

Private Function RebarText(DB As Integer, SP As Integer) As String
    RebarText = "DB" & WP_DB(DB) & " @ " & Format$(WP_SP(SP), "0.00") & " m"
End Function

Private Sub DrawMainBars(d As Design, pixelsPerMetre As Double, clearCover As Double, x0 As Double, x1 As Double, toe As Double, back As Double, y0 As Double, yTop As Double, yBase As Double)
    Dim stemX As Double, toeY As Double, heelY As Double, endCover As Double
    endCover = clearCover * pixelsPerMetre
    stemX = back - (clearCover + WP_DB(d.ASst_DB) / 2000#) * pixelsPerMetre
    toeY = y0 - (clearCover + WP_DB(d.AStoe_DB) / 2000#) * pixelsPerMetre
    heelY = yBase + (clearCover + WP_DB(d.ASheel_DB) / 2000#) * pixelsPerMetre
    ' Dot symbols follow the requested schematic; their count is not a bar takeoff.
    DrawBarDots stemX, yTop + endCover, stemX, yBase - 4
    DrawBarDots x0 + endCover, toeY, toe - 4, toeY
    DrawBarDots back + 4, heelY, x1 - endCover, heelY
    RebarCallout x1 + 65, yTop + 75, "Stem - back face", RebarText(d.ASst_DB, d.ASst_Sp), stemX, (yTop + endCover + yBase - 4) / 2, False
    RebarCallout x0 - 110, yBase - 130, "Toe - bottom", RebarText(d.AStoe_DB, d.AStoe_Sp), (x0 + endCover + toe - 4) / 2, toeY, True
    RebarCallout x1 + 65, yBase - 80, "Heel - top", RebarText(d.ASheel_DB, d.ASheel_Sp), (back + 4 + x1 - endCover) / 2, heelY, False
End Sub

Private Sub DrawBarDots(xStart As Double, yStart As Double, xEnd As Double, yEnd As Double)
    Dim length As Double, intervals As Long, i As Long
    length = Sqr((xEnd - xStart) ^ 2 + (yEnd - yStart) ^ 2)
    If length <= 0 Then Exit Sub
    intervals = 2 * CLng(length / 44): If intervals < 2 Then intervals = 2
    picSketch.DrawWidth = 1
    picSketch.FillStyle = 0: picSketch.FillColor = vbBlack
    For i = 0 To intervals
        picSketch.Circle (xStart + (xEnd - xStart) * i / intervals, yStart + (yEnd - yStart) * i / intervals), 3, vbBlack
    Next i
    picSketch.FillStyle = 1
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

Private Sub DrawHorizontal(left As Double, right As Double, face As Double, atY As Double, text As String)
    Dim tip As Double
    tip = 5: If (right - left) / 3 < tip Then tip = (right - left) / 3
    picSketch.Line (left, face)-(left, atY + 5), RGB(110, 110, 100)
    picSketch.Line (right, face)-(right, atY + 5), RGB(110, 110, 100)
    picSketch.Line (left, atY)-(right, atY)
    picSketch.Line (left + tip, atY - 3)-(left, atY)
    picSketch.Line -(left + tip, atY + 3)
    picSketch.Line (right - tip, atY - 3)-(right, atY)
    picSketch.Line -(right - tip, atY + 3)
    CenterText (left + right) / 2, atY - 19, text
End Sub

Private Sub DrawVertical(top As Double, bottom As Double, face As Double, atX As Double, text As String, rightSide As Boolean)
    Dim x As Double, tip As Double
    tip = 5: If (bottom - top) / 3 < tip Then tip = (bottom - top) / 3
    picSketch.Line (face, top)-(atX, top), RGB(110, 110, 100)
    picSketch.Line (face, bottom)-(atX, bottom), RGB(110, 110, 100)
    picSketch.Line (atX, top)-(atX, bottom)
    picSketch.Line (atX - 3, top + tip)-(atX, top)
    picSketch.Line -(atX + 3, top + tip)
    picSketch.Line (atX - 3, bottom - tip)-(atX, bottom)
    picSketch.Line -(atX + 3, bottom - tip)
    x = atX - picSketch.TextWidth(text) - 8: If rightSide Then x = atX + 8
    picSketch.CurrentX = x: picSketch.CurrentY = (top + bottom - picSketch.TextHeight(text)) / 2
    picSketch.Print text
End Sub

Private Sub CenterText(x As Double, y As Double, text As String)
    picSketch.CurrentX = x - picSketch.TextWidth(text) / 2
    picSketch.CurrentY = y
    picSketch.Print text
End Sub

Private Sub Form_Load()
    If Width > Screen.Width * 0.95 Then Width = Screen.Width * 0.95
    If Height > Screen.Height * 0.9 Then Height = Screen.Height * 0.9
    picSketch.Move 180, 180, ScaleWidth - 360, ScaleHeight - 900
    cmdClose.Move ScaleWidth - cmdClose.Width - 180, ScaleHeight - 540
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
