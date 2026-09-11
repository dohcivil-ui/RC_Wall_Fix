Attribute VB_Name = "modGraphing"
Option Explicit

' One source for plotting: the selected accept CSV's four original columns.
' history(1) corresponds to CSV No.0; unavailable feasible costs stay unavailable.
Public Sub ReadAcceptCostHistory(ByVal path As String, history() As Double, ByRef bestEvaluation As Long, ByRef initialBestAvailable As Boolean)
    Dim f As Integer, csvLine As String, fields() As String, count As Long, capacity As Long
    Dim best As Double, price As Double, j As Integer, populated As Integer
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    bestEvaluation = 0: initialBestAvailable = False: best = NO_SOLUTION_COST
    capacity = 1024: ReDim history(1 To capacity)
    f = FreeFile
    Open path For Input As #f
    If EOF(f) Then Err.Raise 5, , "Empty accept CSV"
    Line Input #f, csvLine
    If Left$(csvLine, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then csvLine = Mid$(csvLine, 4)
    If csvLine <> "No.,Rejected,Passed,Passed and Better value" Then Err.Raise 5, , "Unexpected accept CSV header"
    Do While Not EOF(f)
        Line Input #f, csvLine
        fields = Split(csvLine, ",")
        If UBound(fields) <> 3 Then Err.Raise 5, , "Expected four CSV columns"
        If Not IsNumeric(fields(0)) Then Err.Raise 5, , "Invalid CSV No."
        If CDbl(fields(0)) <> count Then Err.Raise 5, , "CSV must start at No.0 with consecutive rows"
        count = count + 1
        If count > capacity Then capacity = capacity * 2: ReDim Preserve history(1 To capacity)
        populated = 0
        For j = 1 To 3
            If Len(Trim$(fields(j))) > 0 Then
                populated = populated + 1
                If Not IsNumeric(fields(j)) Then Err.Raise 5, , "Invalid CSV price"
                price = CDbl(fields(j))
                If price <= 0 Or price >= NO_SOLUTION_COST Then Err.Raise 5, , "Unavailable CSV price"
            End If
        Next j
        If populated > 1 Then Err.Raise 5, , "Conflicting CSV classifications"
        If Len(Trim$(fields(3))) > 0 Then
            price = CDbl(fields(3))
            If best < NO_SOLUTION_COST And price > best Then Err.Raise 5, , "Best CSV price increased"
            best = price
            ' CSV values are rounded; a later real improvement may print the same price.
            bestEvaluation = count
            If count = 1 Then initialBestAvailable = True
        End If
        history(count) = best
    Loop
    Close #f: f = 0
    If count = 0 Then Err.Raise 5, , "No CSV observations"
    ReDim Preserve history(1 To count)
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If f > 0 Then Close #f
    On Error GoTo 0
    Err.Raise errorNumber, "ReadAcceptCostHistory", errorText
End Sub

Private Function HasCost(ByVal cost As Double) As Boolean
    HasCost = cost > 0 And cost < NO_SOLUTION_COST
End Function

Private Function Lower(ByVal a As Double, ByVal b As Double) As Double
    Lower = a: If b < a Then Lower = b
End Function

Private Function Higher(ByVal a As Double, ByVal b As Double) As Double
    Higher = a: If b > a Then Higher = b
End Function

Private Function RoundUp(ByVal value As Double, ByVal step As Double) As Double
    RoundUp = -Int(-value / step) * step
End Function

' Legacy quantity argument is ignored: it is not a feasible cost-history point.
Public Sub DrawCostGraph(pic As PictureBox, history() As Double, bestEvaluation As Long, Optional InitialCost As Double = 0, Optional method As String = "")
    Dim emptyHistory(1 To 1) As Double
    If method = "" Then method = "Best"
    Call DrawConvergenceGraph(pic, history, bestEvaluation, method, emptyHistory, 0, "")
End Sub

Public Sub DrawComparisonGraph(pic As PictureBox, hca() As Double, hcaBest As Long, ba() As Double, baBest As Long)
    Call DrawConvergenceGraph(pic, ba, baBest, "BA", hca, hcaBest, "HCA")
End Sub

Private Function MethodColor(ByVal method As String) As Long
    If method = "HCA" Then MethodColor = RGB(223, 84, 0) Else MethodColor = RGB(27, 93, 135)
End Function

Private Sub ScanCosts(history() As Double, ByRef minimum As Double, ByRef maximum As Double)
    Dim i As Long
    For i = LBound(history) To UBound(history)
        If HasCost(history(i)) Then
            minimum = Lower(minimum, history(i)): maximum = Higher(maximum, history(i))
        End If
    Next i
End Sub

' Clip each actual staircase segment to one visible range; never bridge unavailable rows.
Private Sub DrawSteps(pic As PictureBox, history() As Double, ByVal color As Long, ByVal lo As Double, ByVal hi As Double, ByVal xLeft As Double, ByVal width As Double, ByVal bottom As Double, ByVal minimum As Double, ByVal yScale As Double)
    Dim i As Long, n0 As Double, n1 As Double, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    pic.ForeColor = color: pic.DrawStyle = 0: pic.DrawWidth = 2
    For i = LBound(history) To UBound(history) - 1
        n0 = i - 1: n1 = i
        If n1 >= lo And n0 <= hi And HasCost(history(i)) And HasCost(history(i + 1)) Then
            x0 = xLeft + (Higher(n0, lo) - lo) / (hi - lo) * width
            x1 = xLeft + (Lower(n1, hi) - lo) / (hi - lo) * width
            y0 = bottom - (history(i) - minimum) * yScale
            y1 = bottom - (history(i + 1) - minimum) * yScale
            pic.Line (x0, y0)-(x1, y0), color
            If n1 >= lo And n1 <= hi Then pic.Line (x1, y0)-(x1, y1), color
        End If
    Next i
End Sub

Private Sub MarkBest(pic As PictureBox, history() As Double, ByVal bestEvaluation As Long, ByVal method As String, ByVal cut As Double, ByVal xLeft As Double, ByVal width As Double, ByVal bottom As Double, ByVal minimum As Double, ByVal yScale As Double, ByVal labelY As Double, ByVal labelX As Double)
    Dim n As Long, x As Double, y As Double, size As Double, color As Long
    If bestEvaluation < LBound(history) Or bestEvaluation > UBound(history) Then Exit Sub
    If Not HasCost(history(bestEvaluation)) Then Exit Sub
    n = bestEvaluation - 1: If n > cut Then Exit Sub
    color = MethodColor(method): size = width * 0.008
    x = xLeft + n / cut * width: y = bottom - (history(bestEvaluation) - minimum) * yScale
    pic.DrawWidth = 1: pic.DrawStyle = 2
    pic.Line (x, y)-(x, labelY + 300), color
    pic.DrawStyle = 0: pic.FillStyle = 0: pic.FillColor = color
    If method = "HCA" Then
        pic.Line (x - size, y - size)-(x + size, y + size), color, BF
    Else
        pic.Circle (x, y), size, color
    End If
    pic.ForeColor = color: pic.FontSize = 13: pic.FontBold = False
    pic.CurrentX = labelX: pic.CurrentY = labelY
    pic.Print method & ": " & Format$(n, "#,##0")
    pic.FillStyle = 1
End Sub

Private Sub DrawConvergenceGraph(pic As PictureBox, first() As Double, ByVal firstBest As Long, ByVal firstMethod As String, second() As Double, ByVal secondBest As Long, ByVal secondMethod As String)
    Dim minimum As Double, maximum As Double, observedMin As Double, padding As Double
    Dim maxNo As Double, cut As Double, tailStart As Double, lastBest As Double, broken As Boolean
    Dim xLeft As Double, xRight As Double, top As Double, bottom As Double, width As Double, height As Double
    Dim mainWidth As Double, tailLeft As Double, tailWidth As Double, yScale As Double
    Dim i As Long, x As Double, y As Double, value As Double, gridStep As Double, xDivisions As Long
    pic.AutoRedraw = True: pic.Cls: pic.BackColor = vbWhite
    pic.DrawStyle = 0: pic.DrawWidth = 1: pic.FillStyle = 1
    pic.FontName = "Tahoma": pic.Font.Charset = 222
    pic.FontSize = 11: pic.FontBold = False
    minimum = NO_SOLUTION_COST: maximum = 0
    Call ScanCosts(first, minimum, maximum)
    If secondMethod <> "" Then Call ScanCosts(second, minimum, maximum)
    If maximum = 0 Then
        pic.CurrentX = 300: pic.CurrentY = 300: pic.ForeColor = vbBlack
        pic.Print "NO_SOLUTION: no value in Passed and Better value."
        Exit Sub
    End If
    observedMin = minimum
    padding = Higher((maximum - minimum) * 0.06, 100)
    minimum = Higher(0, minimum - padding): maximum = maximum + padding
    If maximum > 1000 Then maximum = RoundUp(maximum, 1000) Else maximum = RoundUp(maximum, 100)
    maxNo = Higher(UBound(first), UBound(second)) - 1
    If maxNo >= 100 Then maxNo = RoundUp(maxNo, 100)
    If maxNo < 1 Then maxNo = 1
    lastBest = Higher(firstBest, secondBest) - 1
    If lastBest <= 100 Then
        cut = 100
    ElseIf lastBest < 500 Then
        cut = 500
    Else
        cut = RoundUp(lastBest + 1, 500)
    End If
    tailStart = maxNo - cut / 5
    broken = cut < tailStart And lastBest < cut
    If Not broken Then cut = maxNo
    xLeft = pic.ScaleWidth * 0.20: xRight = pic.ScaleWidth * 0.97
    top = pic.ScaleHeight * 0.06: bottom = pic.ScaleHeight * 0.82
    width = xRight - xLeft: height = bottom - top
    mainWidth = width: tailWidth = 0
    If broken Then mainWidth = width * 0.79: tailLeft = xLeft + width * 0.85: tailWidth = width * 0.15
    yScale = height / (maximum - minimum)
    gridStep = 10
    If maximum - observedMin > 500 Then gridStep = 100
    If maximum - observedMin > 5000 Then gridStep = 1000
    gridStep = Higher(gridStep, Int(((maximum - observedMin) / 4 + gridStep / 2) / gridStep) * gridStep)
    For i = 0 To 4
        value = maximum - (4 - i) * gridStep
        If i = 0 Then value = observedMin
        y = bottom - (value - minimum) * yScale
        pic.DrawStyle = 2: pic.DrawWidth = 1
        pic.Line (xLeft, y)-(xLeft + mainWidth, y), RGB(215, 219, 223)
        If broken Then pic.Line (tailLeft, y)-(xRight, y), RGB(215, 219, 223)
        pic.DrawStyle = 0: pic.ForeColor = vbBlack
        pic.CurrentX = 100: pic.CurrentY = y - 100
        If i = 0 Then pic.Print Format$(value, "#,##0.00") Else pic.Print Format$(value, "#,##0")
    Next i
    xDivisions = 4
    If cut = 100 Or cut = 500 Or cut = 1500 Then xDivisions = 5
    For i = 0 To xDivisions
        x = xLeft + mainWidth * i / xDivisions
        pic.DrawStyle = 2
        pic.Line (x, top)-(x, bottom), RGB(215, 219, 223)
        pic.DrawStyle = 0: pic.CurrentX = x - 120: pic.CurrentY = bottom + 100
        pic.Print Format$(cut * i / xDivisions, "#,##0")
    Next i
    pic.DrawWidth = 2
    pic.Line (xLeft, top)-(xLeft, bottom), vbBlack
    pic.Line (xLeft, bottom)-(xLeft + mainWidth, bottom), vbBlack
    If broken Then
        pic.Line (tailLeft, bottom)-(xRight, bottom), vbBlack
        pic.CurrentX = xRight - 500: pic.CurrentY = bottom + 100: pic.Print Format$(maxNo, "#,##0")
        pic.ForeColor = RGB(90, 90, 90): pic.FontSize = 11
        pic.CurrentX = xLeft + mainWidth * 0.58: pic.CurrentY = top + height * 0.28
        pic.Print "ละช่วง " & Format$(cut, "#,##0") & "-" & Format$(tailStart, "#,##0") & " รอบ"
    End If
    Call DrawSteps(pic, first, MethodColor(firstMethod), 0, cut, xLeft, mainWidth, bottom, minimum, yScale)
    If broken Then Call DrawSteps(pic, first, MethodColor(firstMethod), tailStart, maxNo, tailLeft, tailWidth, bottom, minimum, yScale)
    If secondMethod <> "" Then
        Call DrawSteps(pic, second, MethodColor(secondMethod), 0, cut, xLeft, mainWidth, bottom, minimum, yScale)
        If broken Then Call DrawSteps(pic, second, MethodColor(secondMethod), tailStart, maxNo, tailLeft, tailWidth, bottom, minimum, yScale)
    End If
    If broken Then
        For i = 0 To 1
            x = xLeft + mainWidth: If i = 1 Then x = tailLeft
            For gridStep = 0 To 1
                y = bottom: If gridStep = 1 Then y = bottom - (observedMin - minimum) * yScale
                pic.DrawWidth = 6: pic.Line (x - 75, y + 70)-(x + 75, y - 70), vbWhite
                pic.DrawWidth = 2: pic.Line (x - 75, y + 70)-(x + 75, y - 70), RGB(50, 50, 50)
            Next gridStep
        Next i
    End If
    Call MarkBest(pic, first, firstBest, firstMethod, cut, xLeft, mainWidth, bottom, minimum, yScale, top + height * 0.56, xLeft + mainWidth * 0.20)
    If secondMethod <> "" Then Call MarkBest(pic, second, secondBest, secondMethod, cut, xLeft, mainWidth, bottom, minimum, yScale, top + height * 0.56, xLeft + mainWidth * 0.60)
    pic.ForeColor = vbBlack: pic.FontSize = 12: pic.FontBold = False
    pic.CurrentX = 100: pic.CurrentY = 50: pic.Print "ราคารวม (บาท/ม.)"
    pic.CurrentX = xLeft + width * 0.40: pic.CurrentY = bottom + 450: pic.Print "จำนวนรอบ (No.)"
    pic.CurrentX = xLeft + mainWidth * 0.60: pic.CurrentY = top + 100
    pic.ForeColor = MethodColor(firstMethod): pic.Print firstMethod
    If secondMethod <> "" Then
        pic.CurrentX = xLeft + mainWidth * 0.82: pic.CurrentY = top + 100
        pic.ForeColor = MethodColor(secondMethod): pic.Print secondMethod
    End If
    pic.DrawStyle = 0: pic.DrawWidth = 1: pic.FillStyle = 1: pic.Refresh
End Sub

Public Sub ClearGraph(pic As PictureBox)
    pic.Cls: pic.Refresh
End Sub
