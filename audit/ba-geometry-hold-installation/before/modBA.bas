Attribute VB_Name = "modBA"
Option Explicit

' User-requested control: BA runs the exact HCA engine, without bisection.
' Keep the public entry point and BA-labelled exports/history for the existing UI.
Private csvAcceptData As String
Private csvLoopData As String
Private loopCount As Long
Private bestIterationInRun As Long
Public CostHistory_BA() As Double

Public Function GetConcretePrice_BA(fc As Integer) As Double
    GetConcretePrice_BA = modShared.GetConcretePrice(fc)
End Function

'--------------------------------------------------------------------------------
' Get SD40 Material Properties
'--------------------------------------------------------------------------------
Public Function GetSD40Material_BA(fc As Integer, _
                                   concPrice As Double, _
                                   steelPrice As Double) As MaterialProperties
    GetSD40Material_BA = modShared.GetSD40Material(fc, concPrice, steelPrice)
End Function

Public Function BisectionOptimization(MaxIterations As Long, _
                       wall_height As Double, _
                       backfill_height As Double, _
                       soil_gamma As Double, _
                       concrete_gamma As Double, _
                       friction_angle As Double, _
                       friction_coef As Double, _
                       allowable_bearing As Double, _
                       concrete_cover As Double, _
                       material As MaterialProperties, _
                       Optional useSharedInit As Boolean = False, _
                       Optional sharedTt As Integer = 0, _
                       Optional sharedTb As Integer = 0, _
                       Optional sharedTBase As Integer = 0, _
                       Optional sharedBase As Integer = 0, _
                       Optional sharedLToe As Integer = 0, _
                       Optional sharedStemDB As Integer = 0, _
                       Optional sharedStemSP As Integer = 0, _
                       Optional sharedToeDB As Integer = 0, _
                       Optional sharedToeSP As Integer = 0, _
                       Optional sharedHeelDB As Integer = 0, _
                       Optional sharedHeelSP As Integer = 0, _
                       Optional TrialNumber As Long = 1) As Design
    Dim best As Design, i As Long
    best = modHillClimbing.HillClimbingOptimization( _
        MaxIterations, wall_height, backfill_height, soil_gamma, concrete_gamma, _
        friction_angle, friction_coef, allowable_bearing, concrete_cover, material, _
        useSharedInit, sharedTt, sharedTb, sharedTBase, sharedBase, sharedLToe, _
        sharedStemDB, sharedStemSP, sharedToeDB, sharedToeSP, sharedHeelDB, sharedHeelSP, _
        TrialNumber, True)
    ReDim CostHistory_BA(1 To MaxIterations)
    For i = 1 To MaxIterations
        CostHistory_BA(i) = modDataStructures.CostHistory(i)
    Next i
    BisectionOptimization = best
End Function

Public Sub InitCSVExport_BA()
    csvAcceptData = "No.,Rejected,Passed,Passed and Better value" & vbCrLf
    bestIterationInRun = 0
End Sub

Public Sub InitLoopCounter_BA()
    csvLoopData = "No.,Loop,BestPrice" & vbCrLf
    LastLoopCSVPath = ""
    loopCount = 0
End Sub

'--------------------------------------------------------------------------------
' v3.1: LogIteration_BA - ???????????? HCA
' - Rejected: ??????? column 2 (Invalid)
' - Passed: ??????? column 3 (Valid ???????????? best)
' - Passed and Better value: ??????? column 4 (Valid ????????? best)
'--------------------------------------------------------------------------------
Public Sub LogIteration_BA(iteration As Long, cost As String, IsValid As Boolean, isBetter As Boolean)
    If Not IsValid Then
        ' Rejected - ??????? column 2
        csvAcceptData = csvAcceptData & iteration & "," & cost & ",," & vbCrLf
    ElseIf isBetter Then
        ' Passed and Better value - ??????? column 4
        csvAcceptData = csvAcceptData & iteration & ",,," & cost & vbCrLf
        bestIterationInRun = iteration
    Else
        ' Passed - ??????? column 3
        csvAcceptData = csvAcceptData & iteration & ",," & cost & "," & vbCrLf
    End If
End Sub

'--------------------------------------------------------------------------------
' v3.1: LogLoopResult_BA - ???????????? HCA
' ??????: No., Loop (iteration ????? best), BestPrice
'--------------------------------------------------------------------------------
Public Sub LogLoopResult_BA(bestPrice As Double)
    Dim price As String
    loopCount = loopCount + 1
    If RunBest.IsValid Then price = CsvPrice(bestPrice)
    csvLoopData = csvLoopData & loopCount & "," & bestIterationInRun & "," & price & vbCrLf
End Sub

Public Sub SaveAcceptCSV_BA(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-BA-H" & Replace$(CStr(wallHeight), ",", "."), csvAcceptData, True)
End Sub

Public Sub SaveLoopPriceCSV_BA(wallHeight As Double)
    LastAcceptCSVPath = WriteExportCSV("accept-BA-H" & Replace$(CStr(wallHeight), ",", "."), csvAcceptData, True)
    LastLoopCSVPath = WriteExportCSV("loopPrice-BA-H" & Replace$(CStr(wallHeight), ",", "."), csvLoopData, True)
End Sub
