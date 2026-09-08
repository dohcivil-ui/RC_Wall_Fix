from pathlib import Path
import re
root=Path(__file__).resolve().parent.parent
def read(n): return (root/n).read_bytes().decode('latin1').replace('\r\n','\n')
def write(n,s): (root/n).write_bytes(s.replace('\n','\r\n').encode('latin1'))
s=read('modShared.bas')
s=s.replace('Public LastPrice As Double','''Public EvaluationCount As Long
Public EvaluationBudget As Long
Public RunSeed As Long
Public RunBest As Design
Public RunBestCost As Double
Public RunBestEvaluation As Long
Public RunStatus As String
Public RunFolder As String
Public RunRecoveryCount As Long
Private EvaluationCSV As String
Public LastPrice As Double''')
s+='''
' One evaluation = one complete candidate request (including invalid geometry).
' Initial, reset and neighbor requests all enter here; reporting does not count.
Public Sub BeginSearch(budget As Long, seed As Long, algorithm As String)
    Dim dummy As Single, emptyDesign As Design, suffix As Long, basePath As String
    If budget < 1 Then Err.Raise 5, , "Evaluation budget must be positive"
    EvaluationCount = 0: EvaluationBudget = budget: RunSeed = seed
    RunBest = emptyDesign: RunBestCost = NO_SOLUTION_COST: RunBestEvaluation = 0
    RunRecoveryCount = 0
    RunStatus = "NO_SOLUTION"
    dummy = Rnd(-1): Randomize seed
    If Dir$(App.Path & "\\results", vbDirectory) = "" Then MkDir App.Path & "\\results"
    basePath = App.Path & "\\results\\" & algorithm & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-seed" & CStr(seed)
    RunFolder = basePath
    Do While Dir$(RunFolder, vbDirectory) <> ""
        suffix = suffix + 1: RunFolder = basePath & "-" & CStr(suffix)
    Loop
    MkDir RunFolder
    EvaluationCSV = "evaluation,entry,valid,reason,cost,best_cost,tt,tb,TBase,Base,LToe,StemDB,StemSP,ToeDB,ToeSP,HeelDB,HeelSP" & vbCrLf
End Sub

Public Function EvaluateCandidate(d As Design, entry As String, ByRef candidateCost As Double) As Boolean
    Dim ot As Double, sl As Double, bc As Double, ok As Boolean
    If EvaluationCount >= EvaluationBudget Then Err.Raise 5, , "Evaluation budget exhausted"
    EvaluationCount = EvaluationCount + 1
    candidateCost = NO_SOLUTION_COST
    ok = CheckDesignValid(d, d.ASst_DB, d.ASst_Sp, d.AStoe_DB, d.AStoe_Sp, d.ASheel_DB, d.ASheel_Sp, ot, sl, bc)
    If GeometryOK(d) Then candidateCost = CalculateCost(d)
    d.TotalCost = candidateCost: d.IsValid = ok
    If ok Then
        If Not RunBest.IsValid Or candidateCost < RunBestCost Then
            RunBest = d: RunBestCost = candidateCost: RunBestEvaluation = EvaluationCount
            RunStatus = "SOLUTION_FOUND_CONFIGURED_CHECKS"
        End If
    End If
    EvaluationCSV = EvaluationCSV & EvaluationCount & "," & entry & "," & CStr(ok) & "," & Replace(LastValidationReason, ",", ";") & "," & _
        CsvNumber(candidateCost) & "," & CsvNumber(RunBestCost) & "," & CsvNumber(d.tt) & "," & CsvNumber(d.tb) & "," & _
        CsvNumber(d.TBase) & "," & CsvNumber(d.Base) & "," & CsvNumber(d.LToe) & "," & _
        d.ASst_DB & "," & d.ASst_Sp & "," & d.AStoe_DB & "," & d.AStoe_Sp & "," & d.ASheel_DB & "," & d.ASheel_Sp & vbCrLf
    EvaluateCandidate = ok
End Function

Public Function CsvNumber(value As Double) As String
    CsvNumber = Replace(Format$(value, "0.0000000000"), ",", ".")
End Function

Public Sub FinishSearch()
    Dim f As Integer
    modDataStructures.BestCostIteration = RunBestEvaluation
    If Not RunBest.IsValid Then RunBest.TotalCost = NO_SOLUTION_COST
    f = FreeFile
    Open RunFolder & "\\evaluations.csv" For Output As #f
    Print #f, EvaluationCSV;
    Close #f
    f = FreeFile
    Open RunFolder & "\\run.txt" For Output As #f
    Print #f, "Status=" & RunStatus
    Print #f, "Seed=" & RunSeed & "; Evaluations=" & EvaluationCount & "; Budget=" & EvaluationBudget
    Print #f, "BestEvaluation=" & RunBestEvaluation & "; Recoveries=" & RunRecoveryCount
    Print #f, "H=" & H & "; H1=" & H1 & "; gamma_soil=" & gamma_soil & "; gamma_concrete=" & gamma_concrete
    Print #f, "phi=" & phi & "; mu=" & mu & "; qa_allowable=" & qa & "; clear_cover=" & cover
    Print #f, "fc_prime=" & currentMaterial.fc & "; fy=" & currentMaterial.fy & "; passive_fraction=" & PassiveFactor
    Print #f, "WSDReviewed=" & WSDReviewed & "; source=" & WSDSource
    Print #f, "AllowableShear=" & AllowableShear & "; MinStemRatio=" & MinStemRatio & "; MinBaseRatio=" & MinBaseRatio
    Print #f, FormatResults(RunBest, currentMaterial)
    Close #f
End Sub
'''
write('modShared.bas',s)
for name,algo,func,init,get,gen,hist in [
    ('modBA.bas','BA','BisectionOptimization','InitializeCurrentDesign_BA','GetDesignFromCurrent_BA','GenerateNeighbor_BA','CostHistory_BA'),
    ('modHillClimbing.bas','HCA','HillClimbingOptimization','InitializeCurrentDesign','GetDesignFromCurrent','GenerateNeighbor','modDataStructures.CostHistory')]:
    s=read(name)
    match=re.search(rf'Public Function {func}\([\s\S]*?\nEnd Function',s)
    signature=match.group().split(' As Design',1)[0]+' As Design'
    signature=signature.replace('Optional sharedHeelSP As Integer = 0)', 'Optional sharedHeelSP As Integer = 0, _\n                       Optional RandomSeed As Long = 12345)')
    body='''
    Dim current As Design, neighbor As Design
    Dim currentCost As Double, neighborCost As Double, currentValid As Boolean, ok As Boolean
    Dim saved(1 To 11) As Integer
    Dim Newtt As Integer, Newtb As Integer, NewTBase As Integer, NewBase As Integer, NewLToe As Integer
    Dim NewStemDB As Integer, NewStemSP As Integer, NewToeDB As Integer, NewToeSP As Integer
    Dim NewHeelDB As Integer, NewHeelSP As Integer, i As Long, j As Integer
'''
    if algo=='BA': body+='''    Dim rootMintb As Integer, rootMaxtb As Integer, rootMinTBase As Integer, rootMaxTBase As Integer
    Dim rootMinBase As Integer, rootMaxBase As Integer, innerIterations As Long, roundValid As Boolean
'''
    body+='''    modShared.H = wall_height: modShared.H1 = backfill_height
    modShared.gamma_soil = soil_gamma: modShared.gamma_concrete = concrete_gamma
    modShared.phi = friction_angle: modShared.mu = friction_coef
    modShared.qa = allowable_bearing: modShared.cover = concrete_cover
    modShared.currentMaterial = material
    modShared.currentWSD = CalculateWSDParameters(material.fy, material.fc)
    Call InitializeArrays
'''
    body+=f'    ReDim {hist}(1 To MaxIterations)\n    Call BeginSearch(MaxIterations, RandomSeed, "{algo}")\n    Call {init}\n'
    fields=['tt','tb','TBase','Base','LToe','StemDB','StemSP','ToeDB','ToeSP','HeelDB','HeelSP']
    body+='    If useSharedInit Then\n'
    # Invalid array indices are explicit input errors, never coerced to another design.
    for field,low,high in zip(fields,['TT_MIN','TB_MIN','TBASE_MIN','BASE_MIN','LTOE_MIN','DB_MIN','SP_MIN','DB_MIN','SP_MIN','DB_MIN','SP_MIN'],['TT_MAX','tb_max','TBase_max','BASE_MAX','LTOE_MAX','DB_MAX','SP_MAX','DB_MAX','SP_MAX','DB_MAX','SP_MAX']):
        body+=f'        If shared{field} < {low} Or shared{field} > {high} Then Err.Raise 5, , "Invalid shared {field} index"\n'
        body+=f'        Current{field} = shared{field}\n'
    body+='    End If\n'
    if algo=='BA': body+='''    rootMintb = Mintb: rootMaxtb = Maxtb
    rootMinTBase = MinTBase: rootMaxTBase = MaxTBase
    rootMinBase = MinBase: rootMaxBase = MaxBase
    Countloop = 0
'''
    body+=f'''    current = {get}()
    currentValid = EvaluateCandidate(current, "initial", currentCost)
    If Not currentValid Then currentCost = NO_SOLUTION_COST
    {hist}(EvaluationCount) = RunBestCost
'''
    body+='    Do While EvaluationCount < MaxIterations\n'
    if algo=='BA': body+=f'''        Countloop = Countloop + 1
        If Countloop > 1 Then
            Currenttb = Midtb: CurrentTBase = MidTBase: CurrentBase = MidBase
            If WP_tt(Currenttt) > WP_tb(Currenttb) Then
                For j = TT_MAX To TT_MIN Step -1
                    If WP_tt(j) <= WP_tb(Currenttb) Then Currenttt = j: Exit For
                Next j
            End If
            current = {get}()
            currentValid = EvaluateCandidate(current, "reset", currentCost)
            If Not currentValid Then currentCost = NO_SOLUTION_COST
            {hist}(EvaluationCount) = RunBestCost
        End If
        roundValid = currentValid
        innerIterations = 20 * Countloop
        For i = 1 To innerIterations
            If EvaluationCount >= MaxIterations Then Exit For
'''
    for index,field in enumerate(fields,1): body+=f'        saved({index}) = Current{field}\n'
    body+=f'''        Call {gen}(Newtt, Newtb, NewTBase, NewBase, NewLToe, NewStemDB, NewStemSP, NewToeDB, NewToeSP, NewHeelDB, NewHeelSP)
'''
    for field in fields: body+=f'        Current{field} = New{field}\n'
    body+=f'''        neighbor = {get}()
        ok = EvaluateCandidate(neighbor, "neighbor", neighborCost)
        If ok And (Not currentValid Or neighborCost < currentCost) Then
            current = neighbor: currentCost = neighborCost: currentValid = True
        Else
'''
    for index,field in enumerate(fields,1): body+=f'            Current{field} = saved({index})\n'
    body+=f'''        End If
        {hist}(EvaluationCount) = RunBestCost
        DoEvents
'''
    if algo=='BA': body+='''        If ok Then roundValid = True
        Next i
        ' Preserve the original THREE-variable heuristic, not a monotone proof.
        ' Invalid rounds supply no evidence for removing any part of the domain.
        If Not roundValid Or (Mintb = Maxtb And MinTBase = MaxTBase And MinBase = MaxBase) Then
            Mintb = rootMintb: Maxtb = rootMaxtb
            MinTBase = rootMinTBase: MaxTBase = rootMaxTBase
            MinBase = rootMinBase: MaxBase = rootMaxBase
            MidPricetb = NO_SOLUTION_COST: MidPriceTBase = NO_SOLUTION_COST: MidPrice = NO_SOLUTION_COST
            Midtb = Rand(Mintb, Maxtb): MidTBase = Rand(MinTBase, MaxTBase): MidBase = Rand(MinBase, MaxBase)
            RunRecoveryCount = RunRecoveryCount + 1
        Else
            If currentCost < MidPricetb Then
                Maxtb = Currenttb: MidPricetb = currentCost
            Else
                Mintb = Currenttb
            End If
            If currentCost < MidPriceTBase Then
                MaxTBase = CurrentTBase: MidPriceTBase = currentCost
            Else
                MinTBase = CurrentTBase
            End If
            If currentCost < MidPrice Then
                MaxBase = CurrentBase: MidPrice = currentCost
            Else
                MinBase = CurrentBase
            End If
            Midtb = (Mintb + Maxtb) \\ 2
            MidTBase = (MinTBase + MaxTBase) \\ 2
            MidBase = (MinBase + MaxBase) \\ 2
        End If
'''
    body+=f'''    Loop
    Call FinishSearch
    {func} = RunBest
End Function'''
    if algo=='BA': body=body.replace('    Countloop = 0','    MidPricetb = NO_SOLUTION_COST: MidPriceTBase = NO_SOLUTION_COST: MidPrice = NO_SOLUTION_COST\n    Countloop = 0')
    s=s[:match.start()]+signature+'\n'+body+s[match.end():]
    # Existing public exporters kept for Form1 compatibility; every save is unique.
    s=re.sub(r'filePath = "D:\\([^"\n]+)" & Format\(wallHeight, "0"\) & "\.csv"',r'filePath = UniqueExportPath("\1" & Format(wallHeight, "0"))',s)
    write(name,s)
s=read('modShared.bas')+'''
Public Function UniqueExportPath(stem As String) As String
    Dim n As Long, p As String
    p = RunFolder & "\\" & stem & ".csv"
    Do While Len(Dir$(p)) > 0
        n = n + 1: p = RunFolder & "\\" & stem & "-" & n & ".csv"
    Loop
    UniqueExportPath = p
End Function
'''
write('modShared.bas',s)
