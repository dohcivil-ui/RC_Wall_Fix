Attribute VB_Name = "HcaProbe"
Option Explicit
Private f As Integer
Public Sub Main()
    Dim mat As MaterialProperties, d As Design, trial As Design, r As ProjectDetail
    Dim db As Integer, sp As Integer, part As Integer, valid As Boolean, best As Double, cost As Double
    On Error GoTo Failed
    Call InitializeArrays
    Call EnableProjectChecks
    H = 5: H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4: phi = 30: mu = 0.6: qa = 30: cover = 0.075
    mat = GetSD40Material(320, GetConcretePrice(320), 24): currentMaterial = mat
    currentWSD = CalculateWSDParameters(mat.fy, mat.fc)
    f = FreeFile: Open App.Path & "\probe.txt" For Output As #f
    ' User HCA trial 21, seed 12365: the best recorded design in those 30 trials.
    d.tt = 0.2: d.tb = 0.4: d.TBase = 0.3: d.Base = 3: d.LToe = 1: d.LHeel = 1.6
    d.ASst_DB = 101: d.ASst_Sp = 110: d.AStoe_DB = 100: d.AStoe_Sp = 110: d.ASheel_DB = 101: d.ASheel_Sp = 110
    Print #f, "INPUT H=" & H & "; H1=" & H1 & "; soil=" & gamma_soil & "; conc=" & gamma_concrete & "; phi=" & phi & "; mu=" & mu & "; qa=" & qa & "; cover=" & cover & "; passive=" & PassiveFactor & "; geometry=" & GeometryOK(d) & "; heel=" & CheckHeelLayout(d) & "; fc=" & currentWSD.fcPrime & "; fy=" & currentWSD.fy
    valid = CheckProjectDesign(d, r): best = r.Cost
    Print #f, "USER_HCA valid=" & valid & "; cost=" & best & "; reason=" & LastValidationReason
    For part = 0 To 2
        For db = DB_MIN To DB_MAX
            For sp = SP_MIN To SP_MAX
                trial = d
                Select Case part
                    Case 0: trial.ASst_DB = db: trial.ASst_Sp = sp
                    Case 1: trial.AStoe_DB = db: trial.AStoe_Sp = sp
                    Case 2: trial.ASheel_DB = db: trial.ASheel_Sp = sp
                End Select
                valid = CheckProjectDesign(trial, r)
                If valid Then
                    Print #f, "PAIR part=" & part & "; db=" & db & "; sp=" & sp & "; cost=" & r.Cost
                    If r.Cost < best Then Print #f, "CHEAPER_PAIR"
                End If
            Next sp
        Next db
    Next part
    trial = HillClimbingOptimization(5000, 5, 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, RandomSeed:=12365)
    Print #f, "HCA5000 seed12365 cost=" & RunBestCost & "; eval=" & RunBestEvaluation & "; folder=" & RunFolder
    ' A known BA answer must be admissible to HCA without changing any engineering inputs.
    trial = HillClimbingOptimization(1, 5, 1.2, 1.8, 2.4, 30, 0.6, 30, 0.075, mat, True, 1, 24, 40, 63, 85, 101, 110, 100, 111, 103, 113, 12345)
    Print #f, "HCA_BA_SEED valid=" & RunBest.IsValid & "; cost=" & RunBestCost
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL=" & Err.Description
    Close #f
End Sub
