Attribute VB_Name = "BaselineTest"
Option Explicit
Public Sub Main()
    Dim f As Integer, d As Design, m As Double
    InitializeArrays
    H = 5: H1 = 1.2: gamma_soil = 1.8: phi = 30
    gamma_concrete = 2.4: cover = 0.075
    currentWSD = CalculateWSDParameters(4000, 320)
    d.tb = 0.2: d.TBase = 0.3
    m = CalculateMomentStem()
    f = FreeFile
    Open App.Path & "\baseline-runtime.txt" For Output As #f
    Print #f, "Actual VB6 compiled baseline module (external active modShared)"
    Print #f, "H=5 TBase=0.30 H1=1.20; expected active M=10.3823"
    Print #f, "Actual CalculateMomentStem="; m
    Print #f, "Exact regression holds="; Abs(m - 10.3823) < 0.00001
    Print #f, "Old steel accepts DB12@0.25="; CheckSteelOK(m, 0.125, 100, 113)
    Close #f
End Sub
