    Dim savedSensitivityWSD As WSDParams, sensitivityFile As Integer
    savedSensitivityWSD = currentWSD
    Dim savedSensitivityQa As Double
    savedSensitivityQa = qa: qa = 30
    H = 5: H1 = 1.2: PassiveFactor = 1
    d = Fixture(0.25, 0.6, 2.5, 0.3)
    Call AssertTrue("n sensitivity example overturning", CheckFS_OT(d, ca))
    Call AssertTrue("n sensitivity example sliding", CheckFS_SL(d, ca))
    Call AssertTrue("n sensitivity example bearing", CheckFS_BC(d, ca, e, qt, qh))
    Call Near("n sensitivity stem moment", CalculateMomentStem(d), 8.324000000000)
    sensitivityFile = FreeFile
    Open App.Path & "\modular-ratio-sensitivity-native.csv" For Output As #sensitivityFile
    Print #sensitivityFile, "model,n,fc,fs,flexural_acceptance"
    currentWSD = CalculateWSDParameters(4000, 320)
    currentWSD.n = 9#  ' Isolated comparison only; restore production parameters below.
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("n sensitivity case0 concrete", ca, 142.175674555456)
    Call Near("n sensitivity case0 steel", sa, 1029.856663015879)
    ok = CheckSteelOK(CalculateMomentStem(d), SectionDepth(d.tb, 104), 104, 110)
    Call AssertTrue("n sensitivity case0 flexural decision", ok)
    Print #sensitivityFile, "fixed9," & CsvNumber(currentWSD.n) & "," & CsvNumber(ca) & "," & CsvNumber(sa) & "," & CStr(ok)
    currentWSD = CalculateWSDParameters(4000, 320)
    Call SectionStresses(CalculateMomentStem(d), SectionDepth(d.tb, 104), CalculateAsProv(104, 110), currentWSD.n, ca, sa, ja)
    Call Near("n sensitivity case1 concrete", ca, 148.452949241681)
    Call Near("n sensitivity case1 steel", sa, 1017.452282440440)
    ok = CheckSteelOK(CalculateMomentStem(d), SectionDepth(d.tb, 104), 104, 110)
    Call AssertTrue("n sensitivity case1 flexural decision", Not ok)
    Print #sensitivityFile, "EsEc," & CsvNumber(currentWSD.n) & "," & CsvNumber(ca) & "," & CsvNumber(sa) & "," & CStr(ok)
    Close #sensitivityFile
    currentWSD = savedSensitivityWSD
    qa = savedSensitivityQa
