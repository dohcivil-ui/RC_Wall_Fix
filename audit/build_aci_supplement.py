"""Native supplementary ACI 318-99 calculations; never enable design acceptance."""
from pathlib import Path
P = Path(__file__).resolve().parent
project = (P/'FullSizing.vbp').read_bytes().decode('ascii')
project = project.replace('FullSizingMain', 'AciSupplement').replace('NativeFullSizing', 'NativeAciSupplement').replace('FullSizing.exe', 'AciSupplement.exe')
(P/'AciSupplement.vbp').write_bytes(project.encode('ascii'))
code = '''Attribute VB_Name = "AciSupplement"
Option Explicit

Public Sub Main()
    Dim f As Integer, d As Design, db As Integer, bar As Integer, part As Integer
    Dim depth As Double, stress As Double, limit As Double, thickness As Double
    On Error GoTo Failed
    f = FreeFile
    Open App.Path & "\\aci-supplement-native.csv" For Output As #f
    InitializeArrays
    H = 5#: H1 = 1.2: gamma_soil = 1.8: gamma_concrete = 2.4
    phi = 30#: mu = 0.6: qa = 30#: cover = 0.075
    currentWSD = CalculateWSDParameters(4000, 320)
    If PassiveFactor <> 1# Or WSDCriteriaReady() Then Err.Raise 5, , "Unexpected production settings"
    d.tt = 0.2: d.tb = 0.4: d.TBase = 0.3: d.Base = 2.5
    d.LToe = 1#: d.LHeel = 1.1
    ' 1 psi = 0.0703069579640175 kgf/cm2; convert both sides of the root formula.
    limit = 1.1 * Sqr(320# / 0.0703069579640175) * 0.0703069579640175
    Print #f, "item,value,unit,scope"
    Print #f, "vc_aci99," & CsvNumber(limit) & ",kgf/cm2,A.7.4.1 normal-weight shear and flexure only"
    Print #f, "vc_textbook," & CsvNumber(0.29 * Sqr(320#)) & ",kgf/cm2,Pongnathee printed349 rounded coefficient"
    Print #f, "fs_grade40_50," & CsvNumber(20000# * 0.0703069579640175) & ",kgf/cm2,A.3.2(a) not an SD40 grade assignment"
    Print #f, "fs_grade60_plus," & CsvNumber(24000# * 0.0703069579640175) & ",kgf/cm2,A.3.2(b) not an SD40 grade assignment"
    Print #f, "fy_sd40_psi," & CsvNumber(4000# / 0.0703069579640175) & ",psi,project material conversion only"
    For part = 0 To 2
        If part = 0 Then bar = 16: thickness = d.tb Else bar = 20: thickness = d.TBase
        For db = DB_MIN To DB_MAX
            If WP_DB(db) = bar Then Exit For
        Next db
        depth = SectionDepth(thickness, db)
        stress = MemberShearStress(d, part, depth)
        Print #f, "member_" & part & "_shear," & CsvNumber(stress) & ",kgf/cm2,existing project load case only - not complete ACI load envelope"
    Next part
    Print #f, "stem_horizontal_min_gross," & CsvNumber(0.0025 * 100# * d.tb * 100#) & ",cm2/m,14.1.2 and 14.3.3(b) other deformed bars"
    Print #f, "production_criteria_ready," & CInt(WSDCriteriaReady()) & ",boolean,must remain 0 - supplementary calculation only"
    Print #f, "NATIVE_COMPLETE,1,boolean,no optimizer or acceptance gate changed"
    Close #f
    Exit Sub
Failed:
    Print #f, "FATAL," & Err.Number & ",error," & Err.Description
    Close #f
End Sub
'''
(P/'AciSupplement.bas').write_bytes(code.replace('\n','\r\n').encode('ascii'))
