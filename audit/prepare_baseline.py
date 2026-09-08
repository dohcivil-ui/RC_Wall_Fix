from pathlib import Path
import hashlib, difflib, re
root=Path(__file__).resolve().parent.parent
out=root/'audit'/'baseline'
def put(p,s): p.write_bytes(s.replace('\n','\r\n').encode('ascii'))
rows=[]
for p in sorted(root.glob('modBA.bas*')):
    s=p.read_bytes().decode('latin1')
    bounds=sorted(set(re.findall(r'Private (Min\w+|Max\w+) As Integer',s)))
    rows.append(f'{p.name}: sha256={hashlib.sha256(p.read_bytes()).hexdigest()} bounds={bounds}')
(out/'BA-versions.txt').write_text('\n'.join(rows),encoding='utf-8')
a=(out/'modShared.bas').read_bytes().decode('latin1').splitlines()
b=(out/'modShared.ACTIVE-external.bas').read_bytes().decode('latin1').splitlines()
(out/'active-vs-local.diff').write_bytes('\n'.join(difflib.unified_diff(a,b,fromfile='local',tofile='active-external')).encode('latin1'))
put(out/'BaselineTest.bas','''Attribute VB_Name = "BaselineTest"
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
    Open App.Path & "\\baseline-runtime.txt" For Output As #f
    Print #f, "Actual VB6 compiled baseline module (external active modShared)"
    Print #f, "H=5 TBase=0.30 H1=1.20; expected active M=10.3823"
    Print #f, "Actual CalculateMomentStem="; m
    Print #f, "Exact regression holds="; Abs(m - 10.3823) < 0.00001
    Print #f, "Old steel accepts DB12@0.25="; CheckSteelOK(m, 0.125, 100, 113)
    Close #f
End Sub
''')
put(out/'BaselineTest.vbp','''Type=Exe
Module=BaselineTest; BaselineTest.bas
Module=modDataStructures; modDataStructures.bas
Module=modWSD; modWSD.bas
Module=modShared; modShared.ACTIVE-external.bas
Startup="Sub Main"
Name="BaselineTest"
ExeName32="BaselineTest.exe"
''')
