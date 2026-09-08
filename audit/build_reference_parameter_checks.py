"""Independent reference equations; emits fixtures run by the VB6 regression EXE.
Pongnathee Chapter 10, printed pages 342-343, user-supplied PDF pages 10-11.
No claim that this teaching example establishes EIT 011007-19 compliance.
"""
from pathlib import Path
import math
import json

P = Path(__file__).resolve().parent
rows = []
vb = ['    Dim refParams As WSDParams, refFile As Integer',
      '    refFile = FreeFile',
      '    Open App.Path & "\\wsd-reference-native.csv" For Output As #refFile',
      '    Print #refFile, "fc_prime,fy,n,k_bal,j_bal,R_bal,rho_bal"']
for fc_prime, fy in [(180,3000),(210,4000),(240,4000),(280,4000),(320,4000)]:
    ec = 15100 * math.sqrt(fc_prime)
    n = 2040000 / ec
    fc = .45 * fc_prime
    fs = 1500 if fy == 3000 else 1700
    # At balance, a linear strain diagram gives x/d = eps_c/(eps_c+eps_s).
    eps_c, eps_s = fc/ec, fs/2040000
    k = eps_c/(eps_c+eps_s)
    j = 1-k/3
    R = fc*k*j/2
    rho = fc*k/(2*fs)
    rows.append(dict(fc_prime=fc_prime,fy=fy,Es=2040000,Ec=ec,n=n,k=k,j=j,R=R,rho_bal=rho))
    vb.append(f'    refParams = CalculateWSDParameters({fy}, {fc_prime})')
    for field,value in [('n',n),('k',k),('j',j),('R',R)]:
        vb.append(f'    Call Near("reference fc{fc_prime} {field}", refParams.{field}, {value:.12f})')
    vb.append(f'    Call Near("reference fc{fc_prime} balanced rho", CalculateRhoBalanced(refParams), {rho:.12f})')
    vb.append(f'    Print #refFile, "{fc_prime},{fy}," & CsvNumber(refParams.n) & "," & CsvNumber(refParams.k) & "," & CsvNumber(refParams.j) & "," & CsvNumber(refParams.R) & "," & CsvNumber(CalculateRhoBalanced(refParams))')
    # The actual-bar solver must reproduce both allowable stresses at balance.
    vb.append('    Call SectionStresses(CalculateMomentCapacity(refParams.R, 0.3), 0.3, CalculateRhoBalanced(refParams) * 100# * 30#, refParams.n, ca, sa, ja)')
    for field,value in [('ca',fc),('sa',fs),('ja',j)]:
        vb.append(f'    Call Near("reference fc{fc_prime} actual balanced {field}", {field}, {value:.12f})')
vb.append('    Close #refFile')
# A sparse real layout is not balanced: its neutral axis differs substantially.
n = rows[-1]['n']
area = math.pi*1.6**2/(4*.1)
depth = 31.7
x = (-n*area+math.sqrt((n*area)**2+2*100*n*area*depth))/100
lever = depth-x/3
actual = dict(As=area,d_cm=depth,k=x/depth,j=lever/depth,
              fc=9.7262e5/(50*x*lever),fs=9.7262e5/(area*lever))
vb.append('    refParams = CalculateWSDParameters(4000, 320)')
vb.append('    Call SectionStresses(9.7262, 0.317, CalculateAsProv(101, 110), refParams.n, ca, sa, ja)')
for field,value in [('ca',actual['fc']),('sa',actual['fs']),('ja',actual['j'])]:
    vb.append(f'    Call Near("actual DB16 section {field}", {field}, {value:.12f})')
vb.append('    Call AssertTrue("actual bar j differs from balanced j", Abs(ja - refParams.j) > 0.01)')
vb.append('    On Error Resume Next')
vb.append('    refParams = CalculateWSDParameters(4000, 0)')
vb.append('    Call AssertTrue("zero concrete strength rejected", Err.Number <> 0)')
vb.append('    Err.Clear')
vb.append('    On Error GoTo Fatal')
(P/'reference_parameter_checks.inc').write_text('\n'.join(vb)+'\n',encoding='ascii')
(P/'wsd-reference-independent.json').write_text(json.dumps(dict(materials=rows,actual_DB16=actual,
    pdf_example_note='For fc_prime=210 the PDF rounds n=9.32 to 9. This is not a constant for every concrete strength.',
    model_note='Use unrounded Es/Ec for calculation; round display only. EIT edition applicability remains unverified.'),indent=2),encoding='ascii')
print('Reference parameter fixtures generated: 5 materials and actual-bar/balanced-section distinction.')
