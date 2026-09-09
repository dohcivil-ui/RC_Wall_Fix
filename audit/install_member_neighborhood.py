"""Apply the reviewed member-neighborhood change without re-encoding VB6 text."""
from pathlib import Path
import re
P=Path(__file__).resolve().parent.parent
mask='''
    ' Keep components outside the selected move at the current state.
    If Not SearchMoveIncludes(move(0), 1) Then
        Newtt = Currenttt: Newtb = Currenttb
    End If
    If Not SearchMoveIncludes(move(0), 2) Then
        NewTBase = CurrentTBase: NewBase = CurrentBase: NewLToe = CurrentLToe
    End If
    If Not SearchMoveIncludes(move(0), 3) Then
        NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
    End If
    If Not SearchMoveIncludes(move(0), 4) Then
        NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
    End If
    If Not SearchMoveIncludes(move(0), 5) Then
        NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
    End If
'''
for name,sub in [('modBA.bas','GenerateNeighbor_BA'),('modHillClimbing.bas','GenerateNeighbor')]:
    p=P/name; b=p.read_bytes(); start=b.index(('Private Sub '+sub+'(').encode()); end=b.index(b'End Sub',start)
    body=b[start:end].decode('latin1').replace('\r\n','\n')
    assert 'DrawSearchMove' not in body
    body=body.replace('    Dim Step As Integer','    Dim Step As Integer, move(0 To 11) As Integer')
    body=body.replace('    Dim i As Integer','    Dim i As Integer\n\n    Call DrawSearchMove(move)')
    for index,key in enumerate(('tt','tb','TBase','Base','LToe','StemDB','StemSP','ToeDB','ToeSP','HeelDB','HeelSP'),1):
        if index<=5:
            body,n=re.subn(r'Step = Rand\(-\d, \d\)(\n    New'+key+r' =)',f'Step = move({index})'+r'\1',body)
        else:
            body,n=re.subn(r'Step = Rand\(-2, 2\)\n    New'+key+' = Current'+key+r' \+ Step',f'New{key} = move({index})',body)
        assert n==1,(key,n)
    split=body.index('    NewStemDB = move(6)')
    prior=body[:split];pos=prior.rfind('step = Rand(-2, 2)')
    if pos>=0:prior=prior[:pos]+prior[pos:].replace('step = Rand(-2, 2)','uniform DB/SP pairs',1)
    body=prior+body[split:]+mask
    body='\n'.join(line.rstrip() for line in body.split('\n'))
    b=b[:start]+body.replace('\n','\r\n').encode('latin1')+b[end:]
    b=b.replace(b'Public Const HCA_SEARCH_POLICY As String = "HCA_BA_NEIGHBOR_V1"\r\n\r\n',b'')
    p.write_bytes(b)
p=P/'modShared.bas';b=p.read_bytes()
constant=b'Public Const NEIGHBOR_SEARCH_POLICY As String = "MEMBER_WEIGHTED_V1"\r\n'
assert b'NEIGHBOR_SEARCH_POLICY' not in b
pos=b.index(b'Public Const RESULT_CSV_ROOT')
b=b[:pos]+constant+b[pos:]
shared='''Public Sub DrawSearchMove(ByRef move() As Integer)
    ' Both optimizers use these exact twelve draws, in this order.
    ' Geometry entries are index increments; bar entries are absolute indices.
    move(2) = Rand(-2, 2)       ' tb
    move(1) = Rand(-2, 2)       ' tt
    move(3) = Rand(-5, 5)      ' TBase
    move(5) = Rand(-2, 2)      ' LToe
    move(4) = Rand(-1, 1)      ' Base
    move(6) = Rand(DB_MIN, DB_MAX): move(7) = Rand(SP_MIN, SP_MAX)
    move(8) = Rand(DB_MIN, DB_MAX): move(9) = Rand(SP_MIN, SP_MAX)
    move(10) = Rand(DB_MIN, DB_MAX): move(11) = Rand(SP_MIN, SP_MAX)
    ' Whole design: 7/12; each of the five member moves: 1/12.
    move(0) = Rand(0, 11)
    If move(0) > 5 Then move(0) = 0
End Sub

Public Function SearchMoveIncludes(ByVal moveGroup As Integer, ByVal component As Integer) As Boolean
    ' Components: 1 stem geometry, 2 base geometry, 3/4/5 stem/toe/heel bars.
    Select Case moveGroup
        Case 0: SearchMoveIncludes = True                 ' Whole design
        Case 1: SearchMoveIncludes = (component = 1 Or component = 3)
        Case 2: SearchMoveIncludes = (component = 2 Or component = 4 Or component = 5)
        Case 3 To 5: SearchMoveIncludes = (component = moveGroup)
    End Select
End Function

'''
pos=b.index(b'Public Function Rand(')
b=b[:pos]+shared.replace('\n','\r\n').encode('ascii')+b[pos:]
b=b.replace(b'If RunAlgorithm = "HCA" Then Print #f, "SearchPolicy=" & HCA_SEARCH_POLICY',b'Print #f, "SearchPolicy=" & NEIGHBOR_SEARCH_POLICY')
p.write_bytes(b)
