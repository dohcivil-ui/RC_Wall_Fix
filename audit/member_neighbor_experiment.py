"""Diagnostic candidate: coupled member moves and unrestricted bar-pair draws."""
from pathlib import Path
import re,sys
folder=Path(sys.argv[1]).resolve()
block='''
    Select Case Rand(0, 5)
        Case 0 ' Whole design.
        Case 1 ' Stem geometry and its main bars.
            NewTBase = CurrentTBase: NewBase = CurrentBase: NewLToe = CurrentLToe
            NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
            NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
        Case 2 ' Base geometry and its main bars.
            Newtt = Currenttt: Newtb = Currenttb
            NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
        Case 3 ' Stem main bars only.
            Newtt = Currenttt: Newtb = Currenttb: NewTBase = CurrentTBase
            NewBase = CurrentBase: NewLToe = CurrentLToe
            NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
            NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
        Case 4 ' Toe main bars only.
            Newtt = Currenttt: Newtb = Currenttb: NewTBase = CurrentTBase
            NewBase = CurrentBase: NewBase = CurrentBase: NewLToe = CurrentLToe
            NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
            NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
        Case 5 ' Heel main bars only.
            Newtt = Currenttt: Newtb = Currenttb: NewTBase = CurrentTBase
            NewBase = CurrentBase: NewLToe = CurrentLToe
            NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
            NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
    End Select
'''.replace('NewBase = CurrentBase: NewBase = CurrentBase:', 'NewBase = CurrentBase:')
for name,sub in [('modBA.bas','GenerateNeighbor_BA'),('modHillClimbing.bas','GenerateNeighbor')]:
    p=folder/name;b=p.read_bytes(); start=b.index(('Private Sub '+sub+'(').encode());end=b.index(b'End Sub',start)
    body=b[start:end].decode('latin1')
    for member in ('Stem','Toe','Heel'):
        for kind in ('DB','SP'):
            body=body.replace(f'Step = Rand(-2, 2)\r\n    New{member}{kind} = Current{member}{kind} + Step',f'New{member}{kind} = Rand({kind}_MIN, {kind}_MAX)')
    b=b[:start]+body.encode('latin1')+block.replace('\n','\r\n').encode()+b[end:]
    p.write_bytes(b)
p=folder/'modShared.bas';b=p.read_bytes().replace(b'If RunAlgorithm = "HCA" Then Print #f, "SearchPolicy=" & HCA_SEARCH_POLICY',b'Print #f, "SearchPolicy=MEMBER_NEIGHBOR_EXPERIMENT_V1"');p.write_bytes(b)
