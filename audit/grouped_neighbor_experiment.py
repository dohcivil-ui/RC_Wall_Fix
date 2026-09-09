"""Apply the candidate grouped-neighbor policy only in a supplied diagnostic folder."""
from pathlib import Path
import sys
folder=Path(sys.argv[1]).resolve()
block='''
    ' Select a whole-design or member move with the same rule in BA and HCA.
    ' Existing random step ranges and their draw order are preserved above.
    Select Case Rand(0, 4)
        Case 0 ' All eleven variables.
        Case 1 ' Geometry only.
            NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
            NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
            NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
        Case 2 ' Stem bars only.
            Newtt = Currenttt: Newtb = Currenttb: NewTBase = CurrentTBase
            NewBase = CurrentBase: NewLToe = CurrentLToe
            NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
            NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
        Case 3 ' Toe bars only.
            Newtt = Currenttt: Newtb = Currenttb: NewTBase = CurrentTBase
            NewBase = CurrentBase: NewLToe = CurrentLToe
            NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
            NewHeelDB = CurrentHeelDB: NewHeelSP = CurrentHeelSP
        Case 4 ' Heel bars only.
            Newtt = Currenttt: Newtb = Currenttb: NewTBase = CurrentTBase
            NewBase = CurrentBase: NewLToe = CurrentLToe
            NewStemDB = CurrentStemDB: NewStemSP = CurrentStemSP
            NewToeDB = CurrentToeDB: NewToeSP = CurrentToeSP
    End Select
'''
for name,sub in [('modBA.bas','GenerateNeighbor_BA'),('modHillClimbing.bas','GenerateNeighbor')]:
    p=folder/name; b=p.read_bytes(); start=b.index(('Private Sub '+sub+'(').encode()); end=b.index(b'End Sub',start)
    assert b'Select Case Rand(0, 4)' not in b[start:end]
    p.write_bytes(b[:end]+block.replace('\n','\r\n').encode()+b[end:])
shared=folder/'modShared.bas'
b=shared.read_bytes().replace(b'If RunAlgorithm = "HCA" Then Print #f, "SearchPolicy=" & HCA_SEARCH_POLICY',b'Print #f, "SearchPolicy=GROUPED_NEIGHBOR_EXPERIMENT_V1"')
shared.write_bytes(b)
