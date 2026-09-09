"""Freeze one geometry-proposal candidate and its staged evaluation protocol."""
from pathlib import Path
import difflib
import hashlib
import json

root = Path(__file__).resolve().parent.parent
out = root / 'audit/ba-geometry-hold'
out.mkdir(exist_ok=False)
old = (root / 'audit/ba-lazy-pairs/candidate-modBA.bas').read_bytes().decode('latin1').replace('\r\n', '\n')
marker = '''Private Sub DrawGeometryMove(move() As Integer)
    move(2) = Rand(-1, 1): move(1) = Rand(-1, 1)
    move(3) = Rand(-1, 1): move(5) = Rand(-1, 1): move(4) = Rand(-1, 1)
End Sub'''
replacement = '''Private Sub DrawGeometryMove(move() As Integer)
    move(2) = GeometryStep(): move(1) = GeometryStep()
    move(3) = GeometryStep(): move(5) = GeometryStep(): move(4) = GeometryStep()
End Sub

Private Function GeometryStep() As Integer
    ' Six equiprobable tickets: one step down, four stay, one step up.
    ' Every geometry variable draws independently before the full evaluation.
    Select Case Rand(1, 6)
        Case 1: GeometryStep = -1
        Case 6: GeometryStep = 1
        Case Else: GeometryStep = 0
    End Select
End Function'''
assert old.count(marker) == 1
new = old.replace(marker, replacement)
(out / 'candidate-modBA.bas').write_bytes(new.replace('\n', '\r\n').encode('latin1'))
for name in ('modHillClimbing.bas', 'modShared.bas'):
    (out / ('candidate-' + name)).write_bytes((root / name).read_bytes())
(out / 'change-from-lazy-pairs.diff').write_text(''.join(difflib.unified_diff(old.splitlines(True), new.splitlines(True), fromfile='lazy-pairs/modBA.bas', tofile='geometry-hold/modBA.bas')), encoding='utf-8')
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
project = (root / 'RC_RT_HCA_v2.vbp').read_text(encoding='ascii')
names = [line.split(';')[-1].strip() if line.startswith('Module=') else line[5:] for line in project.splitlines() if line.startswith(('Module=', 'Form='))]
manifests = {
    'root-sources-before.json': {n: sha(root / n) for n in names},
    'root-project-binaries-before.json': {str(p.relative_to(root)): sha(p) for p in [root / 'RC_RT_HCA_v2.vbp', *root.glob('*.exe'), *root.glob('*.frx')]},
    'result-csv-before.json': {str(p.relative_to(root)): sha(p) for p in (root / 'result_csv').rglob('*') if p.is_file()},
}
for name, value in manifests.items():
    (out / name).write_text(json.dumps(value, indent=2) + '\n')
(out / 'PROTOCOL.md').write_text('''# Geometry hold: one candidate, staged exploratory comparison

User authorized further development toward approximately10% overall advantage using complete outcomes, attainment, target-first evaluations and SD. Perfect10/10 or every-run victories are not required. Prior BA geometry/pair disruption is investigated from recorded lazy-pairs traces. No formula/domain/unit/price/check changes, no root/result_csv mutation, no seed controls, no30-trial batch, no unreported outcomes.

## Frozen candidate

BA changes only DrawGeometryMove and adds GeometryStep. Each of the five geometry draws independently uses six equiprobable tickets:1=-1index,2..5=0,6=+1index. Thus P0=2/3 and Pdown=Pup=1/6 before original clamps. All five draws still occur, followed by three independent half-stay coupled-pair draws from previous candidate. No forced member grouping, permanent coordinate fixing, target/height special cases, or rejection filtering during proposal. Existing support, midpointjointtb/TBase/Base, quotas20/40/...,centerupdate,accept/restore and5000budget unchanged. EachBA run1initial+22midpoint+4977neighbor. Every evaluated design goes through all original checks.

HCA/shared are byte-identical copies of original root, retaining grouped widersteps/absolutebars/ttrepair. BA's private samplers do not call shared DrawSearchMove. Compare wholemethods; no causal claim about bisection alone. At four target-boundary geometry coordinates, preservation probability changes(2/3)^4 to(5/6)^4, a2.4414factor for the same incumbent. Early descent can slow and duplicates can increase; local probability is not a whole-run prediction.

## Native checks before performance

Before old sampler must fail newticketcontract. After enumerate all6^5 geometryticketvectors through actual native DrawGeometryMove/GeometryStep with audit-only forcedtickets, plus original20pair-origin/exhaustiveticketchecks. Bounded native mechanism runs preserve checks for allfivegeometry/threesteel draws, strictaccept/restore,midpoint,jointbounds,budgetedgecases,impossiblebearing. Forcedtickets never appear in performance sources. Allcomplete actual performance vectors and exactnativeDoublebytes recorded; source/outputhashes independently checked.

## Stages fixed before outcomes

Stage1: one compiled actualGUI BA10x5000 + originalHCA10x5000 atH3/fc240 and unchanged remainingGUIinputs, labelgeometry-hold-10. No retries or tuning that condition after outcome.

Stage2 is a prespecified fresh replication of EXACT same candidate, one additional BA10x5000+HCA10x5000, labelgeometry-hold-replication-10, only if stage1 known-target attainment is at least HCA and its mean evaluations spent before reaching target (misses use complete5000budget) is at least10% lower. This is a trigger for more evidence, not a universal pass rule imposed on the user's objective. Both stage results must be retained and reported separately plus aggregate descriptive data if stage2runs. No third batch or altered candidate in this condition. These are bounded exploratory10-trial stages, not a30-trial research run or significance test. A favorable first sample followed by an unfavorable replication must not be hidden.

Report target2942.34192 consistently, allfinalprices, everyfirsttarget, successfraction, mean/median/SD of successful targettimes, and capped effort over ALL runs (min(targetfirst,5000), missesvisible). Capped effort is observedbudget spent, not inferred time-to-event after5000. Report attainment at100/250/500/1000/2500/5000evaluations. Compare bothsuccess/effort with SDsupport; no invented compositeweights or cherry-pickedcheckpoint. Relativeeffort reduction=(HCA-BA)/HCA. No automatic promotion or universal/statistical superiority claim.
''', encoding='utf-8')
print('Frozen candidate and protocol:', out)
