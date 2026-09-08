"""Independent sampled/bracketed height check, separate from native root solving."""
import math
import numpy as np

n = 9.0  # User-selected project assumption for this profile suite.

def stem_profile(h, h1, tb, tt, tbase, cover, db, spacing, eta):
    hs, hp = h-tbase, h1-tbase
    assert hs > 0 and 0 <= hp <= hs
    assert tt-cover-db/2000 > 0
    slope = (tb-tt)/hs
    def depth(y): return tb-cover-db/2000-slope*y
    def shear(y):
        return abs(.3*(hs-y)**2-eta*2.7*np.maximum(hp-y,0)**2)/(10*depth(y))
    levels = np.unique(np.append(np.linspace(0,hs,20001), hp))
    values = shear(levels)
    idx = int(np.argmax(values))
    lo, hi = levels[max(0,idx-1)], levels[min(len(levels)-1,idx+1)]
    ratio = (math.sqrt(5)-1)/2
    for _ in range(80):
        left, right = hi-ratio*(hi-lo), lo+ratio*(hi-lo)
        if shear(left) > shear(right): hi = right
        else: lo = left
    candidates = [0,hp,hs,(lo+hi)/2]
    peak_height = max(candidates, key=shear)
    area = math.pi*(db/10)**2/4/spacing
    dc = depth(levels)*100
    neutral = (-n*area+np.sqrt((n*area)**2+200*n*area*dc))/100
    lever = dc-neutral/3
    moment = .1*(hs-levels)**3-eta*.9*np.maximum(hp-levels,0)**3
    fc = np.abs(moment)*1e5/(50*neutral*lever)
    fs = np.abs(moment)*1e5/(area*lever)
    return dict(face_v=float(shear(0)),max_v=float(shear(peak_height)),shear_height=float(peak_height),
                max_M=float(np.max(np.abs(moment))),max_fc=float(fc.max()),max_fs=float(fs.max()),
                moment_height=float(levels[int(np.argmax(np.abs(moment)))]),
                concrete_height=float(levels[int(np.argmax(fc))]),steel_height=float(levels[int(np.argmax(fs))]),stations=len(levels))
