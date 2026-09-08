# ============================================================
# Apply-PatchN-NarrowStep.ps1
# Patch N: Narrowest Step - Rand(-1, 1) for ALL variables
# Date: 2026-05-12
# Supersedes: Patch L (Moderate), Patch M (Full Range)
#
# Rationale: Structural quantities are in meters. Even ±0.05m
#            grid steps are physically meaningful. We want the
#            algorithm to inch toward the optimum, not jump.
#            Goal: slowest convergence, finest exploitation.
#
# Workflow:
#   1. Backup current state (post-Patch-M) for forensics
#   2. Rollback to Patch I+H baseline (from prePatchL backup)
#   3. Apply Patch N on baseline
#   4. Verify
#
# Run:
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchN-NarrowStep.ps1"
# ============================================================

$ErrorActionPreference = 'Stop'

# === Paths ===
$srcDir  = 'D:\rc-rt-optimize-v2\vb6-source'
$baPath  = Join-Path $srcDir 'modBA.bas'
$hcaPath = Join-Path $srcDir 'modHillClimbing.bas'

# === Encoding (CP874) ===
$enc = [System.Text.Encoding]::GetEncoding(874)
$ts  = Get-Date -Format 'yyyyMMdd-HHmmss'

# === Pre-flight ===
foreach ($p in @($baPath, $hcaPath)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "[FAIL] File not found: $p" }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch N: Narrowest Step - Rand(-1, 1) all variables"       -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# ============================================================
# STEP 1: Backup current state (post-Patch-M) for forensics
# ============================================================
Write-Host ""
Write-Host "  [STEP 1] Backup current state (post-Patch-M)" -ForegroundColor Yellow
$baBakPostM  = "$baPath.bak-postPatchM-prePatchN-$ts"
$hcaBakPostM = "$hcaPath.bak-postPatchM-prePatchN-$ts"
Copy-Item -LiteralPath $baPath  -Destination $baBakPostM  -Force
Copy-Item -LiteralPath $hcaPath -Destination $hcaBakPostM -Force
Write-Host "    [OK] $baBakPostM"
Write-Host "    [OK] $hcaBakPostM"

# ============================================================
# STEP 2: Rollback to Patch I+H baseline (from prePatchL backup)
# ============================================================
Write-Host ""
Write-Host "  [STEP 2] Rollback -> Patch I+H baseline" -ForegroundColor Yellow
$baBakPreL  = Get-ChildItem -Path "$baPath.bak-prePatchL-*"  -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
$hcaBakPreL = Get-ChildItem -Path "$hcaPath.bak-prePatchL-*" -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $baBakPreL -or -not $hcaBakPreL) {
    throw "[FAIL] Pre-Patch-L baseline backup not found."
}

Copy-Item -LiteralPath $baBakPreL.FullName  -Destination $baPath  -Force
Copy-Item -LiteralPath $hcaBakPreL.FullName -Destination $hcaPath -Force
Write-Host "    [OK] Restored BA  from: $($baBakPreL.Name)"
Write-Host "    [OK] Restored HCA from: $($hcaBakPreL.Name)"

# ============================================================
# STEP 3: Apply Patch N on Patch I+H baseline
# ============================================================
# Baseline state (Patch I+H):
#   tt: Rand(-2, 2), tb: Rand(-2, 2), TBase: Rand(-5, 5)
#   LToe: Rand(-2, 2), Base: Rand(-1, 1), Steel x6: Rand(-2, 2)
#
# Target (Patch N - all Rand(-1, 1)):
#   tt, tb, TBase, LToe, Steel x6 -> change (10 vars)
#   Base -> NO CHANGE (already Rand(-1, 1))
# ============================================================
Write-Host ""
Write-Host "  [STEP 3] Apply Patch N (Rand(-1, 1) everywhere)" -ForegroundColor Yellow
Write-Host "    tt    : Rand(-2, 2) -> Rand(-1, 1)"
Write-Host "    tb    : Rand(-2, 2) -> Rand(-1, 1)"
Write-Host "    TBase : Rand(-5, 5) -> Rand(-1, 1)"
Write-Host "    LToe  : Rand(-2, 2) -> Rand(-1, 1)"
Write-Host "    Base  : Rand(-1, 1) -> Rand(-1, 1)  [no change, already narrowest]"
Write-Host "    DB x3 : Rand(-2, 2) -> Rand(-1, 1)"
Write-Host "    SP x3 : Rand(-2, 2) -> Rand(-1, 1)"
Write-Host ""

# 10 replacements (Base unchanged)
$replaces = @(
    @{ Tag='tt    '; Old="Step = Rand(-2, 2)`r`n    Newtt = Currenttt + Step";          New="Step = Rand(-1, 1)`r`n    Newtt = Currenttt + Step" }
    @{ Tag='tb    '; Old="Step = Rand(-2, 2)`r`n    Newtb = Currenttb + Step";          New="Step = Rand(-1, 1)`r`n    Newtb = Currenttb + Step" }
    @{ Tag='TBase '; Old="Step = Rand(-5, 5)`r`n    NewTBase = CurrentTBase + Step";    New="Step = Rand(-1, 1)`r`n    NewTBase = CurrentTBase + Step" }
    @{ Tag='LToe  '; Old="Step = Rand(-2, 2)`r`n    NewLToe = CurrentLToe + Step";      New="Step = Rand(-1, 1)`r`n    NewLToe = CurrentLToe + Step" }
    @{ Tag='StemDB'; Old="Step = Rand(-2, 2)`r`n    NewStemDB = CurrentStemDB + Step";  New="Step = Rand(-1, 1)`r`n    NewStemDB = CurrentStemDB + Step" }
    @{ Tag='ToeDB '; Old="Step = Rand(-2, 2)`r`n    NewToeDB = CurrentToeDB + Step";    New="Step = Rand(-1, 1)`r`n    NewToeDB = CurrentToeDB + Step" }
    @{ Tag='HeelDB'; Old="Step = Rand(-2, 2)`r`n    NewHeelDB = CurrentHeelDB + Step";  New="Step = Rand(-1, 1)`r`n    NewHeelDB = CurrentHeelDB + Step" }
    @{ Tag='StemSP'; Old="Step = Rand(-2, 2)`r`n    NewStemSP = CurrentStemSP + Step";  New="Step = Rand(-1, 1)`r`n    NewStemSP = CurrentStemSP + Step" }
    @{ Tag='ToeSP '; Old="Step = Rand(-2, 2)`r`n    NewToeSP = CurrentToeSP + Step";    New="Step = Rand(-1, 1)`r`n    NewToeSP = CurrentToeSP + Step" }
    @{ Tag='HeelSP'; Old="Step = Rand(-2, 2)`r`n    NewHeelSP = CurrentHeelSP + Step";  New="Step = Rand(-1, 1)`r`n    NewHeelSP = CurrentHeelSP + Step" }
)

function Apply-To-File {
    param([string]$Path, [string]$Label)
    Write-Host "    --- $Label ---" -ForegroundColor Yellow
    $text = [System.IO.File]::ReadAllText($Path, $enc)
    $count = 0
    foreach ($r in $replaces) {
        if (-not $text.Contains($r.Old)) {
            throw "[FAIL] $Label : pattern not found for '$($r.Tag.Trim())'"
        }
        $text = $text.Replace($r.Old, $r.New)
        Write-Host "      [OK] $($r.Tag)"
        $count++
    }
    [System.IO.File]::WriteAllText($Path, $text, $enc)
    Write-Host "      Saved: $Path ($count replacements)" -ForegroundColor Green
}

Apply-To-File -Path $baPath  -Label 'BA  (modBA.bas)'
Apply-To-File -Path $hcaPath -Label 'HCA (modHillClimbing.bas)'

# ============================================================
# STEP 4: Verify
# ============================================================
Write-Host ""
Write-Host "  [STEP 4] Verification" -ForegroundColor Yellow
function Show-Count {
    param([string]$Path, [string]$Label, [string]$Pattern, [int]$Expected, [string]$Note)
    $text = [System.IO.File]::ReadAllText($Path, $enc)
    $escaped = [regex]::Escape($Pattern)
    $count = ([regex]::Matches($text, $escaped)).Count
    $status = if ($count -eq $Expected) { '[OK]' } else { '[!!]' }
    $color  = if ($count -eq $Expected) { 'Green' } else { 'Red' }
    Write-Host ("    {0} {1} '{2,-22}' : found {3,2} (expect {4,2} -- {5})" -f $status, $Label, $Pattern, $count, $Expected, $Note) -ForegroundColor $color
}

# Expected post-Patch-N:
#   Rand(-1, 1): 11 (tt, tb, TBase, LToe, Base, 6xSteel)
#   Rand(-2, 2): 0, Rand(-5, 5): 0, all others: 0
foreach ($pair in @(@{ Path=$baPath; Label='BA ' }, @{ Path=$hcaPath; Label='HCA' })) {
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-1, 1)' -Expected 11 -Note 'all 11 variables'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-2, 2)' -Expected 0  -Note 'should be 0'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-5, 5)' -Expected 0  -Note 'should be 0'
    Write-Host ""
}

Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Patch N applied successfully"                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Backups:"
Write-Host "    Post-Patch-M (forensics)  : $baBakPostM"
Write-Host "                                $hcaBakPostM"
Write-Host "    Pre-Patch-L (Patch I+H)   : $($baBakPreL.FullName)"
Write-Host "                                $($hcaBakPreL.FullName)"
Write-Host ""
Write-Host "  Next steps:"
Write-Host "    1. Open VB6 IDE -> reload modBA.bas, modHillClimbing.bas"
Write-Host "    2. Compile -> Run 30 trials @ H=3 (BA, then HCA)"
Write-Host "    3. WATCH bestIter values: if they cluster near 100, convergence"
Write-Host "       is incomplete -- consider increasing maxIter to 500 or 1000"
Write-Host "    4. Save CSVs:"
Write-Host "         loopPrice-BA-H3-patchN.csv"
Write-Host "         loopPrice-HCA-H3-patchN.csv"
Write-Host "    5. Send CSVs back for analysis"
Write-Host ""
