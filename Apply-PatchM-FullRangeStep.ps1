# ============================================================
# Apply-PatchM-FullRangeStep.ps1
# Patch M: Full-Range Step (any-to-any in 1 iteration)
# Date: 2026-05-12
# Supersedes: Patch L (which was Moderate ~2x, not full range)
#
# Workflow:
#   1. Backup current state (post-Patch-L) for forensics
#   2. Rollback Patch L -> restore Patch I+H baseline
#   3. Apply Patch M on Patch I+H baseline
#   4. Verify
#
# Run:
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchM-FullRangeStep.ps1"
# ============================================================

$ErrorActionPreference = 'Stop'

# === Paths ===
$srcDir  = 'D:\rc-rt-optimize-v2\vb6-source'
$baPath  = Join-Path $srcDir 'modBA.bas'
$hcaPath = Join-Path $srcDir 'modHillClimbing.bas'

# === Encoding (CP874 to preserve Thai) ===
$enc = [System.Text.Encoding]::GetEncoding(874)
$ts  = Get-Date -Format 'yyyyMMdd-HHmmss'

# === Pre-flight ===
foreach ($p in @($baPath, $hcaPath)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "[FAIL] File not found: $p" }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch M: Full-Range Step (any-to-any in 1 step)"           -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# ============================================================
# STEP 1: Backup current state (post-Patch-L) for forensics
# ============================================================
Write-Host ""
Write-Host "  [STEP 1] Backup current state (post-Patch-L)" -ForegroundColor Yellow
$baBakPostL  = "$baPath.bak-postPatchL-prePatchM-$ts"
$hcaBakPostL = "$hcaPath.bak-postPatchL-prePatchM-$ts"
Copy-Item -LiteralPath $baPath  -Destination $baBakPostL  -Force
Copy-Item -LiteralPath $hcaPath -Destination $hcaBakPostL -Force
Write-Host "    [OK] $baBakPostL"
Write-Host "    [OK] $hcaBakPostL"

# ============================================================
# STEP 2: Rollback Patch L (restore Patch I+H baseline)
# ============================================================
Write-Host ""
Write-Host "  [STEP 2] Rollback Patch L -> Patch I+H baseline" -ForegroundColor Yellow
$baBakPreL  = Get-ChildItem -Path "$baPath.bak-prePatchL-*"  -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
$hcaBakPreL = Get-ChildItem -Path "$hcaPath.bak-prePatchL-*" -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $baBakPreL -or -not $hcaBakPreL) {
    throw "[FAIL] Pre-Patch-L backup not found. Cannot rollback safely."
}

Copy-Item -LiteralPath $baBakPreL.FullName  -Destination $baPath  -Force
Copy-Item -LiteralPath $hcaBakPreL.FullName -Destination $hcaPath -Force
Write-Host "    [OK] Restored BA  from: $($baBakPreL.Name)"
Write-Host "    [OK] Restored HCA from: $($hcaBakPreL.Name)"

# ============================================================
# STEP 3: Apply Patch M (Full Range) on baseline
# ============================================================
Write-Host ""
Write-Host "  [STEP 3] Apply Patch M (Full Range)" -ForegroundColor Yellow
Write-Host "    tt    : Rand(-2, 2)  -> Rand(-16, 16)"
Write-Host "    tb    : Rand(-2, 2)  -> Rand(-16, 16)"
Write-Host "    TBase : Rand(-5, 5)  -> Rand(-14, 14)"
Write-Host "    LToe  : Rand(-2, 2)  -> Rand(-9, 9)"
Write-Host "    Base  : Rand(-1, 1)  -> Rand(-11, 11)"
Write-Host "    DB x3 : Rand(-2, 2)  -> Rand(-4, 4)"
Write-Host "    SP x3 : Rand(-2, 2)  -> Rand(-3, 3)"
Write-Host ""

$replaces = @(
    # --- Size variables ---
    @{ Tag='tt    '; Old="Step = Rand(-2, 2)`r`n    Newtt = Currenttt + Step";          New="Step = Rand(-16, 16)`r`n    Newtt = Currenttt + Step" }
    @{ Tag='tb    '; Old="Step = Rand(-2, 2)`r`n    Newtb = Currenttb + Step";          New="Step = Rand(-16, 16)`r`n    Newtb = Currenttb + Step" }
    @{ Tag='TBase '; Old="Step = Rand(-5, 5)`r`n    NewTBase = CurrentTBase + Step";    New="Step = Rand(-14, 14)`r`n    NewTBase = CurrentTBase + Step" }
    @{ Tag='LToe  '; Old="Step = Rand(-2, 2)`r`n    NewLToe = CurrentLToe + Step";      New="Step = Rand(-9, 9)`r`n    NewLToe = CurrentLToe + Step" }
    @{ Tag='Base  '; Old="Step = Rand(-1, 1)`r`n    NewBase = CurrentBase + Step";      New="Step = Rand(-11, 11)`r`n    NewBase = CurrentBase + Step" }
    # --- Steel DB x3 ---
    @{ Tag='StemDB'; Old="Step = Rand(-2, 2)`r`n    NewStemDB = CurrentStemDB + Step";  New="Step = Rand(-4, 4)`r`n    NewStemDB = CurrentStemDB + Step" }
    @{ Tag='ToeDB '; Old="Step = Rand(-2, 2)`r`n    NewToeDB = CurrentToeDB + Step";    New="Step = Rand(-4, 4)`r`n    NewToeDB = CurrentToeDB + Step" }
    @{ Tag='HeelDB'; Old="Step = Rand(-2, 2)`r`n    NewHeelDB = CurrentHeelDB + Step";  New="Step = Rand(-4, 4)`r`n    NewHeelDB = CurrentHeelDB + Step" }
    # --- Steel SP x3 ---
    @{ Tag='StemSP'; Old="Step = Rand(-2, 2)`r`n    NewStemSP = CurrentStemSP + Step";  New="Step = Rand(-3, 3)`r`n    NewStemSP = CurrentStemSP + Step" }
    @{ Tag='ToeSP '; Old="Step = Rand(-2, 2)`r`n    NewToeSP = CurrentToeSP + Step";    New="Step = Rand(-3, 3)`r`n    NewToeSP = CurrentToeSP + Step" }
    @{ Tag='HeelSP'; Old="Step = Rand(-2, 2)`r`n    NewHeelSP = CurrentHeelSP + Step";  New="Step = Rand(-3, 3)`r`n    NewHeelSP = CurrentHeelSP + Step" }
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

foreach ($pair in @(@{ Path=$baPath; Label='BA ' }, @{ Path=$hcaPath; Label='HCA' })) {
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-16, 16)' -Expected 2 -Note 'tt, tb'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-14, 14)' -Expected 1 -Note 'TBase'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-11, 11)' -Expected 1 -Note 'Base'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-9, 9)'   -Expected 1 -Note 'LToe'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-4, 4)'   -Expected 3 -Note 'Steel DB x3'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-3, 3)'   -Expected 3 -Note 'Steel SP x3'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-2, 2)'   -Expected 0 -Note 'all replaced'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-1, 1)'   -Expected 0 -Note 'all replaced'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-5, 5)'   -Expected 0 -Note 'all replaced'
    Write-Host ""
}

Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Patch M applied successfully"                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Backups available:"
Write-Host "    Post-Patch-L (forensics)  : $baBakPostL"
Write-Host "                                $hcaBakPostL"
Write-Host "    Pre-Patch-L  (Patch I+H)  : $($baBakPreL.FullName)"
Write-Host "                                $($hcaBakPreL.FullName)"
Write-Host ""
Write-Host "  Next steps:"
Write-Host "    1. Open VB6 IDE -> reload modBA.bas, modHillClimbing.bas"
Write-Host "    2. Compile -> Run 30 trials @ H=3 (BA, then HCA)"
Write-Host "    3. Save CSVs:"
Write-Host "         loopPrice-BA-H3-patchM.csv"
Write-Host "         loopPrice-HCA-H3-patchM.csv"
Write-Host "    4. Send CSVs back for analysis"
Write-Host ""
Write-Host "  Rollback to Patch I+H baseline (if needed):"
Write-Host "    Copy-Item '$($baBakPreL.FullName)'  '$baPath'  -Force"
Write-Host "    Copy-Item '$($hcaBakPreL.FullName)' '$hcaPath' -Force"
Write-Host ""
