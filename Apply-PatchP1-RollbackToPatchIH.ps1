# ============================================================
# Apply-PatchP1-RollbackToPatchIH.ps1
# Patch P-1: Rollback Patches L, M, N, O -> restore Patch I+H baseline
# Date: 2026-05-12
#
# Rationale: Patches L, M, N, O were exploratory and didn't address
#            the real issue (initial cost discrepancy bug).
#            Return to known-good Patch I+H state for systematic debug.
#
# Run:
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchP1-RollbackToPatchIH.ps1"
# ============================================================

$ErrorActionPreference = 'Stop'

# === Paths ===
$srcDir  = 'D:\rc-rt-optimize-v2\vb6-source'
$baPath  = Join-Path $srcDir 'modBA.bas'
$hcaPath = Join-Path $srcDir 'modHillClimbing.bas'

$ts = Get-Date -Format 'yyyyMMdd-HHmmss'

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch P-1: Rollback to Patch I+H baseline"                  -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# === Backup current state (Patch O) ===
$baBakNow  = "$baPath.bak-prePatchP1-$ts"
$hcaBakNow = "$hcaPath.bak-prePatchP1-$ts"
Copy-Item -LiteralPath $baPath  -Destination $baBakNow  -Force
Copy-Item -LiteralPath $hcaPath -Destination $hcaBakNow -Force
Write-Host ""
Write-Host "  [STEP 1] Current state backed up to:"
Write-Host "    $baBakNow"
Write-Host "    $hcaBakNow"

# === Find latest pre-Patch-L backup (= Patch I+H state) ===
$baBakBaseline  = Get-ChildItem -Path "$baPath.bak-prePatchL-*"  -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
$hcaBakBaseline = Get-ChildItem -Path "$hcaPath.bak-prePatchL-*" -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $baBakBaseline -or -not $hcaBakBaseline) {
    throw "[FAIL] Pre-Patch-L (Patch I+H) baseline backup not found!"
}

Write-Host ""
Write-Host "  [STEP 2] Found Patch I+H baseline:"
Write-Host "    BA:  $($baBakBaseline.Name)"
Write-Host "    HCA: $($hcaBakBaseline.Name)"

# === Restore ===
Copy-Item -LiteralPath $baBakBaseline.FullName  -Destination $baPath  -Force
Copy-Item -LiteralPath $hcaBakBaseline.FullName -Destination $hcaPath -Force
Write-Host ""
Write-Host "  [STEP 3] Restored:" -ForegroundColor Green
Write-Host "    $baPath  <- Patch I+H state"
Write-Host "    $hcaPath <- Patch I+H state"

# === Verify ===
$enc = [System.Text.Encoding]::GetEncoding(874)
$baText  = [System.IO.File]::ReadAllText($baPath,  $enc)
$hcaText = [System.IO.File]::ReadAllText($hcaPath, $enc)

Write-Host ""
Write-Host "  [STEP 4] Verification (should match Patch I+H state):" -ForegroundColor Yellow

function CountAndShow {
    param([string]$Text, [string]$Label, [string]$Pattern, [int]$Expected)
    $escaped = [regex]::Escape($Pattern)
    $count = ([regex]::Matches($Text, $escaped)).Count
    $status = if ($count -eq $Expected) { '[OK]' } else { '[!!]' }
    $color  = if ($count -eq $Expected) { 'Green' } else { 'Red' }
    Write-Host ("    {0} {1} '{2,-22}' : {3} (expect {4})" -f $status, $Label, $Pattern, $count, $Expected) -ForegroundColor $color
}

# Patch I+H state expects:
# - Step = Rand(-2, 2) for tt, tb, LToe + 6 Steel = 9 (BA), 9 (HCA)
# - Step = Rand(-5, 5) for TBase = 1
# - Step = Rand(-1, 1) for Base = 1
# - Init: Currenttb = Maxtb (deterministic), no Rand(...) in init
CountAndShow -Text $baText  -Label 'BA ' -Pattern 'Step = Rand(-2, 2)' -Expected 9
CountAndShow -Text $baText  -Label 'BA ' -Pattern 'Step = Rand(-5, 5)' -Expected 1
CountAndShow -Text $baText  -Label 'BA ' -Pattern 'Step = Rand(-1, 1)' -Expected 1
CountAndShow -Text $baText  -Label 'BA ' -Pattern 'Currenttb = Maxtb'  -Expected 1
CountAndShow -Text $hcaText -Label 'HCA' -Pattern 'Step = Rand(-2, 2)' -Expected 9
CountAndShow -Text $hcaText -Label 'HCA' -Pattern 'Step = Rand(-5, 5)' -Expected 1
CountAndShow -Text $hcaText -Label 'HCA' -Pattern 'Step = Rand(-1, 1)' -Expected 1
CountAndShow -Text $hcaText -Label 'HCA' -Pattern 'CurrentStemDB = DB_MAX' -Expected 1

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Rollback to Patch I+H baseline complete"               -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next: Run script 2 (Apply-PatchP2-AddDebugPrint.ps1)"
Write-Host "        to add detailed cost breakdown logs"
Write-Host ""
