# ============================================================
# Apply-PatchL-StepExpansion.ps1
# Patch L: Step Expansion Moderate (Level A)
# Date: 2026-05-12
# Target: modBA.bas + modHillClimbing.bas
# Description: Increase step ranges ~2x for all 11 step variables
#              Applied symmetrically to BA and HCA (fair-init)
#
# Run from any directory (uses absolute paths):
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchL-StepExpansion.ps1"
# Or from VB6 source folder:
#   cd D:\rc-rt-optimize-v2\vb6-source
#   .\Apply-PatchL-StepExpansion.ps1
# ============================================================

$ErrorActionPreference = 'Stop'

# === Paths (absolute, .NET-compatible) ===
$srcDir  = 'D:\rc-rt-optimize-v2\vb6-source'
$baPath  = Join-Path $srcDir 'modBA.bas'
$hcaPath = Join-Path $srcDir 'modHillClimbing.bas'

# === Encoding: Windows-874 (CP874) — preserve Thai characters in comments ===
$enc = [System.Text.Encoding]::GetEncoding(874)

# === Timestamp for backup naming ===
$ts = Get-Date -Format 'yyyyMMdd-HHmmss'

# === Pre-flight checks ===
foreach ($p in @($baPath, $hcaPath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        throw "[FAIL] File not found: $p"
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch L: Step Expansion Moderate (Level A)"               -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Targets:"
Write-Host "    $baPath"
Write-Host "    $hcaPath"
Write-Host ""

# === Backup ===
$baBak  = "$baPath.bak-prePatchL-$ts"
$hcaBak = "$hcaPath.bak-prePatchL-$ts"
Copy-Item -LiteralPath $baPath  -Destination $baBak  -Force
Copy-Item -LiteralPath $hcaPath -Destination $hcaBak -Force
Write-Host "  [OK] Backups created:" -ForegroundColor Green
Write-Host "    $baBak"
Write-Host "    $hcaBak"
Write-Host ""

# === Replacement table (same for both algorithms) ===
# Each entry: { Tag, Old, New } — uses next-line anchor for uniqueness
$replaces = @(
    # --- Size variables ---
    @{ Tag='tt    '; Old="Step = Rand(-2, 2)`r`n    Newtt = Currenttt + Step";          New="Step = Rand(-4, 4)`r`n    Newtt = Currenttt + Step" }
    @{ Tag='tb    '; Old="Step = Rand(-2, 2)`r`n    Newtb = Currenttb + Step";          New="Step = Rand(-4, 4)`r`n    Newtb = Currenttb + Step" }
    @{ Tag='TBase '; Old="Step = Rand(-5, 5)`r`n    NewTBase = CurrentTBase + Step";    New="Step = Rand(-7, 7)`r`n    NewTBase = CurrentTBase + Step" }
    @{ Tag='LToe  '; Old="Step = Rand(-2, 2)`r`n    NewLToe = CurrentLToe + Step";      New="Step = Rand(-4, 4)`r`n    NewLToe = CurrentLToe + Step" }
    @{ Tag='Base  '; Old="Step = Rand(-1, 1)`r`n    NewBase = CurrentBase + Step";      New="Step = Rand(-3, 3)`r`n    NewBase = CurrentBase + Step" }
    # --- Steel x6 (Stem DB+SP, Toe DB+SP, Heel DB+SP) ---
    @{ Tag='StemDB'; Old="Step = Rand(-2, 2)`r`n    NewStemDB = CurrentStemDB + Step";  New="Step = Rand(-3, 3)`r`n    NewStemDB = CurrentStemDB + Step" }
    @{ Tag='StemSP'; Old="Step = Rand(-2, 2)`r`n    NewStemSP = CurrentStemSP + Step";  New="Step = Rand(-3, 3)`r`n    NewStemSP = CurrentStemSP + Step" }
    @{ Tag='ToeDB '; Old="Step = Rand(-2, 2)`r`n    NewToeDB = CurrentToeDB + Step";    New="Step = Rand(-3, 3)`r`n    NewToeDB = CurrentToeDB + Step" }
    @{ Tag='ToeSP '; Old="Step = Rand(-2, 2)`r`n    NewToeSP = CurrentToeSP + Step";    New="Step = Rand(-3, 3)`r`n    NewToeSP = CurrentToeSP + Step" }
    @{ Tag='HeelDB'; Old="Step = Rand(-2, 2)`r`n    NewHeelDB = CurrentHeelDB + Step";  New="Step = Rand(-3, 3)`r`n    NewHeelDB = CurrentHeelDB + Step" }
    @{ Tag='HeelSP'; Old="Step = Rand(-2, 2)`r`n    NewHeelSP = CurrentHeelSP + Step";  New="Step = Rand(-3, 3)`r`n    NewHeelSP = CurrentHeelSP + Step" }
)

# === Apply function ===
function Apply-To-File {
    param([string]$Path, [string]$Label)

    Write-Host "  --- $Label ---" -ForegroundColor Yellow
    $text = [System.IO.File]::ReadAllText($Path, $enc)
    $count = 0

    foreach ($r in $replaces) {
        if (-not $text.Contains($r.Old)) {
            throw "[FAIL] $Label : pattern not found for variable '$($r.Tag.Trim())'"
        }
        $text = $text.Replace($r.Old, $r.New)
        Write-Host "    [OK] $($r.Tag) : Rand changed"
        $count++
    }

    [System.IO.File]::WriteAllText($Path, $text, $enc)
    Write-Host "    Saved: $Path ($count replacements)" -ForegroundColor Green
    Write-Host ""
}

# === Run ===
Apply-To-File -Path $baPath  -Label 'BA  (modBA.bas)'
Apply-To-File -Path $hcaPath -Label 'HCA (modHillClimbing.bas)'

# === Verification: count post-patch patterns ===
Write-Host "  === Verification (post-patch) ===" -ForegroundColor Cyan
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
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-4, 4)' -Expected 3 -Note 'tt, tb, LToe'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-7, 7)' -Expected 1 -Note 'TBase'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-3, 3)' -Expected 7 -Note 'Base + Steel x6'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-2, 2)' -Expected 0 -Note 'all replaced'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-1, 1)' -Expected 0 -Note 'all replaced'
    Show-Count -Path $pair.Path -Label $pair.Label -Pattern 'Step = Rand(-5, 5)' -Expected 0 -Note 'all replaced'
    Write-Host ""
}

Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Patch L applied successfully"                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Note: Comment headers in source files still reference old"
Write-Host "        Rand values. Code is correctly updated; comments are"
Write-Host "        cosmetic and can be updated manually later."
Write-Host ""
Write-Host "  Next steps:"
Write-Host "    1. Open VB6 IDE -> reload modBA.bas, modHillClimbing.bas"
Write-Host "    2. Compile (Ctrl+F5) -> Run 30 trials @ H=3 (BA, then HCA)"
Write-Host "    3. Save CSVs:"
Write-Host "         loopPrice-BA-H3-patchL.csv"
Write-Host "         loopPrice-HCA-H3-patchL.csv"
Write-Host "    4. Send CSVs back to chat for comparison vs Patch I+H baseline"
Write-Host ""
Write-Host "  Rollback (if needed):"
Write-Host "    Copy-Item '$baBak'  '$baPath'  -Force"
Write-Host "    Copy-Item '$hcaBak' '$hcaPath' -Force"
Write-Host ""
