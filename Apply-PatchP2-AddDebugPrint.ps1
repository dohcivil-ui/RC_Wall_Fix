# ============================================================
# Apply-PatchP2-AddDebugPrint.ps1
# Patch P-2: Add detailed Debug.Print for initial cost breakdown
# Date: 2026-05-12
# Run AFTER Apply-PatchP1-RollbackToPatchIH.ps1
#
# Adds 12 Debug.Print lines each to BA and HCA showing:
#   - Indices and meter-values of all design variables
#   - modShared.H, cover
#   - Material prices
#   - Initial cost
#   - Cost breakdown: H_stem, V_stem, V_base
#
# These lines print to Immediate Window with prefix [P2_BA] or [P2_HCA]
# Compare BA and HCA output side-by-side to find bug.
#
# Run:
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchP2-AddDebugPrint.ps1"
# ============================================================

$ErrorActionPreference = 'Stop'

$srcDir  = 'D:\rc-rt-optimize-v2\vb6-source'
$baPath  = Join-Path $srcDir 'modBA.bas'
$hcaPath = Join-Path $srcDir 'modHillClimbing.bas'
$enc = [System.Text.Encoding]::GetEncoding(874)
$ts  = Get-Date -Format 'yyyyMMdd-HHmmss'

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch P-2: Add Debug.Print for initial cost diagnosis"     -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# === Backup ===
$baBak  = "$baPath.bak-prePatchP2-$ts"
$hcaBak = "$hcaPath.bak-prePatchP2-$ts"
Copy-Item -LiteralPath $baPath  -Destination $baBak  -Force
Copy-Item -LiteralPath $hcaPath -Destination $hcaBak -Force
Write-Host ""
Write-Host "  Backups:" -ForegroundColor Yellow
Write-Host "    $baBak"
Write-Host "    $hcaBak"

# ============================================================
# BA: Insert debug lines after "tb Bisection: Min=..." Debug.Print line
# ============================================================
$baOld = "    Debug.Print `"tb Bisection: Min=`" & Mintb & `", Max=`" & Maxtb & `", Mid=`" & Midtb`r`n"
$baDebugBlock = @"
    Debug.Print "[P2_BA] ttIdx="    & Currenttt    & ", tt_m="    & Format(WP_tt(Currenttt), "0.0000")
    Debug.Print "[P2_BA] tbIdx="    & Currenttb    & ", tb_m="    & Format(WP_tb(Currenttb), "0.0000")
    Debug.Print "[P2_BA] TBIdx="    & CurrentTBase & ", TB_m="    & Format(WP_TBase(CurrentTBase), "0.0000")
    Debug.Print "[P2_BA] BaseIdx="  & CurrentBase  & ", Base_m="  & Format(WP_Base(CurrentBase), "0.0000")
    Debug.Print "[P2_BA] LToeIdx="  & CurrentLToe  & ", LToe_m="  & Format(WP_LToe(CurrentLToe), "0.0000")
    Debug.Print "[P2_BA] LHeel_m="  & Format(current.LHeel, "0.0000")
    Debug.Print "[P2_BA] StemDB="   & WP_DB(CurrentStemDB)   & "mm, StemSP="   & Format(WP_SP(CurrentStemSP), "0.000") & "m"
    Debug.Print "[P2_BA] ToeDB="    & WP_DB(CurrentToeDB)    & "mm, ToeSP="    & Format(WP_SP(CurrentToeSP), "0.000")  & "m"
    Debug.Print "[P2_BA] HeelDB="   & WP_DB(CurrentHeelDB)   & "mm, HeelSP="   & Format(WP_SP(CurrentHeelSP), "0.000") & "m"
    Debug.Print "[P2_BA] modShared.H=" & modShared.H & ", cover=" & modShared.cover
    Debug.Print "[P2_BA] cP=" & modShared.currentMaterial.concretePrice & ", sP=" & modShared.currentMaterial.steelPrice
    Dim p2H_stem As Double, p2V_stem As Double, p2V_base As Double
    p2H_stem = modShared.H - WP_TBase(CurrentTBase)
    p2V_stem = 0.5 * (WP_tt(Currenttt) + WP_tb(Currenttb)) * p2H_stem
    p2V_base = WP_Base(CurrentBase) * WP_TBase(CurrentTBase)
    Debug.Print "[P2_BA] H_stem=" & Format(p2H_stem, "0.0000") & ", V_stem=" & Format(p2V_stem, "0.0000") & ", V_base=" & Format(p2V_base, "0.0000")
"@

$baNew = $baOld + $baDebugBlock + "`r`n"

$baText = [System.IO.File]::ReadAllText($baPath, $enc)
if (-not $baText.Contains($baOld)) {
    throw "[FAIL] BA anchor pattern not found"
}
$baText = $baText.Replace($baOld, $baNew)
[System.IO.File]::WriteAllText($baPath, $baText, $enc)
Write-Host ""
Write-Host "  [OK] BA debug block inserted" -ForegroundColor Green

# ============================================================
# HCA: Insert debug lines after "LHeel=..." Debug.Print line
# Anchor uniquely on the variable declaration after it
# ============================================================
$hcaOld = "    Dim estimatedCost As Double`r`n"
$hcaDebugBlock = @"
    Debug.Print "[P2_HCA] ttIdx="    & Currenttt    & ", tt_m="    & Format(WP_tt(Currenttt), "0.0000")
    Debug.Print "[P2_HCA] tbIdx="    & Currenttb    & ", tb_m="    & Format(WP_tb(Currenttb), "0.0000")
    Debug.Print "[P2_HCA] TBIdx="    & CurrentTBase & ", TB_m="    & Format(WP_TBase(CurrentTBase), "0.0000")
    Debug.Print "[P2_HCA] BaseIdx="  & CurrentBase  & ", Base_m="  & Format(WP_Base(CurrentBase), "0.0000")
    Debug.Print "[P2_HCA] LToeIdx="  & CurrentLToe  & ", LToe_m="  & Format(WP_LToe(CurrentLToe), "0.0000")
    Debug.Print "[P2_HCA] LHeel_m="  & Format(current.LHeel, "0.0000")
    Debug.Print "[P2_HCA] StemDB="   & WP_DB(CurrentStemDB)   & "mm, StemSP="   & Format(WP_SP(CurrentStemSP), "0.000") & "m"
    Debug.Print "[P2_HCA] ToeDB="    & WP_DB(CurrentToeDB)    & "mm, ToeSP="    & Format(WP_SP(CurrentToeSP), "0.000")  & "m"
    Debug.Print "[P2_HCA] HeelDB="   & WP_DB(CurrentHeelDB)   & "mm, HeelSP="   & Format(WP_SP(CurrentHeelSP), "0.000") & "m"
    Debug.Print "[P2_HCA] modShared.H=" & modShared.H & ", cover=" & modShared.cover
    Debug.Print "[P2_HCA] cP=" & modShared.currentMaterial.concretePrice & ", sP=" & modShared.currentMaterial.steelPrice
    Dim p2H_stem As Double, p2V_stem As Double, p2V_base As Double
    p2H_stem = modShared.H - WP_TBase(CurrentTBase)
    p2V_stem = 0.5 * (WP_tt(Currenttt) + WP_tb(Currenttb)) * p2H_stem
    p2V_base = WP_Base(CurrentBase) * WP_TBase(CurrentTBase)
    Debug.Print "[P2_HCA] H_stem=" & Format(p2H_stem, "0.0000") & ", V_stem=" & Format(p2V_stem, "0.0000") & ", V_base=" & Format(p2V_base, "0.0000")
"@

$hcaNew = $hcaOld + $hcaDebugBlock + "`r`n"

$hcaText = [System.IO.File]::ReadAllText($hcaPath, $enc)
if (-not $hcaText.Contains($hcaOld)) {
    throw "[FAIL] HCA anchor pattern not found"
}
$hcaText = $hcaText.Replace($hcaOld, $hcaNew)
[System.IO.File]::WriteAllText($hcaPath, $hcaText, $enc)
Write-Host "  [OK] HCA debug block inserted" -ForegroundColor Green

# ============================================================
# Verify
# ============================================================
$baText  = [System.IO.File]::ReadAllText($baPath,  $enc)
$hcaText = [System.IO.File]::ReadAllText($hcaPath, $enc)
$baCount  = ([regex]::Matches($baText,  '\[P2_BA\]')).Count
$hcaCount = ([regex]::Matches($hcaText, '\[P2_HCA\]')).Count

Write-Host ""
Write-Host "  Verification:" -ForegroundColor Yellow
Write-Host "    BA  [P2_BA] markers: $baCount  (expect 12)"
Write-Host "    HCA [P2_HCA] markers: $hcaCount (expect 12)"

if ($baCount -ne 12 -or $hcaCount -ne 12) {
    Write-Host "  [!!] Marker counts differ from expected" -ForegroundColor Red
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Debug print added"                                     -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:"
Write-Host "    1. Open VB6 IDE"
Write-Host "    2. View -> Immediate Window (Ctrl+G)"
Write-Host "    3. Set Number of Trials = 1, MaxIter = 100"
Write-Host "    4. Click BA  -> watch Immediate Window for [P2_BA] lines"
Write-Host "    5. Click HCA -> watch Immediate Window for [P2_HCA] lines"
Write-Host "    6. Copy both sets of lines and send to chat"
Write-Host ""
