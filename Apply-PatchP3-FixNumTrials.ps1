# ============================================================
# Apply-PatchP3-FixNumTrials.ps1
# Patch P-3: Fix BA's hardcoded numTrials = 30 in Form1.frm
# Date: 2026-05-12
#
# Bug: cmdBA_Click in Form1.frm sets numTrials = 30 (hardcoded)
#      instead of reading from txtTrials.Text like HCA does.
#      Result: BA ignores user's trial count setting.
#
# Fix: Change to numTrials = CInt(txtTrials.Text) (same as HCA)
#
# Run:
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchP3-FixNumTrials.ps1"
# ============================================================

$ErrorActionPreference = 'Stop'

$srcDir   = 'D:\rc-rt-optimize-v2\vb6-source'
$formPath = Join-Path $srcDir 'Form1.frm'
$enc = [System.Text.Encoding]::GetEncoding(874)
$ts  = Get-Date -Format 'yyyyMMdd-HHmmss'

if (-not (Test-Path -LiteralPath $formPath)) {
    throw "[FAIL] Form1.frm not found at $formPath"
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch P-3: Fix BA numTrials hardcoded bug"                  -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Backup
$formBak = "$formPath.bak-prePatchP3-$ts"
Copy-Item -LiteralPath $formPath -Destination $formBak -Force
Write-Host ""
Write-Host "  Backup: $formBak"

# Apply
$text = [System.IO.File]::ReadAllText($formPath, $enc)

# Use the unique 2-line pattern (maxIter line + numTrials line)
$old = "    maxIter = CLng(txtMaxIter.Text)`r`n    numTrials = 30`r`n"
$new = "    maxIter = CLng(txtMaxIter.Text)`r`n    numTrials = CInt(txtTrials.Text)`r`n"

if (-not $text.Contains($old)) {
    throw "[FAIL] Pattern not found. Maybe already patched?"
}
$text = $text.Replace($old, $new)
[System.IO.File]::WriteAllText($formPath, $text, $enc)

Write-Host ""
Write-Host "  [OK] BA numTrials now reads from txtTrials.Text" -ForegroundColor Green

# Verify
$check = [System.IO.File]::ReadAllText($formPath, $enc)
$count_old = ([regex]::Matches($check, '    numTrials = 30')).Count
$count_new = ([regex]::Matches($check, 'numTrials = CInt\(txtTrials\.Text\)')).Count

Write-Host ""
Write-Host "  Verification:" -ForegroundColor Yellow
Write-Host "    'numTrials = 30' remaining   : $count_old (expect 0)"
Write-Host "    'CInt(txtTrials.Text)' calls : $count_new (expect 2 - BA and HCA)"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Patch P-3 applied successfully"                        -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next: Open VB6 IDE -> reload Form1.frm -> Compile"
Write-Host "        Test by setting Number of Trials = 1, click BA"
Write-Host "        Should see only Trial 1 in Immediate Window"
Write-Host ""
