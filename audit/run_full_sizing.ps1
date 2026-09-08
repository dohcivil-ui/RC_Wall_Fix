$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $taskRoot
$evidence = Join-Path $PSScriptRoot ('full-sizing-checks\' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
python audit\build_full_sizing.py
if ($LASTEXITCODE -ne 0) { throw 'Native harness generation failed' }
$compileArgs = '/make "{0}\FullSizing.vbp" /out "{1}\compile.log" /outdir "{0}"' -f $PSScriptRoot,$evidence
$process = Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList $compileArgs -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(30000)) { throw 'Compile timeout' }
$log = Get-Content -Raw -LiteralPath (Join-Path $evidence 'compile.log')
if ($log -notmatch 'succeeded' -or $log -match 'failed|Compile Error') { throw 'Compile failed' }
$process = Start-Process -FilePath (Join-Path $PSScriptRoot 'FullSizing.exe') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
while (-not $process.WaitForExit(10000)) { Write-Output 'Actual VB6 recalculation is running' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'full-sizing-vb6.md'), (Join-Path $PSScriptRoot 'full-sizing-summary.csv'), (Join-Path $PSScriptRoot 'full-sizing-bar-options.csv'), (Join-Path $PSScriptRoot 'full-sizing-improvements.csv'), (Join-Path $PSScriptRoot 'full-sizing-inputs.json') -Destination $evidence
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'FullSizing.exe'), (Join-Path $PSScriptRoot 'FullSizingMain.bas'), (Join-Path $taskRoot 'modShared.bas'), (Join-Path $taskRoot 'modWSD.bas'), (Join-Path $taskRoot 'modBA.bas') | Format-List | Out-File (Join-Path $evidence 'hashes.txt')
python audit\verify_full_sizing.py
if ($LASTEXITCODE -ne 0) { throw 'Independent verification found a discrepancy; inspect native evidence' }
python audit\verify_recalculation_details.py
if ($LASTEXITCODE -ne 0) { throw 'Member force/cost reconciliation failed' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'recalculation-details.json') -Destination $evidence
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'full-sizing-independent-checks.txt'), (Join-Path $PSScriptRoot 'full-sizing-boundary-cases.csv'), (Join-Path $PSScriptRoot 'full-sizing-flags-0.bin.gz'), (Join-Path $PSScriptRoot 'full-sizing-flags-1.bin.gz'), (Join-Path $PSScriptRoot 'full-sizing-stem-profiles.json') -Destination $evidence
Write-Output "Actual VB6 evidence: $evidence"
