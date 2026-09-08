$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $taskRoot
$sizingEvidence = Join-Path $PSScriptRoot ('sizing-checks\' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $sizingEvidence -Force | Out-Null
python audit\build_sizing.py
if ($LASTEXITCODE -ne 0) { throw 'Sizing harness generation failed' }
$compileArgs = '/make "{0}\Sizing.vbp" /out "{1}\compile.log" /outdir "{0}"' -f $PSScriptRoot,$sizingEvidence
$process = Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList $compileArgs -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(30000)) { throw 'Sizing compile timeout' }
$log = Get-Content -Raw -LiteralPath (Join-Path $sizingEvidence 'compile.log')
if ($log -notmatch 'succeeded' -or $log -match 'failed|Compile Error') { throw 'Sizing compile failed' }
$process = Start-Process -FilePath (Join-Path $PSScriptRoot 'Sizing.exe') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(30000)) { throw 'Sizing run timeout' }
python audit\verify_sizing.py
if ($LASTEXITCODE -ne 0) { throw 'Independent sizing verification failed' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'sizing-options.csv'), (Join-Path $PSScriptRoot 'sizing-vb6-report.md'), (Join-Path $PSScriptRoot 'sizing-independent-checks.txt') -Destination $sizingEvidence
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'Sizing.exe'), (Join-Path $PSScriptRoot 'SizingMain.bas'), (Join-Path $taskRoot 'modShared.bas'), (Join-Path $taskRoot 'modWSD.bas') | Format-List | Out-File (Join-Path $sizingEvidence 'hashes.txt')
Write-Output "Actual VB6 sizing evidence: $sizingEvidence"
