$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)
$projectEvidence = Join-Path $PSScriptRoot ('project-checks\' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff', [Globalization.CultureInfo]::InvariantCulture))
New-Item -ItemType Directory -Path $projectEvidence -Force | Out-Null
# Existing regression uses explicit historical/synthetic mode; new tests use production mode.
& powershell -NoProfile -ExecutionPolicy Bypass -File audit\run_checks.ps1
if ($LASTEXITCODE -ne 0) { throw 'Existing regression failed' }
python audit\build_project_regression.py
if ($LASTEXITCODE -ne 0) { throw 'Project fixture generation failed' }
$projectArgs = '/make "{0}\ProjectRegression.vbp" /out "{1}\compile-project.log" /outdir "{0}"' -f $PSScriptRoot,$projectEvidence
$projectBuild = Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList $projectArgs -WindowStyle Hidden -PassThru
if (-not $projectBuild.WaitForExit(30000)) { throw 'Project compile timeout' }
$projectLog = Get-Content -Raw -LiteralPath (Join-Path $projectEvidence 'compile-project.log')
if ($projectLog -notmatch 'succeeded' -or $projectLog -match 'failed|Compile Error') { throw 'Project compile failed' }
$projectTest = Start-Process -FilePath (Join-Path $PSScriptRoot 'ProjectRegression.exe') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
if (-not $projectTest.WaitForExit(30000)) { throw 'Project test timeout' }
Get-ChildItem -LiteralPath $PSScriptRoot -File | Where-Object { $_.Name -match '^project-(regression\.txt|fixtures\.csv|(H[345]|anchorage_excluded|secondary_excluded)-(detail\.csv|report\.txt))$' } | Copy-Item -Destination $projectEvidence
$projectNative = Get-Content -Raw -LiteralPath (Join-Path $projectEvidence 'project-regression.txt')
Write-Output $projectNative
if ($projectNative -match 'FATAL|FAIL:' -or $projectNative -notmatch 'PROJECT checks=\d+; failures=0') { throw 'Project native verification failed' }
python audit\verify_project_checks.py $projectEvidence
if ($LASTEXITCODE -ne 0) { throw 'Independent project verification failed' }
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'ProjectRegression.exe'),(Join-Path $PSScriptRoot 'RC_RT_HCA_v2.exe'),(Join-Path $PSScriptRoot '..\modProjectChecks.bas'),(Join-Path $PSScriptRoot '..\modShared.bas'),(Join-Path $PSScriptRoot '..\modWSD.bas'),(Join-Path $PSScriptRoot '..\Form1.frm') | Format-List | Out-File (Join-Path $projectEvidence 'hashes.txt') -Encoding utf8
python audit\source_checks.py
if ($LASTEXITCODE -ne 0) { throw 'Source/trace verification failed' }
Write-Output "Evidence: $projectEvidence"
