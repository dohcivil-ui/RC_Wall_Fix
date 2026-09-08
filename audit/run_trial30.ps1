$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $taskRoot
$evidence = Join-Path $PSScriptRoot ('trial30-checks\' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
$existingFolders = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'results') -Directory | Select-Object -ExpandProperty Name)
$existingFolders | ConvertTo-Json | Out-File (Join-Path $evidence 'prior-folders.json') -Encoding utf8
python audit\build_trial30.py
if ($LASTEXITCODE -ne 0) { throw 'Trial harness generation failed' }
$compileArgs = '/make "{0}\Trial30.vbp" /out "{1}\compile.log" /outdir "{0}"' -f $PSScriptRoot,$evidence
$process = Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList $compileArgs -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(30000)) { throw 'Compile timeout' }
$log = Get-Content -Raw -LiteralPath (Join-Path $evidence 'compile.log')
if ($log -notmatch 'succeeded' -or $log -match 'failed|Compile Error') { throw 'Compile failed' }
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'Trial30.exe'), (Join-Path $PSScriptRoot 'Trial30.bas'), (Join-Path $taskRoot 'Form1.frm'), (Join-Path $taskRoot 'modShared.bas'), (Join-Path $taskRoot 'modWSD.bas'), (Join-Path $taskRoot 'modBA.bas') | Format-List | Out-File (Join-Path $evidence 'hashes.txt')
$process = Start-Process -FilePath (Join-Path $PSScriptRoot 'Trial30.exe') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
while (-not $process.WaitForExit(10000)) {
    $completed = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'results') -Directory | Where-Object { $_.Name -notin $existingFolders -and (Test-Path -LiteralPath (Join-Path $_.FullName 'run.txt')) })
    Write-Output "VB6 BA trials completed: $($completed.Count)/30"
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'trial30-native.txt') -Destination $evidence
$native = Get-Content -Raw -LiteralPath (Join-Path $evidence 'trial30-native.txt')
if ($native -notmatch 'NATIVE_TRIAL30_COMPLETE' -or $native -match 'FATAL') { throw 'Native trial run incomplete; inspect evidence' }
python audit\verify_trial30.py $evidence
if ($LASTEXITCODE -ne 0) { throw 'Trial evidence validation failed' }
Write-Output "Actual VB6 trial evidence: $evidence"
