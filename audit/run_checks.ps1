$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$compiler = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE'
$runEvidence = Join-Path $PSScriptRoot ('checks\' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $runEvidence -Force | Out-Null
Set-Location -LiteralPath $projectRoot
python audit\independent_checks.py
if ($LASTEXITCODE -ne 0) { throw 'Independent checks failed' }
python audit\build_force_chain_checks.py
if ($LASTEXITCODE -ne 0) { throw 'Independent force-chain fixtures failed' }
python audit\build_reference_parameter_checks.py
if ($LASTEXITCODE -ne 0) { throw 'Reference parameter fixtures failed' }
python audit\build_harness.py
if ($LASTEXITCODE -ne 0) { throw 'Harness generation failed' }

# Derive GUI harness from the real project so its module/control list is exact.
$encoding = [Text.Encoding]::GetEncoding(28591)
$guiProject = [IO.File]::ReadAllText((Join-Path $projectRoot 'RC_RT_HCA_v2.vbp'), $encoding)
$guiProject = $guiProject.Replace('Startup="Form1"', 'Startup="Sub Main"')
$guiProject = $guiProject.Replace('Name="Project1"', 'Name="GuiTestProject"')
$guiProject = [regex]::Replace($guiProject, '(?m)^(Module=[^;]+; )([^\r\n]+)', '$1..\$2')
$guiProject = $guiProject.Replace('Form=Form1.frm', 'Form=..\Form1.frm')
$guiProject += "`r`nModule=GuiRegression; GuiRegression.bas`r`nExeName32=`"GuiRegression.exe`"`r`n"
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'GuiRegression.vbp'), $guiProject, $encoding)

foreach ($name in @('RC_RT_HCA_v2', 'Regression', 'GuiRegression')) {
    $vbpPath = if ($name -eq 'RC_RT_HCA_v2') { Join-Path $projectRoot "$name.vbp" } else { Join-Path $PSScriptRoot "$name.vbp" }
    # VB6 /out appends. A fresh path prevents an old success masking a new error.
    $logPath = Join-Path $runEvidence "compile-$name.log"
    $args = '/make "{0}" /out "{1}" /outdir "{2}"' -f $vbpPath, $logPath, $PSScriptRoot
    $process = Start-Process -FilePath $compiler -ArgumentList $args -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(30000)) { throw "Compiler still running for $name (PID $($process.Id))" }
    $logText = [IO.File]::ReadAllText($logPath)
    Write-Output $logText.Trim()
    if ($logText -notmatch 'succeeded') { throw "Compile failed: $name" }
    if ($logText -match 'failed|Compile Error') { throw "Compile failed: $name" }
    Copy-Item -LiteralPath $logPath -Destination (Join-Path $PSScriptRoot "final-compile-$name.log")
}
foreach ($name in @('Regression', 'GuiRegression')) {
    $exePath = Join-Path $PSScriptRoot "$name.exe"
    $process = Start-Process -FilePath $exePath -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(30000)) { throw "Test still running: $name (PID $($process.Id))" }
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'wsd-reference-native.csv'), (Join-Path $PSScriptRoot 'wsd-reference-independent.json') -Destination $runEvidence
$native = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'native-regression.txt')
$gui = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'gui-regression.txt')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'native-regression.txt') -Destination $runEvidence
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'gui-regression.txt') -Destination $runEvidence
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'force-chain-native.csv'), (Join-Path $PSScriptRoot 'force-chain-independent.json'), (Join-Path $PSScriptRoot 'stem-shear-native.csv'), (Join-Path $PSScriptRoot 'stem-shear-independent.json') -Destination $runEvidence
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'RC_RT_HCA_v2.exe'), (Join-Path $PSScriptRoot 'Regression.exe'), (Join-Path $PSScriptRoot 'GuiRegression.exe') | Format-List | Out-File (Join-Path $runEvidence 'binary-hashes.txt')
if ($native -notmatch 'TOTAL checks=\d+; failures=0' -or $native -match 'FATAL|FAIL:') { throw 'Native regression failed; read audit/native-regression.txt' }
if ($gui -notmatch 'GUI failures=0' -or $gui -match 'FATAL') { throw 'GUI regression failed; read audit/gui-regression.txt' }
Write-Output ($native -split "`r?`n" | Where-Object { $_ -match '^TOTAL' })
Write-Output ($gui -split "`r?`n" | Where-Object { $_ -match '^GUI failures=' })
python audit\verify_check_report.py
if ($LASTEXITCODE -ne 0) { throw 'Per-check report differs from independent calculations' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'h5-vb6-checks.md'), (Join-Path $PSScriptRoot 'report-independent-checks.txt') -Destination $runEvidence
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $projectRoot 'modShared.bas'), (Join-Path $projectRoot 'modWSD.bas'), (Join-Path $PSScriptRoot 'RegressionMain.bas') | Format-List | Out-File (Join-Path $runEvidence 'report-source-hashes.txt')
python audit\source_checks.py
if ($LASTEXITCODE -ne 0) { throw 'Source/trace checks failed' }
