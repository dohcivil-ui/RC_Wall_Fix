param([switch]$SingleTrial)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)
$evidence = Join-Path $PSScriptRoot ('sketch-checks\' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff', [Globalization.CultureInfo]::InvariantCulture))
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
if ($SingleTrial) { python audit\build_sketch_regression.py --single-trial } else { python audit\build_sketch_regression.py }
if ($LASTEXITCODE -ne 0) { throw 'Sketch harness generation failed' }
$args = '/make "{0}\SketchRegression.vbp" /out "{1}\compile-sketch.log" /outdir "{0}"' -f $PSScriptRoot,$evidence
$build = Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList $args -WindowStyle Hidden -PassThru
if (-not $build.WaitForExit(30000)) { throw 'Sketch compile timeout' }
$log = Get-Content -Raw -LiteralPath (Join-Path $evidence 'compile-sketch.log')
if ($log -notmatch 'succeeded' -or $log -match 'failed|Compile Error') { throw $log }
$test = Start-Process -FilePath (Join-Path $PSScriptRoot 'SketchRegression.exe') -ArgumentList ('"{0}"' -f $evidence) -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
if (-not $test.WaitForExit(30000)) { throw 'Sketch test timeout' }
$result = Get-Content -Raw -LiteralPath (Join-Path $evidence 'sketch-regression.txt')
Write-Output $result
if ($result -match 'FAIL:|FATAL' -or $result -notmatch 'SKETCH checks=\d+; failures=0') { throw 'Sketch verification failed' }
python audit\render_sketch_evidence.py $evidence
if ($LASTEXITCODE -ne 0) { throw 'Native image conversion failed' }
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'SketchRegression.exe'),(Join-Path $PSScriptRoot '..\frmBestDesign.frm'),(Join-Path $PSScriptRoot '..\Form1.frm') | Format-List | Out-File (Join-Path $evidence 'hashes.txt') -Encoding utf8
Write-Output "Evidence: $evidence"
