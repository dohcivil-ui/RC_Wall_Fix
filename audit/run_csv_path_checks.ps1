$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)
$evidence = Join-Path $PSScriptRoot ('csv-path-checks\' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff', [Globalization.CultureInfo]::InvariantCulture))
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
$encoding = [Text.Encoding]::GetEncoding(28591)
$project = [IO.File]::ReadAllText((Join-Path $PWD 'RC_RT_HCA_v2.vbp'), $encoding)
$project = $project.Replace('Startup="Form1"', 'Startup="Sub Main"').Replace('Name="Project1"', 'Name="CsvPathsTest"')
$project = [regex]::Replace($project, '(?m)^(Module=[^;]+; |Form=)([^\r\n]+)', '$1..\$2')
$project += "`r`nModule=CsvPathRegression; CsvPathRegression.bas`r`nExeName32=`"CsvPathRegression.exe`"`r`n"
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'CsvPathRegression.vbp'), $project, $encoding)
foreach ($name in @('RC_RT_HCA_v2', 'CsvPathRegression')) {
    $vbp = if ($name -eq 'RC_RT_HCA_v2') { Join-Path $PWD "$name.vbp" } else { Join-Path $PSScriptRoot "$name.vbp" }
    $log = Join-Path $evidence "compile-$name.log"
    $arguments = '/make "{0}" /out "{1}" /outdir "{2}"' -f $vbp,$log,$PSScriptRoot
    $build = Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList $arguments -WindowStyle Hidden -PassThru
    if (-not $build.WaitForExit(30000)) { throw 'Compile timeout' }
    $text = Get-Content -Raw -LiteralPath $log
    if ($text -notmatch 'succeeded' -or $text -match 'failed|Compile Error') { throw $text }
    Write-Output $text.Trim()
}
$test = Start-Process -FilePath (Join-Path $PSScriptRoot 'CsvPathRegression.exe') -ArgumentList ('"{0}"' -f $evidence) -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru
if (-not $test.WaitForExit(30000)) { throw 'CSV path test timeout' }
$result = Get-Content -Raw -LiteralPath (Join-Path $evidence 'native-csv-paths.txt')
Write-Output $result
if ($result -notmatch 'CSV_PATH checks=\d+; failures=0' -or $result -match 'FAIL:|FATAL') { throw 'CSV path verification failed' }
Get-FileHash -Algorithm SHA256 -LiteralPath 'Form1.frm','modShared.bas','modBatch.bas','audit\RC_RT_HCA_v2.exe','audit\CsvPathRegression.exe' | Format-List | Out-File (Join-Path $evidence 'hashes.txt') -Encoding utf8
Write-Output "Evidence: $evidence"
