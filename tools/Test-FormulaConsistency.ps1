$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
[xml]$ribbon = Get-Content -LiteralPath (Join-Path $repository 'src\FormuRay.AddIn\Ribbon\FormuRayRibbon.xml') -Raw
$namespaces = [System.Xml.XmlNamespaceManager]::new($ribbon.NameTable)
$namespaces.AddNamespace('m', 'http://schemas.microsoft.com/office/2009/07/customui')
$hardcode = $ribbon.SelectSingleNode('//m:toggleButton[@id="tglHardcodedDetectorVsto"]', $namespaces)
$consistency = if ($hardcode) { $hardcode.SelectSingleNode('following-sibling::*[1]', $namespaces) } else { $null }
if (-not $consistency -or $consistency.LocalName -ne 'toggleButton' -or
    $consistency.id -ne 'tglFormulaConsistencyVsto' -or
    $consistency.onAction -ne 'OnFormulaConsistency' -or
    $consistency.getPressed -ne 'GetFormulaConsistencyPressed') {
    throw 'Formula Consistency Ribbon toggle must immediately follow HardCode Detector.'
}
Write-Output 'PASS Formula Consistency Ribbon toggle position and callbacks'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -property installationPath
if (-not $installation) { throw 'Visual Studio was not found.' }
$compiler = Join-Path $installation 'MSBuild\Current\Bin\Roslyn\csc.exe'
$pia = Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll'
$output = Join-Path ([IO.Path]::GetTempPath()) ('FormuRay-FormulaConsistencySmoke-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output | Out-Null
foreach ($project in @('FormuRay.Core', 'FormuRay.Excel')) {
    Copy-Item -LiteralPath (Join-Path $repository "src\$project\bin\Debug\$project.dll") -Destination $output
}
$arguments = @('/nologo', '/target:exe', '/platform:anycpu', "/out:$output\FormulaConsistencySmoke.exe",
    "/link:$pia", '/reference:System.Drawing.dll',
    "/reference:$output\FormuRay.Core.dll", "/reference:$output\FormuRay.Excel.dll",
    (Join-Path $repository 'tests\FormuRay.Excel.Smoke\FormulaConsistencySmoke.cs'))
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw 'Formula Consistency smoke runner compilation failed.' }
& (Join-Path $output 'FormulaConsistencySmoke.exe')
$result = $LASTEXITCODE
Write-Output "Smoke artifacts: $output"
if ($result -ne 0) { throw 'Formula Consistency Excel smoke checks failed.' }
