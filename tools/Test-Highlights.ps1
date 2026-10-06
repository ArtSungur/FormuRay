param()
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -property installationPath
if (-not $installation) { throw 'Visual Studio was not found.' }
$compiler = Join-Path $installation 'MSBuild\Current\Bin\Roslyn\csc.exe'
$pia = Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll'
$office = Join-Path $installation 'Common7\IDE\PrivateAssemblies\Office15\Office.dll'
$output = Join-Path ([IO.Path]::GetTempPath()) ('FormuRay-HighlightSmoke-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output | Out-Null
foreach ($project in @('FormuRay.Core', 'FormuRay.Excel')) {
    Copy-Item -LiteralPath (Join-Path $repository "src\$project\bin\Debug\$project.dll") -Destination $output
}
$arguments = @('/nologo', '/debug+', '/target:exe', '/platform:anycpu', "/out:$output\HighlightSmoke.exe",
    "/link:$pia", "/link:$office", '/reference:Microsoft.CSharp.dll', '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll',
    "/reference:$output\FormuRay.Core.dll", "/reference:$output\FormuRay.Excel.dll",
    (Join-Path $repository 'tests\FormuRay.Excel.Smoke\HighlightSmoke.cs'))
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw 'Smoke runner compilation failed.' }
$runArguments = @($repository)

& (Join-Path $output 'HighlightSmoke.exe') @runArguments
$result = $LASTEXITCODE
Write-Output "Smoke artifacts: $output"
if ($result -ne 0) { throw 'Excel smoke checks failed.' }
