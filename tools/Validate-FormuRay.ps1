[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure([string]$message) {
    $script:failures.Add($message)
}

$xmlPath = Join-Path $repoRoot 'src\customUI14.xml'
try {
    [xml]$ribbonXml = Get-Content -LiteralPath $xmlPath -Raw
}
catch {
    Add-Failure "Ribbon XML is invalid: $($_.Exception.Message)"
}

if ($null -ne $ribbonXml) {
    $namespace = [System.Xml.XmlNamespaceManager]::new($ribbonXml.NameTable)
    $namespace.AddNamespace('m', 'http://schemas.microsoft.com/office/2009/07/customui')

    $managerButton = $ribbonXml.SelectSingleNode('//m:button[@id="btnNamedRangesManager"]', $namespace)
    if ($null -eq $managerButton) {
        Add-Failure 'Named Ranges Manager Ribbon button is missing.'
    }
    elseif ($managerButton.label -ne 'Named Ranges Manager') {
        Add-Failure 'Named Ranges Manager Ribbon label is not exact.'
    }

    $callbackSource = Get-Content -LiteralPath (Join-Path $repoRoot 'src\modRibbonCallbacks.bas') -Raw
    foreach ($node in $ribbonXml.SelectNodes('//*[@onAction]', $namespace)) {
        $callback = [string]$node.onAction
        if ($callbackSource -notmatch "(?im)^Public Sub $([regex]::Escape($callback))\b") {
            Add-Failure "Ribbon callback has no matching Public Sub: $callback"
        }
    }
}

$requiredFiles = @(
    'src\modNamedRangesService.bas',
    'src\modNamedRangesManagerFormBuilder.bas',
    'src\clsNamedRangesManagerButton.cls',
    'src\frmNamedRangesManager.frm',
    'src\frmNamedRangesManager.frx',
    'src\frmFormuRay.frm',
    'src\frmFormuRay.frx',
    'src\frmFormuRayKeySettings.frm',
    'src\frmFormuRayKeySettings.frx'
)
foreach ($relativePath in $requiredFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $repoRoot $relativePath))) {
        Add-Failure "Required file is missing: $relativePath"
    }
}

$serviceSource = Get-Content -LiteralPath (Join-Path $repoRoot 'src\modNamedRangesService.bas') -Raw
$builderSource = Get-Content -LiteralPath (Join-Path $repoRoot 'src\modNamedRangesManagerFormBuilder.bas') -Raw
$formSource = Get-Content -LiteralPath (Join-Path $repoRoot 'src\frmNamedRangesManager.frm') -Raw

if ($builderSource -notmatch '(?im)^Public Sub BuildNamedRangesManagerForm\(\)\s*$') {
    Add-Failure 'Named Ranges form builder must expose a parameterless macro for the Excel macro list.'
}

foreach ($requiredText in @('Clean #REF', 'Clean #VALUE', 'Clean Hidden', 'Refresh', 'Close')) {
    if ($formSource -notmatch [regex]::Escape($requiredText)) {
        Add-Failure "Manager form is missing required UI text: $requiredText"
    }
}
foreach ($requiredText in @('targetWorkbook Is ThisWorkbook', 'targetWorkbook.IsAddin',
                             'nameObject.RefersTo', 'nameObject.RefersToLocal', 'nameObject.Value',
                             'targetWorkbook.Names', 'worksheetObject.Names', 'seen.Exists',
                             'NRM_CLEAN_REF', 'NRM_CLEAN_VALUE', 'NRM_CLEAN_HIDDEN',
                             'ChrW$(1057)', 'ChrW$(1067)', 'ChrW$(1047)', 'ChrW$(1063)')) {
    if ($serviceSource -notmatch [regex]::Escape($requiredText)) {
        Add-Failure "Names service is missing required behavior marker: $requiredText"
    }
}
if ($serviceSource -match 'ThisWorkbook\.Names') {
    Add-Failure 'Names service must not enumerate ThisWorkbook.Names.'
}

$managerFrxPath = Join-Path $repoRoot 'src\frmNamedRangesManager.frx'
if (Test-Path -LiteralPath $managerFrxPath) {
    $managerFrxText = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($managerFrxPath))
    foreach ($controlName in @('lstComponents', 'txtFormula', 'cmdRefresh', 'cmdClose', 'lblElement', 'lblValue')) {
        if ($managerFrxText -notmatch [regex]::Escape($controlName)) {
            Add-Failure "Manager FRX is missing the native base control: $controlName"
        }
    }
}

$legacyFormulaName = 'Formula' + 'Audit'
$legacyArtName = 'Art' + 'Explorer'
$legacyFormulaNameSpaced = 'Formula' + ' ' + 'Audit'
$legacyArtNameSpaced = 'Art' + ' ' + 'Explorer'
$oldNamePattern = [regex]::Escape($legacyFormulaName) + '(?:XLAM)?|' +
                  [regex]::Escape($legacyFormulaNameSpaced) + '(?: XLAM)?\b|' +
                  [regex]::Escape($legacyArtName) + '|' + [regex]::Escape($legacyArtNameSpaced)
$localArtifactPattern = '[\\/](?:\.git|\.vs|bin|obj|TestResults)(?:[\\/]|$)'
$textFiles = Get-ChildItem -LiteralPath $repoRoot -Recurse -File |
    Where-Object { $_.FullName.Substring($repoRoot.Length) -notmatch $localArtifactPattern -and $_.Extension -in '.md', '.bas', '.cls', '.frm', '.xml', '.ps1', '.txt', '.bat' }

foreach ($file in $textFiles) {
    $relativePath = $file.FullName.Substring($repoRoot.Length).TrimStart('\')
    $lineNumber = 0
    foreach ($line in Get-Content -LiteralPath $file.FullName) {
        $lineNumber++
        if ($line -notmatch $oldNamePattern) { continue }

        $allowed =
            ($relativePath -eq 'src\modLicenseRuntime.bas' -and $line -match ([regex]::Escape($legacyFormulaName + 'XL'))) -or
            ($relativePath -eq 'src\modHotkeySettings.bas' -and $line -match ([regex]::Escape($legacyArtName) + '(Precedents|Dependents)Hotkey')) -or
            ($relativePath -in @('README.md', 'CHANGELOG.md', 'docs\Licensing.md') -and $line -match ([regex]::Escape($legacyFormulaName + 'XL'))) -or
            ($relativePath -in @('src\FormuRay.AddIn\Ribbon\FormuRayRibbon.xml', 'docs\VstoAuditMigration.md') -and $line -match 'Formula Audit Explorer') -or
            ($relativePath -eq 'tools\Validate-FormuRay.ps1')

        if (-not $allowed) {
            Add-Failure "Unexpected legacy product name at ${relativePath}:$lineNumber"
        }
    }
}

$legacyFileNames = Get-ChildItem -LiteralPath $repoRoot -Recurse -Force |
    Where-Object { $_.FullName.Substring($repoRoot.Length) -notmatch $localArtifactPattern -and $_.Name -match $oldNamePattern }
foreach ($item in $legacyFileNames) {
    Add-Failure "Legacy product name remains in path: $($item.FullName)"
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Output 'FormuRay validation passed.'
