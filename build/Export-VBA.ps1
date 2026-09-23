[CmdletBinding()]
param(
    [string]$InputWorkbook = 'E:\TaiHoaDonDienTu_v6.3.xlsm',
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force

$InputWorkbook = (Resolve-Path -LiteralPath $InputWorkbook).Path
$templateDir = Join-Path $RepositoryRoot 'template'
$moduleDir = Join-Path $RepositoryRoot 'src\modules'
$classDir = Join-Path $RepositoryRoot 'src\classes'
$formDir = Join-Path $RepositoryRoot 'src\forms'
$testDir = Join-Path $RepositoryRoot 'tests'
foreach ($dir in @($templateDir, $moduleDir, $classDir, $formDir, $testDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

Get-ChildItem -LiteralPath $moduleDir -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $classDir -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $formDir -File -ErrorAction SilentlyContinue | Remove-Item -Force

$templatePath = Join-Path $templateDir 'App_Template.xlsm'
Copy-Item -LiteralPath $InputWorkbook -Destination $templatePath -Force

Invoke-WithExcelWorkbook -Path $InputWorkbook -ReadOnly -Action {
    param($workbook, $excel)
    foreach ($component in @($workbook.VBProject.VBComponents)) {
        $extension = $null
        $destination = $null
        switch ([int]$component.Type) {
            1 { $extension = '.bas'; $destination = $moduleDir }
            2 { $extension = '.cls'; $destination = $classDir }
            3 { $extension = '.frm'; $destination = $formDir }
            default { }
        }
        if ($null -ne $extension) {
            $exportPath = Join-Path $destination ($component.Name + $extension)
            $component.Export($exportPath)
        }
        Release-ComObject $component
    }
    $inventory = Get-WorkbookInventory -Workbook $workbook -Excel $excel -Path $InputWorkbook
    $inventory | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $testDir 'original-inventory.json') -Encoding utf8
}

$missingFrx = @()
Get-ChildItem -LiteralPath $formDir -Filter '*.frm' | ForEach-Object {
    $frx = [IO.Path]::ChangeExtension($_.FullName, '.frx')
    if (-not (Test-Path -LiteralPath $frx)) { $missingFrx += $_.Name }
}
if ($missingFrx.Count -gt 0) { throw "Missing FRX files for: $($missingFrx -join ', ')" }
Write-Host "Exported VBA source and inventory. Template: $templatePath"
