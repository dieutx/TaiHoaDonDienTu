[CmdletBinding()]
param([Parameter(Mandatory)][string]$WorkbookPath)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking

$resolvedPath = [IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $resolvedPath)) { throw "Workbook not found: $resolvedPath" }

$excel = $null
$workbook = $null
$menu = $null
$lookup = $null
$reference = $null
$referenceLookup = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($resolvedPath, 0, $true)
    $menu = $workbook.Worksheets.Item('MENU')
    $lookup = $workbook.Worksheets.Item('LinkTraCuu')

    $cachedLogin = @($menu.Range('D7:E7').Value2 | ForEach-Object { [string]$_ } | Where-Object { $_.Length -gt 0 })
    if ($cachedLogin.Count -gt 0) { throw 'MENU!D7:E7 still contains cached login data.' }
    $buyerTaxIds = 0
    $sellerTaxIds = 0
    for ($row = 2; $row -le 122; $row++) {
        if (-not [string]::IsNullOrWhiteSpace([string]$lookup.Range("B$row").Value2)) { $buyerTaxIds++ }
        if (-not [string]::IsNullOrWhiteSpace([string]$lookup.Range("C$row").Value2)) { $sellerTaxIds++ }
    }
    if ($buyerTaxIds -ne 117 -or $sellerTaxIds -ne 4) {
        throw "LinkTraCuu lookup tax IDs changed: B=$buyerTaxIds (expected 117), C=$sellerTaxIds (expected 4)."
    }

    $templatePath = [IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $PSScriptRoot) 'template\App_Template.xlsm'))
    if ($resolvedPath -ine $templatePath) {
        $reference = $excel.Workbooks.Open($templatePath, 0, $true)
        $referenceLookup = $reference.Worksheets.Item('LinkTraCuu')
        foreach ($column in @('B', 'C')) {
            for ($row = 2; $row -le 122; $row++) {
                $address = "$column$row"
                if ([string]$lookup.Range($address).Formula -cne [string]$referenceLookup.Range($address).Formula) {
                    throw "LinkTraCuu!$address differs from the template lookup mapping."
                }
            }
        }
    }

    foreach ($sheet in @($workbook.Worksheets)) {
        $values = $sheet.UsedRange.Value2
        if ($values -is [Array]) {
            foreach ($value in $values) {
                $text = [string]$value
                if ($text -match '^eyJ[^.]+\.[^.]+\.[^.]+$') { throw "JWT-like cell remains in sheet $($sheet.Name)." }
                if ($text -match '(?i)Bearer\s+eyJ|JSESSIONID=|TS0114b13e=') { throw "Session data remains in sheet $($sheet.Name)." }
            }
        }
        Release-ComObject $sheet
    }

    Write-Host 'PUBLIC WORKBOOK TEST PASSED'
} finally {
    if ($null -ne $reference) { try { $reference.Close($false) } catch {} }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $lookup
    Release-ComObject $referenceLookup
    Release-ComObject $reference
    Release-ComObject $menu
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
