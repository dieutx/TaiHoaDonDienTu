[CmdletBinding()]
param([Parameter(Mandatory)][string]$WorkbookPath)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking

$resolvedPath = [IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $resolvedPath)) { throw "Workbook not found: $resolvedPath" }

$excel = $null
$workbook = $null
$menu = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.UserName = 'TaiHoaDonDienTu'
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($resolvedPath, 0, $false)
    $menu = $workbook.Worksheets.Item('MENU')
    $menu.Range('D7:E7').ClearContents()
    # LinkTraCuu!B:C contains the lookup tax IDs needed to select invoice links.
    # They are public routing data, not cached login/session data.
    $workbook.RemovePersonalInformation = $true

    foreach ($propertyName in @('Author', 'Last Author', 'Company', 'Manager')) {
        try { $workbook.BuiltinDocumentProperties.Item($propertyName).Value = 'TaiHoaDonDienTu' } catch {}
    }
    for ($index = $workbook.CustomDocumentProperties.Count; $index -ge 1; $index--) {
        try { $workbook.CustomDocumentProperties.Item($index).Delete() } catch {}
    }

    $workbook.Save()
    $workbook.Close($true)
    Release-ComObject $menu; $menu = $null
    Release-ComObject $workbook; $workbook = $null
    $excel.Quit()
    Release-ComObject $excel; $excel = $null
    Write-Host "SANITIZED: $resolvedPath"
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $menu
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
