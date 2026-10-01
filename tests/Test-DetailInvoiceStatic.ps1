[CmdletBinding()]
param([string]$RepositoryRoot = '')

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
$writeExcel = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGhiExcel.bas')
$build = Get-Content -Raw (Join-Path $RepositoryRoot 'build\Build-Excel.ps1')
$runtime = Get-Content -Raw (Join-Path $RepositoryRoot 'build\Test-DetailInvoiceRuntime.ps1')
$relatedRuntime = Get-Content -Raw (Join-Path $RepositoryRoot 'build\Test-RelatedInvoiceRuntime.ps1')
$updateRuntime = Get-Content -Raw (Join-Path $RepositoryRoot 'tests\Test-UpdateRuntime.ps1')
$userFormSmoke = Get-Content -Raw (Join-Path $RepositoryRoot 'build\Invoke-UserFormSmoke.ps1')
$commonBuild = Get-Content -Raw (Join-Path $RepositoryRoot 'build\ExcelBuild.Common.psm1')
$upgradeForm = Get-Content -Raw (Join-Path $RepositoryRoot 'build\Upgrade-FormUI.ps1')

function Has([string]$Text, [string]$Pattern) {
    return [regex]::IsMatch($Text, $Pattern, 'IgnoreCase,Multiline')
}

$checks = @(
    [pscustomobject]@{
        Name = 'Template code precedes invoice series'
        Pass = Has $writeExcel 'arrColName\s*=\s*Array\("khmshdon",\s*"khhdon",\s*"shdon"'
    },
    [pscustomobject]@{
        Name = 'Explicit common-column mapping is used'
        Pass = (Has $writeExcel 'DETAIL_COL_TEMPLATE As Long = 1') -and
            (Has $writeExcel 'DETAIL_COL_SERIES As Long = 2') -and
            (Has $writeExcel 'ws\.Cells\(row_ct, arrCol\(col\)\)')
    },
    [pscustomobject]@{
        Name = 'Seller and buyer tax IDs are written as text'
        Pass = (Has $writeExcel 'Case "nbmst", "nmmst"[\s\S]*?NumberFormat = "@"[\s\S]*?Value2 = CStr')
    },
    [pscustomobject]@{
        Name = 'All common fields fill down through column O'
        Pass = Has $writeExcel 'Range\("A" & frow & ":O" & lrow\)[\s\S]*?FillDown'
    },
    [pscustomobject]@{
        Name = 'Build migrates both detail sheets'
        Pass = (Has $build "ChiTietHD_Mua', 'ChiTietHD_Ban") -and
            (Has $build '\$firstColumn\.Insert\(\)') -and
            (Has $build "Columns\.Item\(8\)\.NumberFormat = '@'") -and
            (Has $build "Columns\.Item\(14\)\.NumberFormat = '@'")
    },
    [pscustomobject]@{
        Name = 'Build runs detail runtime regression'
        Pass = Has $build "Test-DetailInvoiceRuntime\.ps1'"
    },
    [pscustomobject]@{
        Name = 'Runtime covers both directions and leading zeros'
        Pass = (Has $runtime 'CodexTestPurchaseDetail') -and
            (Has $runtime 'CodexTestSalesDetail') -and
            (Has $runtime '"0000000001"') -and
            (Has $runtime '"0000000004"')
    },
    [pscustomobject]@{
        Name = 'Build avoids redundant detail column-width COM setter'
        Pass = (Has $build '\[Math\]::Abs\(\$detailTemplateWidth - \$detailSeriesWidth\) -gt 0\.01') -and
            (Has $build '\$detailTemplateColumn\.PasteSpecial\(8\)')
    },
    [pscustomobject]@{
        Name = 'Runtime keeps assertion row separate from ByRef write cursor'
        Pass = (Has $runtime 'Dim writeRow As Long') -and
            (Has $runtime 'writeRow = targetRow') -and
            (Has $runtime 'ghiExcel_ChiTiet sampleJson, writeRow, direction')
    },
    [pscustomobject]@{
        Name = 'Excel runtime scripts resolve workbook paths'
        Pass = (Has $relatedRuntime '\$BuiltWorkbook = \(Resolve-Path -LiteralPath \$BuiltWorkbook\)\.Path') -and
            (Has $updateRuntime '\$BuiltWorkbook = \(Resolve-Path -LiteralPath \$BuiltWorkbook\)\.Path') -and
            (Has $userFormSmoke '\$BuiltWorkbook = \(Resolve-Path -LiteralPath \$BuiltWorkbook\)\.Path') -and
            (Has $commonBuild '\$resolvedPath = \(Resolve-Path -LiteralPath \$Path\)\.Path') -and
            (Has $commonBuild 'Workbooks\.Open\(\$resolvedPath') -and
            (Has $upgradeForm '\$OutputPath = \[IO\.Path\]::GetFullPath\(\$OutputPath\)')
    }
)

$parseErrors = @()
foreach ($script in @(
    (Join-Path $RepositoryRoot 'build\Build-Excel.ps1'),
    (Join-Path $RepositoryRoot 'build\ExcelBuild.Common.psm1'),
    (Join-Path $RepositoryRoot 'build\Test-DetailInvoiceRuntime.ps1'),
    (Join-Path $RepositoryRoot 'build\Test-RelatedInvoiceRuntime.ps1'),
    (Join-Path $RepositoryRoot 'build\Invoke-UserFormSmoke.ps1'),
    (Join-Path $RepositoryRoot 'build\Upgrade-FormUI.ps1'),
    (Join-Path $RepositoryRoot 'tests\Test-UpdateRuntime.ps1'),
    $PSCommandPath
)) {
    $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($script, [ref]$null, [ref]$errors)
    foreach ($error in @($errors)) { $parseErrors += "$(Split-Path -Leaf $script): $($error.Message)" }
}
$checks += [pscustomobject]@{ Name = 'PowerShell scripts parse'; Pass = ($parseErrors.Count -eq 0) }

foreach ($check in $checks) {
    $status = if ($check.Pass) { 'PASS' } else { 'FAIL' }
    Write-Host "[$status] $($check.Name)"
}
if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { Write-Host $_ } }
$failed = @($checks | Where-Object { -not $_.Pass })
if ($failed.Count -gt 0) { throw "$($failed.Count) detail-invoice static check(s) failed." }
Write-Host 'DETAIL-INVOICE STATIC TEST PASSED'
