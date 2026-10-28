[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$reportDir = Join-Path (Split-Path -Parent $BuiltWorkbook) 'regression'
$backupDir = Join-Path $reportDir 'prior-reports'
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$reports = @('built-inventory.json','comparison-report.json','detail-invoice-runtime-result.json','retry-ui-report.md','retry-ui-test-results.json','related-runtime-result.json')
foreach ($name in $reports) {
    $path = Join-Path $root ('tests\' + $name)
    if (Test-Path -LiteralPath $path) { Copy-Item -LiteralPath $path -Destination (Join-Path $backupDir $name) -Force }
}
$checks = @(
    @('tests\Test-DetailInvoiceStatic.ps1', @()),
    @('tests\Test-RetryUiStatic.ps1', @()),
    @('build\Test-Build.ps1', @('-BuiltWorkbook',$BuiltWorkbook,'-RunUserFormInstantiation')),
    @('build\Test-PublicWorkbook.ps1', @('-WorkbookPath',$BuiltWorkbook)),
    @('build\Test-DetailInvoiceRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook)),
    @('build\Test-CompatibilityFixesRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook)),
    @('build\Test-CoreFixesRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook)),
    @('build\Test-ReviewSourceRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook)),
    @('build\Test-ModernUiRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook)),
    @('build\Test-ParallelJsonRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook)),
    @('build\Test-DatePdfRuntime.ps1', @('-BuiltWorkbook',$BuiltWorkbook))
)
$powerShellExe = (Get-Process -Id $PID).Path
try {
    foreach ($check in $checks) {
        Write-Output ('Running ' + $check[0])
        $log = Join-Path $reportDir (([IO.Path]::GetFileNameWithoutExtension($check[0])) + '.txt')
        $script = Join-Path $root $check[0]
        $checkArguments = $check[1]
        & $powerShellExe -Mta -NoProfile -ExecutionPolicy Bypass -File $script @checkArguments *> $log
        if ($LASTEXITCODE -ne 0) { Get-Content -LiteralPath $log; throw ('FAILED: ' + $check[0]) }
        Write-Output 'PASS'
    }
    Write-Output 'DATE/PDF REGRESSION CHECKS PASSED'
} finally {
    foreach ($name in $reports) {
        $path = Join-Path $root ('tests\' + $name)
        if (Test-Path -LiteralPath $path) { Copy-Item -LiteralPath $path -Destination (Join-Path $reportDir $name) -Force }
        $backup = Join-Path $backupDir $name
        if (Test-Path -LiteralPath $backup) { Copy-Item -LiteralPath $backup -Destination $path -Force }
    }
}
