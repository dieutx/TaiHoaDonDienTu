[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version,
    [string]$RepositoryRoot,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
$RepositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $RepositoryRoot "dist\TaiHoaDonDienTu_v$Version.xlsm" }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$powerShellExe = (Get-Process -Id $PID).Path

function Invoke-ReleaseCheck([string]$RelativePath, [string[]]$CheckArguments = @()) {
    $script = Join-Path $RepositoryRoot $RelativePath
    if (-not (Test-Path -LiteralPath $script)) { throw "Missing release check: $RelativePath" }
    Write-Output "Running $RelativePath"
    # Excel COM cleanup can deadlock the Windows PowerShell STA at process exit.
    & $powerShellExe -Mta -NoProfile -ExecutionPolicy Bypass -File $script @CheckArguments
    if ($LASTEXITCODE -ne 0) { throw "Release blocked: $RelativePath returned exit code $LASTEXITCODE." }
}

Invoke-ReleaseCheck 'tests\Test-DetailInvoiceStatic.ps1' @('-RepositoryRoot', $RepositoryRoot)
Invoke-ReleaseCheck 'tests\Test-RetryUiStatic.ps1' @('-RepositoryRoot', $RepositoryRoot)
Invoke-ReleaseCheck 'tests\Test-ReleaseGate.ps1'
Invoke-ReleaseCheck 'build\Build-Excel.ps1' @('-Version', $Version, '-RepositoryRoot', $RepositoryRoot, '-OutputPath', $OutputPath)
Invoke-ReleaseCheck 'build\Test-Build.ps1' @('-BuiltWorkbook', $OutputPath, '-RepositoryRoot', $RepositoryRoot, '-RunUserFormInstantiation')
Invoke-ReleaseCheck 'build\Test-RelatedInvoiceRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-InvoiceDateRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-ReviewFixesRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-ExtendedReviewRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-CompatibilityFixesRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-LoginFlowRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-CoreFixesRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-DownloaderParityRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-ReviewSourceRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-FinalRetryAfterRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-ProgressClockRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-ModernUiRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'build\Test-ParallelJsonRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Invoke-ReleaseCheck 'tests\Test-UpdateRuntime.ps1' @('-BuiltWorkbook', $OutputPath)
Write-Output "RELEASE CHECKS PASSED: $OutputPath"
