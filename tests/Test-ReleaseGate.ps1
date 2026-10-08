[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$fixtureRoot = Join-Path $env:TEMP ('hddt-release-gate-' + [guid]::NewGuid().ToString('N'))
$fixtureTests = Join-Path $fixtureRoot 'tests'
New-Item -ItemType Directory -Path $fixtureTests -Force | Out-Null
$marker = Join-Path $fixtureRoot 'unexpected-next-check'
$firstScript = Join-Path $fixtureTests 'Test-DetailInvoiceStatic.ps1'
$secondScript = Join-Path $fixtureTests 'Test-RetryUiStatic.ps1'
$log = Join-Path $fixtureRoot 'gate.log'
$runner = Join-Path (Split-Path -Parent $PSScriptRoot) 'build\Test-Release.ps1'
$powerShellExe = (Get-Process -Id $PID).Path
try {
    [IO.File]::WriteAllText($firstScript, "param([string]`$RepositoryRoot)`r`nexit 17`r`n")
    [IO.File]::WriteAllText($secondScript, "param([string]`$RepositoryRoot)`r`n[IO.File]::WriteAllText('$marker', 'unexpected')`r`nexit 0`r`n")
    $previousErrorPreference = $ErrorActionPreference
    try {
        # A failing child is expected here; PowerShell 5.1 represents its stderr
        # as NativeCommandError even when redirected to the fixture log.
        $ErrorActionPreference = 'Continue'
        & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $runner -Version '6.7.6' -RepositoryRoot $fixtureRoot *> $log
        $gateExitCode = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previousErrorPreference }
    if ($gateExitCode -eq 0 -or (Test-Path -LiteralPath $marker)) { throw 'A failed check did not stop the release suite.' }
    if ((Get-Content -Raw -LiteralPath $log) -notmatch 'returned exit code 17') { throw 'Release failure did not name the failing check.' }
    Write-Output 'RELEASE-GATE TEST PASSED: exit 17 blocks all later checks'
} finally {
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -match '^hddt-release-gate-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
