[CmdletBinding()]
param(
    [string]$BuiltWorkbook,
    [string]$RepositoryRoot,
    [switch]$RunUserFormInstantiation
)

$ErrorActionPreference = 'Stop'
# $PSScriptRoot is not reliable while parameter defaults are evaluated, so the
# repository root is resolved here in the body instead.
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force
if ([string]::IsNullOrWhiteSpace($BuiltWorkbook)) {
    $BuiltWorkbook = (Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'dist') -Filter 'App_v*.xlsm' | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}
if (-not (Test-Path -LiteralPath $BuiltWorkbook)) { throw "Built workbook not found: $BuiltWorkbook" }
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$originalPath = Join-Path $RepositoryRoot 'tests\original-inventory.json'
if (-not (Test-Path -LiteralPath $originalPath)) { throw "Original inventory not found: $originalPath" }
$original = Get-Content -LiteralPath $originalPath -Raw | ConvertFrom-Json
$testDir = Join-Path $RepositoryRoot 'tests'
$builtInventoryPath = Join-Path $testDir 'built-inventory.json'
$reportPath = Join-Path $testDir 'comparison-report.json'

function Read-GdtErrorReportSheet([object]$Workbook, [string]$SheetName, [int]$ColumnCount) {
    $snapshot = [ordered]@{ Exists = $false; Headers = @(); AutoFilterMode = $false; FreezePanes = $null }
    $sheet = $null
    try { $sheet = $Workbook.Worksheets.Item($SheetName) } catch { return $snapshot }
    try {
        $snapshot.Exists = $true
        $headers = @()
        for ($column = 1; $column -le $ColumnCount; $column++) { $headers += [string]$sheet.Cells.Item(1, $column).Value2 }
        $snapshot.Headers = $headers
        $snapshot.AutoFilterMode = [bool]$sheet.AutoFilterMode
        try { $sheet.Activate(); $snapshot.FreezePanes = [bool]$Workbook.Windows.Item(1).FreezePanes } catch { $snapshot.FreezePanes = $null }
    } finally {
        Release-ComObject $sheet
    }
    return $snapshot
}

$errorReportSheetSnapshot = [ordered]@{ Exists = $false; Headers = @(); AutoFilterMode = $false; FreezePanes = $null }
$updateIntegrationSnapshot = [ordered]@{ WorkbookOpenCode = ''; FormCode = '' }
Write-Host 'Opening built workbook for inventory...'
Invoke-WithExcelWorkbook -Path $BuiltWorkbook -ReadOnly -Action {
    param($workbook, $excel)
    $inventory = Get-WorkbookInventory -Workbook $workbook -Excel $excel -Path $BuiltWorkbook
    $inventory | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $builtInventoryPath -Encoding utf8
    $captured = Read-GdtErrorReportSheet -Workbook $workbook -SheetName 'BaoCao_LoiTaiHD' -ColumnCount 17
    foreach ($key in @($captured.Keys)) { $errorReportSheetSnapshot[$key] = $captured[$key] }
    foreach ($componentName in @('ThisWorkbook', 'frmUpdate')) {
        $component = $workbook.VBProject.VBComponents.Item($componentName)
        $codeModule = $component.CodeModule
        $codeText = if ($codeModule.CountOfLines -gt 0) { $codeModule.Lines(1, $codeModule.CountOfLines) } else { '' }
        if ($componentName -eq 'ThisWorkbook') { $updateIntegrationSnapshot.WorkbookOpenCode = $codeText }
        else { $updateIntegrationSnapshot.FormCode = $codeText }
        Release-ComObject $codeModule; Release-ComObject $component
    }
}
Write-Host 'Built inventory captured; comparing metadata...'
$built = Get-Content -LiteralPath $builtInventoryPath -Raw | ConvertFrom-Json

function Compare-Set($label, $left, $right) {
    $differences = @(Compare-Object -ReferenceObject @($left | Sort-Object) -DifferenceObject @($right | Sort-Object))
    [ordered]@{ Name = $label; Pass = ($differences.Count -eq 0); Differences = $differences }
}

function Normalize-OnAction([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    return ($value -replace "^'?(?:[^']+\.xlsm)'?!", '')
}

# The retry/error-report feature adds one dedicated worksheet at build time.
# It is a deliberate addition, so it is reported as an expected structural
# addition instead of failing the comparison that protects original sheets.
$expectedAddedWorksheetNames = @('BaoCao_LoiTaiHD')
$expectedAddedComponentNames = @('modGdtRetry', 'modGdtErrorReport', 'clsGdtRequestResult', 'clsGdtRetryItem', 'Sheet10', 'modVersion', 'modUpdate', 'frmUpdate')
$expectedAddedNamePatterns = @('^BaoCao_LoiTaiHD!')
$originalWorksheetNames = @($original.Worksheets | ForEach-Object { $_.Name })
$expectedAddedWorksheets = @($built.Worksheets | Where-Object { $_.Name -in $expectedAddedWorksheetNames -and $_.Name -notin $originalWorksheetNames } | ForEach-Object { $_.Name })
$unexpectedNewWorksheets = @($built.Worksheets | Where-Object { $_.Name -notin $originalWorksheetNames -and $_.Name -notin $expectedAddedWorksheetNames } | ForEach-Object { $_.Name })
$comparableBuiltWorksheets = @($built.Worksheets | Where-Object { $_.Name -in $originalWorksheetNames })

$checks = @()
$checks += Compare-Set 'Worksheets' @($original.Worksheets | ForEach-Object { "$($_.Name)|$($_.CodeName)|$($_.Visibility)" }) @($comparableBuiltWorksheets | ForEach-Object { "$($_.Name)|$($_.CodeName)|$($_.Visibility)" })
$checks += [ordered]@{ Name = 'UnexpectedNewWorksheets'; Pass = ($unexpectedNewWorksheets.Count -eq 0); Differences = @($unexpectedNewWorksheets) }
$checks += [ordered]@{ Name = 'ExpectedAddedWorksheets'; Pass = ($expectedAddedWorksheets.Count -eq $expectedAddedWorksheetNames.Count); Differences = @($expectedAddedWorksheetNames | Where-Object { $_ -notin $expectedAddedWorksheets }) }
$comparableBuiltNames = @($built.NamedRanges | Where-Object { $name = $_.Name; -not (@($expectedAddedNamePatterns | Where-Object { $name -match $_ }).Count -gt 0) })
$checks += Compare-Set 'NamedRanges' @($original.NamedRanges | ForEach-Object { "$($_.Name)|$($_.RefersTo)|$($_.Visible)" }) @($comparableBuiltNames | ForEach-Object { "$($_.Name)|$($_.RefersTo)|$($_.Visible)" })
$checks += [ordered]@{
    Name = 'ExpectedAddedNamedRanges'
    Pass = (@($built.NamedRanges | Where-Object { $name = $_.Name; @($expectedAddedNamePatterns | Where-Object { $name -match $_ }).Count -gt 0 }).Count -ge 1)
    Differences = @()
    Explanation = 'AutoFilter on the added report sheet creates its hidden filter database name.'
}
$originalNameKeys = @($original.NamedRanges | ForEach-Object { $_.Name })
$unexpectedAddedNames = @($built.NamedRanges | Where-Object { $name = $_.Name; ($name -notin $originalNameKeys) -and -not (@($expectedAddedNamePatterns | Where-Object { $name -match $_ }).Count -gt 0) } | ForEach-Object { $_.Name })
$checks += [ordered]@{
    Name = 'UnexpectedNewNamedRanges'
    Pass = ($unexpectedAddedNames.Count -eq 0)
    Differences = @($unexpectedAddedNames)
}
$checks += Compare-Set 'Tables' @($original.Worksheets | ForEach-Object { $s=$_; $_.Tables | ForEach-Object { "$($s.Name)|$($_.Name)|$($_.Range)" } }) @($comparableBuiltWorksheets | ForEach-Object { $s=$_; $_.Tables | ForEach-Object { "$($s.Name)|$($_.Name)|$($_.Range)" } })
$checks += Compare-Set 'ShapesAndMacroAssignments' @($original.Worksheets | ForEach-Object { $s=$_; $_.Shapes | ForEach-Object { "$($s.Name)|$($_.Name)|$($_.Type)|$(Normalize-OnAction $_.OnAction)|$($_.FormControlType)" } }) @($comparableBuiltWorksheets | ForEach-Object { $s=$_; $_.Shapes | ForEach-Object { "$($s.Name)|$($_.Name)|$($_.Type)|$(Normalize-OnAction $_.OnAction)|$($_.FormControlType)" } })
$checks += Compare-Set 'ActiveXControls' @($original.Worksheets | ForEach-Object { $s=$_; $_.ActiveXControls | ForEach-Object { "$($s.Name)|$($_.Name)|$($_.ProgId)" } }) @($comparableBuiltWorksheets | ForEach-Object { $s=$_; $_.ActiveXControls | ForEach-Object { "$($s.Name)|$($_.Name)|$($_.ProgId)" } })
$checks += Compare-Set 'VbaReferences' @($original.VbaReferences | ForEach-Object { "$($_.Name)|$($_.Guid)|$($_.Major)|$($_.Minor)|$($_.IsBroken)" }) @($built.VbaReferences | ForEach-Object { "$($_.Name)|$($_.Guid)|$($_.Major)|$($_.Minor)|$($_.IsBroken)" })
$originalComponentKeys = @($original.VbaComponents | ForEach-Object { "$($_.Name)|$($_.Type)" })
$comparableBuiltComponents = @($built.VbaComponents | Where-Object { "$($_.Name)|$($_.Type)" -in $originalComponentKeys })
$addedBuiltComponents = @($built.VbaComponents | Where-Object { "$($_.Name)|$($_.Type)" -notin $originalComponentKeys })
$unexpectedBuiltComponents = @($addedBuiltComponents | Where-Object { $_.Name -notin $expectedAddedComponentNames })
$missingExpectedComponents = @($expectedAddedComponentNames | Where-Object { $_ -notin @($addedBuiltComponents | ForEach-Object { $_.Name }) })
$checks += Compare-Set 'VbaComponents' @($original.VbaComponents | ForEach-Object { "$($_.Name)|$($_.Type)" }) @($comparableBuiltComponents | ForEach-Object { "$($_.Name)|$($_.Type)" })
$checks += [ordered]@{
    Name = 'UnexpectedNewVbaComponents'
    Pass = ($unexpectedBuiltComponents.Count -eq 0)
    Differences = @($unexpectedBuiltComponents | ForEach-Object { "$($_.Name)|$($_.Type)" })
}
$checks += [ordered]@{
    Name = 'ExpectedAddedVbaComponents'
    Pass = ($missingExpectedComponents.Count -eq 0)
    Differences = @($missingExpectedComponents)
    Explanation = 'Components added by the retry/error-report feature, plus the document module of the added worksheet.'
}
$originalFormNames = @($original.VbaComponents | Where-Object Type -eq 3 | ForEach-Object { $_.Name })
$checks += Compare-Set 'UserFormControls' @($original.VbaComponents | Where-Object Type -eq 3 | ForEach-Object { $f=$_; $_.Controls | ForEach-Object { "$($f.Name)|$($_.Name)|$($_.ProgId)" } }) @($built.VbaComponents | Where-Object { $_.Type -eq 3 -and $_.Name -in $originalFormNames } | ForEach-Object { $f=$_; $_.Controls | ForEach-Object { "$($f.Name)|$($_.Name)|$($_.ProgId)" } })
$requiredUpdateControls = @('lblHeader','lblCurrentTitle','lblCurrent','lblNewTitle','lblNew','lblDateTitle','lblDate','lblNotesTitle','lblNotes','lblQuestion','cmdDownload','cmdContinue')
$actualUpdateControls = @($built.VbaComponents | Where-Object Name -eq 'frmUpdate' | ForEach-Object { $_.Controls | ForEach-Object { $_.Name } })
$checks += [ordered]@{ Name = 'UpdateFormControls'; Pass = (@($requiredUpdateControls | Where-Object { $_ -notin $actualUpdateControls }).Count -eq 0); Differences = @($requiredUpdateControls | Where-Object { $_ -notin $actualUpdateControls }) }
$checks += [ordered]@{ Name = 'WorkbookOpenChecksForUpdate'; Pass = ($updateIntegrationSnapshot.WorkbookOpenCode -match '(?is)Sub\s+Workbook_Open\s*\(\s*\).*?CheckForUpdate'); Differences = @() }
$checks += [ordered]@{ Name = 'UpdateFormHandlers'; Pass = ($updateIntegrationSnapshot.FormCode -match 'Private Sub cmdDownload_Click' -and $updateIntegrationSnapshot.FormCode -match 'Private Sub cmdContinue_Click'); Differences = @() }

# The 17 error-report headings come from the reporting module at build time, so
# the Telex source literals are decoded here and compared with the workbook.
function Get-VbaLogicalLines([string]$Text) {
    $logicalLines = @()
    $buffer = ''
    foreach ($line in @($Text -split "`r?`n")) {
        $trimmed = $line.Trim()
        if ($trimmed.EndsWith('_')) { $buffer += $trimmed.Substring(0, $trimmed.Length - 1) + ' '; continue }
        $logicalLines += ($buffer + $trimmed)
        $buffer = ''
    }
    return $logicalLines
}

function Get-UniConvertTables([string]$MsgBoxSource) {
    $telex = @()
    $codes = @()
    foreach ($line in @(Get-VbaLogicalLines $MsgBoxSource)) {
        if ($line -match '^Telex_Type = Array\((.*)\)$') { $telex = @([regex]::Matches($Matches[1], '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }) }
        elseif ($line -match '^CharCode = Array\((.*)\)$') { $codes = @([regex]::Matches($Matches[1], 'ChrW\((\d+)\)') | ForEach-Object { [int]$_.Groups[1].Value }) }
    }
    return New-Object PSObject -Property @{ Telex = $telex; Codes = $codes }
}

function Convert-TelexToUnicode([string]$Text, [object]$Tables) {
    $result = $Text
    for ($index = 0; $index -lt $Tables.Codes.Count -and $index -lt $Tables.Telex.Count; $index++) {
        $result = $result.Replace($Tables.Telex[$index], [string][char]$Tables.Codes[$index])
    }
    return $result
}

function Get-DeclaredErrorReportHeaders([string]$ReportSource, [object]$Tables) {
    $headers = @()
    foreach ($line in @($ReportSource -split "`r?`n")) {
        if ($line -notmatch '^\s*headers\(1,\s*(\d+)\)\s*=\s*(.+?)\s*$') { continue }
        $column = [int]$Matches[1]
        $expression = $Matches[2]
        if ($expression -match '^UniConvert\("([^"]*)"\)$') { $headers += (Convert-TelexToUnicode $Matches[1] $Tables) }
        elseif ($expression -match '^"([^"]*)"$') { $headers += $Matches[1] }
        else { $headers += "<unsupported expression for column $column>" }
    }
    return $headers
}

$reportSource = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGdtErrorReport.bas')
$msgBoxSource = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modMsgboxTV.bas')
$uniTables = Get-UniConvertTables $msgBoxSource
$declaredHeaders = @()
if ($uniTables.Codes.Count -gt 0) { $declaredHeaders = @(Get-DeclaredErrorReportHeaders $reportSource $uniTables) }
$headerDifferences = @(Compare-Object -ReferenceObject $declaredHeaders -DifferenceObject @($errorReportSheetSnapshot.Headers) -SyncWindow 0)
$checks += [ordered]@{
    Name = 'ErrorReportSheet'
    Pass = ($errorReportSheetSnapshot.Exists -and $uniTables.Codes.Count -gt 0 -and $declaredHeaders.Count -eq 17 -and $errorReportSheetSnapshot.Headers.Count -eq 17 -and $headerDifferences.Count -eq 0 -and $errorReportSheetSnapshot.AutoFilterMode)
    Differences = $headerDifferences
    DeclaredHeaders = $declaredHeaders
    ActualHeaders = @($errorReportSheetSnapshot.Headers)
}
$checks += [ordered]@{
    Name = 'ErrorReportSheetFrozenHeader'
    Pass = ($errorReportSheetSnapshot.FreezePanes -eq $true)
    Differences = @("FreezePanes=$($errorReportSheetSnapshot.FreezePanes)")
}
$unsafeBuildScripts = @()
foreach ($script in @(Get-ChildItem -Path (Join-Path $PSScriptRoot '*.ps1'), (Join-Path $PSScriptRoot '*.psm1'))) {
    $bytes = [IO.File]::ReadAllBytes($script.FullName)
    $hasByteOrderMark = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $hasNonAscii = (@($bytes | Where-Object { $_ -gt 127 }).Count -gt 0)
    if ($hasNonAscii -and -not $hasByteOrderMark) { $unsafeBuildScripts += $script.Name }
}
$checks += [ordered]@{
    Name = 'BuildScriptsEncodingSafe'
    Pass = ($unsafeBuildScripts.Count -eq 0)
    Differences = @($unsafeBuildScripts)
    Explanation = 'Windows PowerShell 5.1 reads a BOM-less script with the active ANSI code page, so a non-ASCII literal corrupts the script at parse time. Such scripts need a UTF-8 BOM.'
}

# Code hashes are compared with the source text rather than with the migration
# baseline: the baseline was captured from a previous build output, and parts of
# it come from ANSI exports that had already replaced unsupported Vietnamese
# characters with '?', so it can never be reproduced from the current source.
$sourceCodeMismatches = @()
foreach ($component in @($built.VbaComponents)) {
    $sourcePath = $null
    switch ([int]$component.Type) {
        1 { $sourcePath = Join-Path $RepositoryRoot "src\modules\$($component.Name).bas" }
        2 { $sourcePath = Join-Path $RepositoryRoot "src\classes\$($component.Name).cls" }
        3 { $sourcePath = Join-Path $RepositoryRoot "src\forms\$($component.Name).frm" }
    }
    if (-not $sourcePath -or -not (Test-Path -LiteralPath $sourcePath)) { continue }
    $expectedHash = Get-Sha256Text ((Get-VbaCodeTextFromSource -Path $sourcePath).ToLowerInvariant())
    if ($expectedHash -ne [string]$component.CodeSha256IgnoringCase) { $sourceCodeMismatches += $component.Name }
}
$baselineCodeDifferences = @(Compare-Object -ReferenceObject @($original.VbaComponents | ForEach-Object { "$($_.Name)|$($_.CodeSha256)" }) -DifferenceObject @($comparableBuiltComponents | ForEach-Object { "$($_.Name)|$($_.CodeSha256)" }))
$checks += [ordered]@{
    Name = 'VbaCodeMatchesSource'
    Pass = ($sourceCodeMismatches.Count -eq 0)
    Differences = @($sourceCodeMismatches)
    Explanation = 'Every reimported module, class and form contains the source code, compared without case because VBIDE rewrites member casing on compile, and without module attributes because VBIDE hides them.'
}

$brokenReferences = @($built.VbaReferences | Where-Object IsBroken)
$expectedSerializationDifferences = @()
foreach ($originalSheet in @($original.Worksheets)) {        $builtSheet = $comparableBuiltWorksheets | Where-Object Name -eq $originalSheet.Name | Select-Object -First 1
    foreach ($originalShape in @($originalSheet.Shapes | Where-Object { -not [string]::IsNullOrWhiteSpace($_.OnAction) })) {
        $builtShape = $builtSheet.Shapes | Where-Object Name -eq $originalShape.Name | Select-Object -First 1
        if ($null -ne $builtShape -and $originalShape.OnAction -ne $builtShape.OnAction -and (Normalize-OnAction $originalShape.OnAction) -eq (Normalize-OnAction $builtShape.OnAction)) {
            $expectedSerializationDifferences += [ordered]@{
                Category = 'Expected serialization difference'
                Item = "$($originalSheet.Name)/$($originalShape.Name)"
                Original = [string]$originalShape.OnAction
                Built = [string]$builtShape.OnAction
                Explanation = 'Excel rebound the workbook-qualified OnAction prefix to the new output filename; the macro target is unchanged.'
            }
        }
    }
}
$formSmoke = [ordered]@{ Attempted = $false; Pass = $null; Results = @(); Note = 'NOT VERIFIED by the default test. Use -RunUserFormInstantiation in a suitable interactive Excel session.' }
if ($RunUserFormInstantiation) {
    $formSmoke.Attempted = $true
    $formSmoke.Note = 'Each UserForm was instantiated in an isolated Excel/PowerShell process using a temporary VBA test module and then unloaded.'
    $powerShellExe = (Get-Process -Id $PID).Path
    $helper = Join-Path $PSScriptRoot 'Invoke-UserFormSmoke.ps1'
    foreach ($form in @($built.VbaComponents | Where-Object Type -eq 3)) {
        & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $helper -BuiltWorkbook $BuiltWorkbook -FormName ([string]$form.Name) *> $null
        $ok = ($LASTEXITCODE -eq 0)
        $formSmoke.Results += [ordered]@{ Name = [string]$form.Name; Pass = $ok; ExitCode = $LASTEXITCODE }
    }
    $formSmoke.Pass = (@($formSmoke.Results | Where-Object { -not $_.Pass }).Count -eq 0)
}

$structuralPass = (@($checks | Where-Object { -not $_.Pass }).Count -eq 0) -and ($brokenReferences.Count -eq 0)
$pass = $structuralPass -and ($formSmoke.Pass -ne $false)
$status = if (-not $pass) { 'FAIL' } elseif (-not $formSmoke.Attempted) { 'PARTIAL' } else { 'PASS' }
$report = [ordered]@{
    Status = $status
    BuiltWorkbook = $BuiltWorkbook
    ReopenVerified = $true
    Checks = $checks
    BrokenReferences = $brokenReferences
    UserFormSmokeTest = $formSmoke
    ExpectedSerializationDifferences = $expectedSerializationDifferences
    CriticalBusinessMacroExecution = [ordered]@{ Status = 'NOT VERIFIED'; Reason = 'Business macros perform external/network and workbook mutations; no safe representative inputs were supplied.' }
    VbaCompile = [ordered]@{
        Status = 'NOT VERIFIED (no supported compile API)'
        Reason = 'Excel exposes no supported programmatic compile API, so compilation is only evidenced by running code: the workbook is reopened and every UserForm is instantiated.'
        RuntimeEvidence = $(if ($formSmoke.Attempted) {
            if ($formSmoke.Pass) { "All $($formSmoke.Results.Count) UserForms instantiated and unloaded in isolated processes." } else { 'At least one UserForm failed to instantiate.' }
        } else { 'Not attempted in this run; use -RunUserFormInstantiation.' })
    }
    ExpectedStructuralAdditions = [ordered]@{
        Worksheets = @($expectedAddedWorksheets)
        VbaComponents = @($expectedAddedComponentNames | Where-Object { $_ -in @($addedBuiltComponents | ForEach-Object { $_.Name }) })
        NamedRanges = @($built.NamedRanges | Where-Object { $name = $_.Name; @($expectedAddedNamePatterns | Where-Object { $name -match $_ }).Count -gt 0 } | ForEach-Object { $_.Name })
        Explanation = 'Feature-added worksheet. Sheet names, CodeNames, visibility, tables, shapes and controls of the original worksheets are still compared unchanged.'
    }
    CodeDifferencesVsMigrationBaseline = $baselineCodeDifferences
    ErrorReportSheet = $errorReportSheetSnapshot
}
$report | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $reportPath -Encoding utf8
if (-not $pass) { throw "BUILD FAILED. See $reportPath" }
Write-Host "STRUCTURAL TESTS PASSED ($status). See $reportPath"
