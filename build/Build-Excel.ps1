[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^v?\d+(\.\d+){1,3}([-.][0-9A-Za-z.-]+)?$')][string]$Version,
    [string]$RepositoryRoot,
    [string]$OutputPath,
    [switch]$SkipTest
)

$ErrorActionPreference = 'Stop'
# $PSScriptRoot is not reliable while parameter defaults are evaluated, so the
# repository root is resolved here in the body instead.
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force
$normalizedVersion = $Version.TrimStart('v')
$errorReportSheetName = 'BaoCao_LoiTaiHD'
$templatePath = Join-Path $RepositoryRoot 'template\App_Template.xlsm'
$distDir = Join-Path $RepositoryRoot 'dist'
$outputPath = if ($OutputPath) { [IO.Path]::GetFullPath($OutputPath) } else { Join-Path $distDir "App_v$normalizedVersion.xlsm" }
if (-not (Test-Path -LiteralPath $templatePath)) { throw "Template not found: $templatePath" }
$outputDirectory = [IO.Path]::GetDirectoryName($outputPath)
if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
Copy-Item -LiteralPath $templatePath -Destination $outputPath -Force
# The template came from a download and carries the Zone.Identifier stream, so
# every copy would open in Protected View with macros disabled.  The generated
# workbook is a local build artifact, so the inherited mark is cleared here.
try { Unblock-File -LiteralPath $outputPath -ErrorAction Stop } catch { Write-Host ('Could not clear the inherited download mark: ' + $_.Exception.Message) }

# VBIDE imports text files through the active Windows ANSI code page, which is
# also what [Text.Encoding]::Default reports.  Git source is UTF-8, so it is
# staged as ANSI before import; files that are already ANSI (older exports) are
# decoded with that code page, which round-trips their bytes.  The ANSI encoder
# is strict on purpose: a character the code page cannot represent fails the
# build instead of silently becoming '?'.  Legacy files that already lost
# characters that way are out of recovery range.
#
# Line endings are normalized to CRLF as well.  A class module imported from an
# LF-only file leaves the saved workbook in a state where VBA cannot compile, so
# every macro fails with 'Cannot run the macro ... all macros may be disabled'.
$stageDir = Join-Path $env:TEMP ('hddt-vba-import-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
$strictUtf8 = New-Object Text.UTF8Encoding($false, $true)
$ansi = [Text.Encoding]::Default
$ansiCodePage = $ansi.CodePage
$strictAnsi = [Text.Encoding]::GetEncoding($ansiCodePage, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
function ConvertTo-CrlfText([AllowEmptyString()][string]$Text) {
    return (($Text -replace "`r`n", "`n") -replace "`n", "`r`n")
}
function Get-StagedVbaFile([IO.FileInfo]$File) {
    $destination = Join-Path $stageDir $File.Name
    $bytes = [IO.File]::ReadAllBytes($File.FullName)
    try {
        $text = ConvertTo-CrlfText ($strictUtf8.GetString($bytes))
        [IO.File]::WriteAllText($destination, $text, $strictAnsi)
    } catch [Text.EncoderFallbackException] {
        throw "$($File.Name) contains characters that code page $ansiCodePage cannot represent. Keep VBA source ASCII and use UniConvert(...) or ChrW(...) for UI text."
    } catch [Text.DecoderFallbackException] {
        $text = ConvertTo-CrlfText ($ansi.GetString($bytes))
        [IO.File]::WriteAllText($destination, $text, $ansi)
    }
    if ($File.Extension -ieq '.frm') {
        $frx = [IO.Path]::ChangeExtension($File.FullName, '.frx')
        if (Test-Path -LiteralPath $frx) { Copy-Item -LiteralPath $frx -Destination ([IO.Path]::ChangeExtension($destination, '.frx')) -Force }
    }
    return $destination
}

$excel = $null
$workbook = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($outputPath, 0, $false)
    Write-Host 'Opened output workbook; replacing VBA components...'

    $remove = @()
    foreach ($component in @($workbook.VBProject.VBComponents)) {
        if ([int]$component.Type -in @(1, 2, 3)) { $remove += $component } else { Release-ComObject $component }
    }
    foreach ($component in $remove) {
        $workbook.VBProject.VBComponents.Remove($component)
        Release-ComObject $component
    }
    Write-Host ('Removed ' + $remove.Count + ' components; importing source...')

    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'src\modules') -Filter '*.bas' | Sort-Object Name)) {
        $imported = $workbook.VBProject.VBComponents.Import((Get-StagedVbaFile $file)); Release-ComObject $imported
    }
    $versionCode = $workbook.VBProject.VBComponents.Item('modVersion').CodeModule
    $versionText = $versionCode.Lines(1, $versionCode.CountOfLines)
    if ($versionText -notmatch ('(?m)^Public Const CURRENT_VERSION As String = "' + [regex]::Escape($normalizedVersion) + '"\r?$')) {
        throw "Imported modVersion does not match requested build version $normalizedVersion. Actual: $($versionText -replace '\r?\n', ' | ')"
    }
    Release-ComObject $versionCode
    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'src\classes') -Filter '*.cls' | Sort-Object Name)) {
        $imported = $workbook.VBProject.VBComponents.Import((Get-StagedVbaFile $file)); Release-ComObject $imported
    }
    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'src\forms') -Filter '*.frm' | Sort-Object Name)) {
        $frx = [IO.Path]::ChangeExtension($file.FullName, '.frx')
        if (-not (Test-Path -LiteralPath $frx)) { throw "Missing FRX for $($file.Name)" }
        $imported = $workbook.VBProject.VBComponents.Import((Get-StagedVbaFile $file)); Release-ComObject $imported
    }

    # Create the update form in the VBA designer so its controls and event
    # handlers are part of the built workbook without hand-editing an FRX blob.
    $updateForm = $workbook.VBProject.VBComponents.Add(3) # vbext_ct_MSForm
    $updateForm.Name = 'frmUpdate'
    $updateForm.Properties.Item('Caption').Value = 'TaiHoaDonDienTu'
    $updateForm.Properties.Item('Width').Value = 410
    $updateForm.Properties.Item('Height').Value = 375
    $updateForm.Properties.Item('StartUpPosition').Value = 1
    $updateControls = @(
        @('Label','lblHeader',20,18,350,24),
        @('Label','lblCurrentTitle',20,54,150,18), @('Label','lblCurrent',180,54,190,18),
        @('Label','lblNewTitle',20,81,150,18), @('Label','lblNew',180,81,190,18),
        @('Label','lblDateTitle',20,108,150,18), @('Label','lblDate',180,108,190,18),
        @('Label','lblNotesTitle',20,141,350,18), @('Label','lblNotes',20,166,350,90),
        @('Label','lblQuestion',20,268,350,20),
        @('CommandButton','cmdDownload',20,300,160,30),
        @('CommandButton','cmdContinue',190,300,190,30)
    )
    foreach ($spec in $updateControls) {
        $control = $updateForm.Designer.Controls.Add(('Forms.' + $spec[0] + '.1'), $spec[1], $true)
        $control.Left = $spec[2]; $control.Top = $spec[3]
        $control.Width = $spec[4]; $control.Height = $spec[5]
        if ($spec[1] -eq 'lblNotes') { $control.WordWrap = $true }
        Release-ComObject $control
    }
    $updateForm.CodeModule.AddFromFile((Join-Path $RepositoryRoot 'src\forms\frmUpdate.code.txt'))
    Release-ComObject $updateForm

    $workbookComponent = $workbook.VBProject.VBComponents.Item('ThisWorkbook')
    $workbookCode = $workbookComponent.CodeModule
    $allWorkbookCode = if ($workbookCode.CountOfLines -gt 0) { $workbookCode.Lines(1, $workbookCode.CountOfLines) } else { '' }
    if ($allWorkbookCode -match '(?im)^\s*(?:Private\s+)?Sub\s+Workbook_Open\s*\(\s*\)') {
        if ($allWorkbookCode -notmatch '(?im)^\s*CheckForUpdate\s*$') {
            for ($line = 1; $line -le $workbookCode.CountOfLines; $line++) {
                if ($workbookCode.Lines($line, 1) -match '^\s*(?:Private\s+)?Sub\s+Workbook_Open\s*\(\s*\)') {
                    $workbookCode.InsertLines($line + 1, '    CheckForUpdate')
                    break
                }
            }
        }
    } else {
        $workbookCode.AddFromString("Private Sub Workbook_Open()`r`n    CheckForUpdate`r`nEnd Sub")
    }
    Release-ComObject $workbookCode; Release-ComObject $workbookComponent

    # Build-time sheet setup.  Macros stay disabled for the whole build so
    # Workbook_Open never runs, therefore the sheet is created through COM and
    # the heading text is decoded from the reporting module instead of being
    # duplicated here as a non-ASCII literal (Windows PowerShell 5.1 reads a
    # BOM-less script with the active ANSI code page).
    Write-Host 'Source imported; creating the error report sheet...'
    $reportHeaders = @(Get-GdtErrorReportHeaders -RepositoryRoot $RepositoryRoot)
    if ($reportHeaders.Count -ne 17) { throw "Expected 17 error report headings in modGdtErrorReport, found $($reportHeaders.Count)." }
    $originalSheet = $workbook.ActiveSheet
    $reportSheet = $null
    try { $reportSheet = $workbook.Worksheets.Item($errorReportSheetName) } catch {
        $reportSheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
        $reportSheet.Name = $errorReportSheetName
    }
    for ($column = 1; $column -le $reportHeaders.Count; $column++) { $reportSheet.Cells.Item(1, $column).Value2 = $reportHeaders[$column - 1] }
    $headerRange = $reportSheet.Range('A1:Q1')
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 15917529
    $headerRange.WrapText = $true
    $headerRange.AutoFilter() | Out-Null
    $columnWidths = @(7, 19, 14, 12, 20, 18, 18, 16, 16, 18, 45, 12, 45, 14, 13, 24, 35)
    for ($column = 1; $column -le $columnWidths.Count; $column++) { $reportSheet.Columns.Item($column).ColumnWidth = $columnWidths[$column - 1] }
    $reportSheet.Columns.Item(2).NumberFormat = 'dd/mm/yyyy hh:mm:ss'
    $reportSheet.Columns.Item(9).NumberFormat = 'dd/mm/yyyy'
    $reportSheet.Activate()
    $excel.ActiveWindow.FreezePanes = $false
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.FreezePanes = $true
    $originalSheet.Activate()
    Release-ComObject $headerRange; Release-ComObject $reportSheet; Release-ComObject $originalSheet
    Write-Host 'Error report sheet ready; saving workbook...'

    $relatedHeaders = Get-GdtRelatedInvoiceHeaders -RepositoryRoot $RepositoryRoot
    if ($relatedHeaders.Count -ne 8) { throw "Expected 8 related-invoice headings in modGhiExcel, found $($relatedHeaders.Count)." }
    foreach ($summarySheetName in @('TongHopHD_Mua', 'TongHopHD_Ban')) {
        $summarySheet = $workbook.Worksheets.Item($summarySheetName)
        $formatSource = $summarySheet.Range('BA2')
        $relatedHeaderRange = $summarySheet.Range('BE2:BL2')
        $formatSource.Copy() | Out-Null
        $relatedHeaderRange.PasteSpecial(-4122) | Out-Null # xlPasteFormats
        foreach ($column in $relatedHeaders.Keys) { $summarySheet.Cells.Item(2, [int]$column).Value2 = $relatedHeaders[$column] }
        $summarySheet.Columns.Item(65).Clear() | Out-Null
        $summarySheet.Columns.Item(57).ColumnWidth = 32
        for ($column = 58; $column -le 63; $column++) { $summarySheet.Columns.Item($column).ColumnWidth = 18 }
        $summarySheet.Columns.Item(64).ColumnWidth = 45
        $summarySheet.Columns.Item(62).NumberFormat = 'dd/mm/yyyy'
        Release-ComObject $relatedHeaderRange; Release-ComObject $formatSource; Release-ComObject $summarySheet
    }
    $excel.CutCopyMode = $false
    Write-Host 'Related-invoice columns ready on purchase and sales summaries.'

    $workbook.SaveAs($outputPath, 52)
    Write-Host 'Workbook saved; closing build Excel instance...'
    $workbook.Close($true); Release-ComObject $workbook; $workbook = $null
} catch {
    $failure = $_
    Write-Host ('BUILD FAILED: ' + $failure.Exception.Message)
    if (Test-Path -LiteralPath $outputPath) { try { Remove-Item -LiteralPath $outputPath -Force } catch {} }
    throw $failure
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $workbook; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $stageDir) { Remove-Item -LiteralPath $stageDir -Recurse -Force }
}
Write-Host 'Build Excel instance closed.'

if (-not $SkipTest) {
    $powerShellExe = (Get-Process -Id $PID).Path
    & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Build.ps1') -BuiltWorkbook $outputPath -RepositoryRoot $RepositoryRoot
    if ($LASTEXITCODE -ne 0) { throw "BUILD FAILED: structural test process returned exit code $LASTEXITCODE." }
    & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-PublicWorkbook.ps1') -WorkbookPath $outputPath
    if ($LASTEXITCODE -ne 0) { throw "BUILD FAILED: public workbook test process returned exit code $LASTEXITCODE." }
}
Write-Host "BUILD GENERATED; STRUCTURAL TESTS PASSED: $outputPath"
