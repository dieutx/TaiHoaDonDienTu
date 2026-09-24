Set-StrictMode -Version Latest

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function New-ExcelApplication {
    # Late binding through the CLSID is used instead of New-Object -ComObject
    # because that cmdlet resolves the COM interop class and casts it to the
    # Excel interface IIDs.  When Excel's interface registration is incomplete
    # (HKCR\Interface\{000208D5-...} missing) the cast fails with 0x80040155
    # (Interface not registered), while this IDispatch path still works.
    $excelType = [Type]::GetTypeFromCLSID([Guid]'00024500-0000-0000-C000-000000000046')
    if ($null -eq $excelType) { throw 'Excel.Application CLSID 00024500-0000-0000-C000-000000000046 is not registered.' }
    return [Activator]::CreateInstance($excelType)
}

function Get-VbaLogicalLines {
    param([AllowEmptyString()][string]$Text)
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

function Get-UniConvertTables {
    # UniConvert maps an ASCII Telex source to Vietnamese.  The tables are read
    # from the module that defines the function so the build can produce the
    # same text as the runtime without any non-ASCII literal in a script.
    param([Parameter(Mandatory)][string]$MsgBoxSource)
    $telex = @()
    $codes = @()
    foreach ($line in @(Get-VbaLogicalLines $MsgBoxSource)) {
        if ($line -match '^Telex_Type = Array\((.*)\)$') { $telex = @([regex]::Matches($Matches[1], '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }) }
        elseif ($line -match '^CharCode = Array\((.*)\)$') { $codes = @([regex]::Matches($Matches[1], 'ChrW\((\d+)\)') | ForEach-Object { [int]$_.Groups[1].Value }) }
    }
    return New-Object PSObject -Property @{ Telex = $telex; Codes = $codes }
}

function ConvertFrom-TelexSource {
    param([AllowEmptyString()][string]$Text, [Parameter(Mandatory)][object]$Tables)
    $result = $Text
    for ($index = 0; $index -lt $Tables.Codes.Count -and $index -lt $Tables.Telex.Count; $index++) {
        $result = $result.Replace($Tables.Telex[$index], [string][char]$Tables.Codes[$index])
    }
    return $result
}

function Get-GdtErrorReportHeaders {
    # Single source of truth for the report headings: the literals inside
    # FormatGdtErrorReport in the reporting module.
    param([Parameter(Mandatory)][string]$RepositoryRoot)
    $tables = Get-UniConvertTables -MsgBoxSource (Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modMsgboxTV.bas'))
    if ($tables.Codes.Count -eq 0) { throw 'Could not read the UniConvert tables from modMsgboxTV.bas.' }
    $headers = @()
    foreach ($line in @((Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGdtErrorReport.bas')) -split "`r?`n")) {
        if ($line -notmatch '^\s*headers\(1,\s*(\d+)\)\s*=\s*(.+?)\s*$') { continue }
        $expression = $Matches[2]
        if ($expression -match '^UniConvert\("([^"]*)"\)$') { $headers += (ConvertFrom-TelexSource $Matches[1] $tables) }
        elseif ($expression -match '^"([^"]*)"$') { $headers += $Matches[1] }
        else { throw "Unsupported heading expression in modGdtErrorReport: $expression" }
    }
    return $headers
}

function Get-GdtRelatedInvoiceHeaders {
    param([Parameter(Mandatory)][string]$RepositoryRoot)
    $tables = Get-UniConvertTables -MsgBoxSource (Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modMsgboxTV.bas'))
    if ($tables.Codes.Count -eq 0) { throw 'Could not read the UniConvert tables from modMsgboxTV.bas.' }
    $source = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGhiExcel.bas')
    $headers = [ordered]@{}
    foreach ($match in [regex]::Matches($source, 'Cells\(2,\s*(\d+)\)\.Value\s*=\s*UniConvert\("([^"]*)"\)')) {
        $column = [int]$match.Groups[1].Value
        if ($column -ge 57 -and $column -le 64) {
            $headers[[string]$column] = ConvertFrom-TelexSource $match.Groups[2].Value $tables
        }
    }
    return $headers
}

function Get-Sha256Text {
    param([AllowEmptyString()][string]$Text)
    $normalized = ($Text -replace "`r`n", "`n").Trim()
    $bytes = [Text.Encoding]::UTF8.GetBytes($normalized)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $hash.Dispose() }
}

function Get-VbaComponentTypeName {
    param([int]$Type)
    switch ($Type) {
        1 { 'StandardModule' }
        2 { 'ClassModule' }
        3 { 'UserForm' }
        100 { 'DocumentModule' }
        default { "Unknown($Type)" }
    }
}

function Get-VbaSourceText {
    # Files written by newer tools are UTF-8; files from the earlier export are
    # in the active ANSI code page.  This mirrors how the build stages them.
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    $strictUtf8 = New-Object Text.UTF8Encoding($false, $true)
    try { return $strictUtf8.GetString($bytes) }
    catch { return [Text.Encoding]::Default.GetString($bytes) }
}

function Get-VbaCodeTextFromSource {
    # VBIDE keeps module and member attributes out of CodeModule, so the source
    # text is reduced to the same shape before it is compared with a workbook.
    param([Parameter(Mandatory)][string]$Path)
    $lines = @((Get-VbaSourceText -Path $Path) -split "`r?`n")
    $codeStart = 0
    while ($codeStart -lt $lines.Count) {
        $line = $lines[$codeStart]
        if ($line.Trim().Length -eq 0) { $codeStart++; continue }
        if ($line -match '^\s') { $codeStart++; continue }
        if ($line -match '^(VERSION|BEGIN|END|Attribute|MultiUse)\b') { $codeStart++; continue }
        break
    }
    $codeLines = @()
    for ($index = $codeStart; $index -lt $lines.Count; $index++) {
        if ($lines[$index] -match '^\s*Attribute\s+[A-Za-z_][A-Za-z0-9_]*\.') { continue }
        $codeLines += $lines[$index]
    }
    return (($codeLines -join "`n").Trim())
}

function Get-WorkbookInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Workbook,
        [Parameter(Mandatory)][object]$Excel,
        [Parameter(Mandatory)][string]$Path
    )

    $worksheets = @()
    foreach ($sheet in @($Workbook.Worksheets)) {
        $tables = @()
        foreach ($table in @($sheet.ListObjects)) {
            $tables += [ordered]@{ Name = [string]$table.Name; Range = [string]$table.Range.Address($false, $false) }
            Release-ComObject $table
        }

        $shapes = @()
        foreach ($shape in @($sheet.Shapes)) {
            $onAction = ''
            try { $onAction = [string]$shape.OnAction } catch {}
            $controlType = $null
            try { $controlType = [int]$shape.FormControlType } catch {}
            $shapes += [ordered]@{
                Name = [string]$shape.Name
                Type = [int]$shape.Type
                OnAction = $onAction
                FormControlType = $controlType
            }
            Release-ComObject $shape
        }

        $oleObjects = @()
        $oleCollection = $sheet.OLEObjects()
        for ($oleIndex = 1; $oleIndex -le [int]$oleCollection.Count; $oleIndex++) {
            $ole = $oleCollection.Item($oleIndex)
            $progId = ''
            try { $progId = [string]$ole.progID } catch {}
            $oleObjects += [ordered]@{ Name = [string]$ole.Name; ProgId = $progId }
            Release-ComObject $ole
        }
        Release-ComObject $oleCollection

        $worksheets += [ordered]@{
            Name = [string]$sheet.Name
            CodeName = [string]$sheet.CodeName
            Visibility = [int]$sheet.Visible
            Tables = $tables
            Shapes = $shapes
            ActiveXControls = $oleObjects
        }
        Release-ComObject $sheet
    }

    $names = @()
    foreach ($name in @($Workbook.Names)) {
        $names += [ordered]@{ Name = [string]$name.Name; RefersTo = [string]$name.RefersTo; Visible = [bool]$name.Visible }
        Release-ComObject $name
    }

    $references = @()
    foreach ($reference in @($Workbook.VBProject.References)) {
        $references += [ordered]@{
            Name = [string]$reference.Name
            Description = [string]$reference.Description
            Guid = [string]$reference.Guid
            Major = [int]$reference.Major
            Minor = [int]$reference.Minor
            FullPath = $(try { [string]$reference.FullPath } catch { '' })
            IsBroken = [bool]$reference.IsBroken
        }
        Release-ComObject $reference
    }

    $components = @()
    $entryPoints = @()
    $hardCodedWorkbookReferences = @()
    foreach ($component in @($Workbook.VBProject.VBComponents)) {
        $codeModule = $component.CodeModule
        $lineCount = [int]$codeModule.CountOfLines
        $code = if ($lineCount -gt 0) { [string]$codeModule.Lines(1, $lineCount) } else { '' }
        $typeName = Get-VbaComponentTypeName ([int]$component.Type)
        $item = [ordered]@{
            Name = [string]$component.Name
            Type = [int]$component.Type
            TypeName = $typeName
            LineCount = $lineCount
            CodeSha256 = Get-Sha256Text $code
            # VBA is case insensitive and the editor rewrites member names to its
            # own casing when it compiles a project, so the source round trip is
            # compared without case.
            CodeSha256IgnoringCase = Get-Sha256Text ($code.ToLowerInvariant())
        }

        if ($typeName -eq 'UserForm') {
            $controls = @()
            $designer = $null
            try {
                $designer = $component.Designer
                foreach ($control in @($designer.Controls)) {
                    $controls += [ordered]@{ Name = [string]$control.Name; Type = [string]$control.GetType().FullName; ProgId = $(try { [string]$control.progID } catch { '' }) }
                    Release-ComObject $control
                }
                $item.Caption = $(try { [string]$designer.Caption } catch { '' })
                $item.Controls = $controls
            } catch {
                $item.Caption = ''
                $item.Controls = @()
                $item.ControlInventoryError = $_.Exception.Message
            } finally {
                Release-ComObject $designer
            }
        }

        $componentEntries = [regex]::Matches($code, '(?im)^\s*(Public\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)')
        foreach ($match in $componentEntries) {
            $entryPoints += [ordered]@{ Component = [string]$component.Name; Declaration = $match.Value.Trim(); Name = $match.Groups[5].Value }
        }
        foreach ($pattern in @('Workbooks\s*\(\s*"[^"]+"\s*\)', 'Windows\s*\(\s*"[^"]+"\s*\)', 'Application\.Run\s+"[^"]+"')) {
            foreach ($match in [regex]::Matches($code, $pattern, 'IgnoreCase')) {
                $hardCodedWorkbookReferences += [ordered]@{ Component = [string]$component.Name; Text = $match.Value; Classification = 'Requires review' }
            }
        }
        $components += $item
        Release-ComObject $codeModule
        Release-ComObject $component
    }

    return [ordered]@{
        SchemaVersion = 1
        CapturedAtUtc = [DateTime]::UtcNow.ToString('o')
        Workbook = [ordered]@{
            FileName = [IO.Path]::GetFileName($Path)
            FullPath = [IO.Path]::GetFullPath($Path)
            FileSize = (Get-Item -LiteralPath $Path).Length
            FileSha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
            ExcelVersion = [string]$Excel.Version
            VBProjectName = [string]$Workbook.VBProject.Name
            VBProjectProtection = [int]$Workbook.VBProject.Protection
        }
        Worksheets = $worksheets
        NamedRanges = $names
        VbaReferences = $references
        VbaComponents = $components
        EntryPoints = $entryPoints
        HardCodedWorkbookReferences = $hardCodedWorkbookReferences
    }
}

function Invoke-WithExcelWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][scriptblock]$Action,
        [switch]$ReadOnly
    )
    $excel = $null
    $workbook = $null
    try {
        $excel = New-ExcelApplication
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $workbook = $excel.Workbooks.Open($Path, 0, [bool]$ReadOnly)
        & $Action $workbook $excel -ErrorAction Stop
    } finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

Export-ModuleMember -Function Release-ComObject, New-ExcelApplication, Get-WorkbookInventory, Invoke-WithExcelWorkbook, Get-GdtErrorReportHeaders, Get-GdtRelatedInvoiceHeaders, Get-Sha256Text, Get-VbaCodeTextFromSource
