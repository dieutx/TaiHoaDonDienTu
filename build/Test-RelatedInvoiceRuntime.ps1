[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$excel = $null
$workbook = $null
$testModule = $null
$sheet = $null
$reportRange = $null
$relatedRange = $null
$queryRelatedRange = $null
$scoRelatedRange = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $testModule = $workbook.VBProject.VBComponents.Add(1)
    $testModule.Name = 'modCodexRelatedSmokeTest'

    $sampleItems = @(
        '{"khmshdon":1,"khhdon":"C25THN","shdon":6,"khhdgoc":"C24THN","khmshdgoc":"1","shdgoc":18510,"lhdgoc":1,"tthai":4}',
        '{"khmshdon":1,"khhdon":"C25THN","shdon":30,"khhdgoc":"C25THN","khmshdgoc":"1","shdgoc":6,"lhdgoc":1,"tthai":4}',
        '{"khmshdon":1,"khhdon":"C25THN","shdon":31,"khhdgoc":"C25THN","khmshdgoc":"1","shdgoc":30,"lhdgoc":1,"tthai":4}',
        '{"khmshdon":1,"khhdon":"C25THN","shdon":5507,"khhdgoc":"C25THN","khmshdgoc":"1","shdgoc":31,"lhdgoc":1,"tthai":4}',
        '{"khmshdon":1,"khhdon":"C25THN","shdon":9996,"khhdgoc":"C25THN","khmshdgoc":"1","shdgoc":5507,"lhdgoc":1,"tthai":2}'
    )
    $codeLines = @(
        'Option Explicit',
        'Public Function CodexRelatedTrivial() As Boolean',
        'CodexRelatedTrivial = True',
        'End Function',
        'Public Function CodexTestResolvedErrorsRemoved() As Boolean',
        'On Error GoTo Failed',
        'Dim ws As Worksheet',
        'Dim lastRow As Long',
        'Dim rowIndex As Long',
        'Dim found As Boolean',
        'Set ws = EnsureGdtErrorReportSheet()',
        'UpsertGdtErrorReport "Ban ra", "query", "CODEX-TEST-MST", "1", "C99TST", "999", DateSerial(2026, 1, 1), "CODEX-TEST-STAGE", "/fixture", 500, "fixture", 1, 0, "failed"',
        'MarkGdtRetrySuccess "Ban ra", "query", "CODEX-TEST-MST", "1", "C99TST", "999", "CODEX-TEST-STAGE", 2',
        'lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',
        'For rowIndex = 2 To lastRow',
        '    If CStr(ws.Cells(rowIndex, 5).Value2) = "CODEX-TEST-MST" And CStr(ws.Cells(rowIndex, 8).Value2) = "999" Then found = True',
        'Next rowIndex',
        'CodexTestResolvedErrorsRemoved = Not found',
        'Exit Function',
        'Failed:',
        'CodexTestResolvedErrorsRemoved = False',
        'End Function',
        'Public Function CodexCallRelativeNothing() As Boolean',
        'On Error GoTo Failed',
        'WriteRelativeInvoiceData Nothing, 2, 3, "1", "C24THN", "18510"',
        'CodexCallRelativeNothing = True',
        'Exit Function',
        'Failed:',
        'CodexCallRelativeNothing = False',
        'End Function',
        'Public Function CodexCallRelatedEmpty() As Boolean',
        'On Error GoTo Failed',
        'WriteRelatedInformationResponse "[]", 2, 3',
        'CodexCallRelatedEmpty = True',
        'Exit Function',
        'Failed:',
        'CodexCallRelatedEmpty = False',
        'End Function',
        'Public Function CodexTestStatusSixRelated() As Boolean',
        'On Error GoTo Failed',
        'Dim queryJson As String',
        'Dim scoJson As String',
        'Dim queryText As String',
        'Dim scoText As String',
        'queryJson = "{""hdtbssrses"":[{""ngay"":""2024-01-17T17:00:00Z"",""ten"":""Thong bao query"",""loai"":1,""ldo"":""Ly do query"",""kqtnhan"":1}]}"',
        'scoJson = "{""mtthdtbssrs"":[{""ngay"":""2024-07-04T17:00:00Z"",""ten"":""Thong bao sco"",""loai"":1,""ldo"":""Ly do sco"",""kqtnhan"":0},{""ngay"":""2024-07-04T17:00:00Z"",""ten"":""Thong bao sco"",""loai"":1,""ldo"":""Ly do sco"",""kqtnhan"":0},{""ngay"":""2024-07-04T17:00:00Z"",""ten"":""Thong bao sco"",""loai"":1,""ldo"":""Ly do sco"",""kqtnhan"":1}]}"',
        'WriteNoRelativeInvoiceData 2, 5',
        'WriteRelatedInformationResponse queryJson, 2, 5',
        'WriteRelatedInformationResponse scoJson, 2, 6',
        'queryText = CStr(ThisWorkbook.Sheets("TongHopHD_Ban").Range("BM5").Value)',
        'scoText = CStr(ThisWorkbook.Sheets("TongHopHD_Ban").Range("BM6").Value)',
        'CodexTestStatusSixRelated = _',
        '    (ThisWorkbook.Sheets("TongHopHD_Ban").Range("BE5").Value = UniConvert("Khoong cos thoong tin hieen thij")) And _',
        '    (UBound(Split(queryText, vbCrLf)) = 0) And _',
        '    (InStr(1, queryText, "18/01/2024", vbTextCompare) > 0) And _',
        '    (InStr(1, queryText, UniConvert("Tisnh chaast Hury"), vbTextCompare) > 0) And _',
        '    (InStr(1, queryText, UniConvert("Cow quan thuees tieesp nhaajn."), vbTextCompare) > 0) And _',
        '    (UBound(Split(scoText, vbCrLf)) = 2) And _',
        '    (InStr(1, scoText, "05/07/2024", vbTextCompare) > 0) And _',
        '    (InStr(1, scoText, UniConvert("Cow quan thuees khoong tieesp nhaajn."), vbTextCompare) > 0) And _',
        '    (InStr(1, scoText, UniConvert("Cow quan thuees tieesp nhaajn."), vbTextCompare) > 0)',
        'Exit Function',
        'Failed:',
        'CodexTestStatusSixRelated = False',
        'End Function',
        'Public Function CodexTestRelatedInvoice() As Boolean',
        'On Error GoTo Failed',
        'Dim relatedItems As Object',
        'Dim reportText As String',
        'Dim sampleJson As String'
    )
    $codeLines += 'sampleJson = "["'
    for ($index = 0; $index -lt $sampleItems.Count; $index++) {
        $chunk = $sampleItems[$index].Replace('"', '""')
        if ($index -gt 0) { $codeLines += 'sampleJson = sampleJson & ","' }
        $codeLines += ('sampleJson = sampleJson & "' + $chunk + '"')
    }
    $codeLines += 'sampleJson = sampleJson & "]"'
    $codeLines += @(
        'Set relatedItems = JsonConverter.ParseJSON(sampleJson)',
        'WriteRelativeInvoiceData relatedItems, 2, 3, "1", "C24THN", "18510"',
        'WriteRelatedInformationResponse "[]", 2, 3',
        'reportText = CStr(ThisWorkbook.Sheets("TongHopHD_Ban").Range("BE3").Value)',
        'Dim reportLines As Variant',
        'reportLines = Split(reportText, vbCrLf)',
        'CodexTestRelatedInvoice = (InStr(1, reportText, "9996", vbTextCompare) > 0) And _',
        '    (InStr(1, reportText, "5507", vbTextCompare) > 0) And _',
        '    (InStr(1, reportText, "18510", vbTextCompare) > 0) And _',
        '    (UBound(Split(reportText, vbCrLf)) = 5) And _',
        '    (InStr(1, reportLines(0), "9996", vbTextCompare) > 0) And _',
        '    (InStr(1, reportLines(5), "18510", vbTextCompare) > 0) And _',
        '    (Len(ThisWorkbook.Sheets("TongHopHD_Ban").Range("BM3").Value) > 0)',
        'Exit Function',
        'Failed:',
        'CodexTestRelatedInvoice = False',
        'End Function'
    )
    $code = $codeLines -join "`r`n"
    $testModule.CodeModule.AddFromString($code)
    $trivialOk = [bool]$excel.Run("'$($workbook.Name)'!CodexRelatedTrivial")
    if (-not $trivialOk) { throw 'Trivial VBA smoke macro failed.' }
    $resolvedErrorRemoved = [bool]$excel.Run("'$($workbook.Name)'!CodexTestResolvedErrorsRemoved")
    if (-not $resolvedErrorRemoved) { throw 'Resolved error row was not removed.' }
    $relativeNothingOk = [bool]$excel.Run("'$($workbook.Name)'!CodexCallRelativeNothing")
    if (-not $relativeNothingOk) { throw 'Empty relative writer failed.' }
    $relatedEmptyOk = [bool]$excel.Run("'$($workbook.Name)'!CodexCallRelatedEmpty")
    if (-not $relatedEmptyOk) { throw 'Empty related writer failed.' }
    $ok = [bool]$excel.Run("'$($workbook.Name)'!CodexTestRelatedInvoice")
    $statusSixOk = [bool]$excel.Run("'$($workbook.Name)'!CodexTestStatusSixRelated")
    $sheet = $workbook.Worksheets.Item('TongHopHD_Ban')
    $reportRange = $sheet.Range('BE3')
    $relatedRange = $sheet.Range('BM3')
    $queryRelatedRange = $sheet.Range('BM5')
    $scoRelatedRange = $sheet.Range('BM6')
    $reportText = [string]$reportRange.Value2
    $relatedText = [string]$relatedRange.Value2
    $queryRelatedText = [string]$queryRelatedRange.Value2
    $scoRelatedText = [string]$scoRelatedRange.Value2
    Release-ComObject $reportRange; $reportRange = $null
    Release-ComObject $relatedRange; $relatedRange = $null
    Release-ComObject $queryRelatedRange; $queryRelatedRange = $null
    Release-ComObject $scoRelatedRange; $scoRelatedRange = $null
    Release-ComObject $sheet; $sheet = $null

    $workbook.VBProject.VBComponents.Remove($testModule)
    Release-ComObject $testModule; $testModule = $null
    $workbook.Close($false); Release-ComObject $workbook; $workbook = $null

    [ordered]@{
        Pass = ($ok -and $statusSixOk -and $resolvedErrorRemoved)
        ResolvedErrorRemoved = $resolvedErrorRemoved
        RelativeLineCount = @($reportText -split "`r?`n").Count
        ContainsInvoice18510 = $reportText.Contains('18510')
        ContainsInvoice9996 = $reportText.Contains('9996')
        FirstLineIsNewest = @($reportText -split "`r?`n")[0].Contains('9996')
        LastLineIsCurrent = @($reportText -split "`r?`n")[-1].Contains('18510')
        EmptyRelatedHandled = -not [string]::IsNullOrWhiteSpace($relatedText)
        StatusSixSkipsRelative = $statusSixOk
        QueryNoticeLineCount = @($queryRelatedText -split "`r?`n").Count
        ScoNoticeLineCount = @($scoRelatedText -split "`r?`n").Count
        ScoHasAcceptedAndRejected = $statusSixOk
        QueryPreview = $queryRelatedText
        ScoPreview = $scoRelatedText
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'tests\related-runtime-result.json') -Encoding utf8

    if (-not ($ok -and $statusSixOk -and $resolvedErrorRemoved)) { throw 'Related-invoice runtime smoke test failed.' }
    Write-Host 'RELATED-INVOICE RUNTIME TEST PASSED'
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $reportRange; Release-ComObject $relatedRange; Release-ComObject $queryRelatedRange; Release-ComObject $scoRelatedRange; Release-ComObject $sheet
    Release-ComObject $testModule; Release-ComObject $workbook; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
