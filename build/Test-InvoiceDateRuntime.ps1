[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BuiltWorkbook,
    [string]$BaselineWorkbook
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$excel = $null; $workbook = $null; $module = $null
$results = [ordered]@{}
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    if ($BaselineWorkbook) {
        $baselinePath = (Resolve-Path -LiteralPath $BaselineWorkbook).Path
        $workbook = $excel.Workbooks.Open($baselinePath, 0, $true)
        $versionModule = $workbook.VBProject.VBComponents.Item('modVersion').CodeModule
        if ($versionModule.Lines(1, $versionModule.CountOfLines) -notmatch 'CURRENT_VERSION As String = "6\.7\.5"') {
            throw 'Baseline must be the v6.7.5 workbook.'
        }
        Release-ComObject $versionModule
        $module = $workbook.VBProject.VBComponents.Add(1)
        $module.Name = 'modInvoiceDateRegression'
        $module.CodeModule.AddFromString(@'
Public Function BaselineDate() As String
    BaselineDate = Format$(ISODateValue("2026-09-11T17:00:00.000Z"), "yyyy-mm-dd")
End Function
'@)
        $actual = [string]$excel.Run("'$($workbook.Name)'!BaselineDate")
        if ($actual -ne '2026-09-11') { throw "v6.7.5 regression did not reproduce: $actual" }
        $results['Baseline675'] = 'REPRODUCED: UTC 17:00 => 2026-09-11; expected Vietnam 2026-09-12'
        Write-Output $results['Baseline675']
        Release-ComObject $module; $module = $null
        $workbook.Close($false); Release-ComObject $workbook; $workbook = $null
    }

    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $true)
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'modInvoiceDateRegression'
    $module.CodeModule.AddFromString(@'
Option Explicit
Private stage As String
Private count As Long
Private Sub ExpectDate(ByVal inputValue As Variant, ByVal expected As Date)
    stage = CStr(inputValue)
    If ISODateValue(inputValue) <> expected Then Err.Raise 5, , "Unexpected calendar date"
    count = count + 1
End Sub
Private Sub ExpectInvalid(ByVal inputValue As Variant)
    Dim result As Date, failure As Long
    stage = "invalid input"
    On Error Resume Next
    result = ISODateValue(inputValue)
    failure = Err.Number
    On Error GoTo 0
    If failure <> vbObjectError + 513 Then Err.Raise 5, , "Invalid date/time was not rejected"
    count = count + 1
End Sub
Public Function InvoiceDateRegression() As String
    On Error GoTo Failed
    ExpectDate "2026-09-11T16:59:59.999Z", DateSerial(2026, 9, 11)
    ExpectDate "2026-09-11T17:00:00.000Z", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12T00:00:00Z", DateSerial(2026, 9, 12)
    ExpectDate "2026-10-05T17:00:00Z", DateSerial(2026, 10, 6)
    ExpectDate "2026-12-31T17:00:00Z", DateSerial(2027, 1, 1)
    ExpectDate "2024-02-28T17:00:00Z", DateSerial(2024, 2, 29)
    ExpectDate "2024-02-29T17:00:00Z", DateSerial(2024, 3, 1)
    ExpectDate "2026-09-12T00:00:00+07:00", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12T00:00:00+0700", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-11T12:00:00-05:00", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-11T12:00:00-0500", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12T00:30:00+08:00", DateSerial(2026, 9, 11)
    ExpectDate "2026-09-11T22:30:00+05:30", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12T23:59:59", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12T23:59:59.123", DateSerial(2026, 9, 12)
    ExpectDate " 2026-09-11t17:00:00z ", DateSerial(2026, 9, 12)
    ExpectDate "2026-09-12 00:00:00+07:00", DateSerial(2026, 9, 12)
    ExpectDate DateSerial(2026, 9, 12) + TimeSerial(23, 59, 59), DateSerial(2026, 9, 12)
    ExpectInvalid "2026-02-29"
    ExpectInvalid "2026-13-01"
    ExpectInvalid "2026-09-00"
    ExpectInvalid "2026-09-12T24:00:00Z"
    ExpectInvalid "2026-09-12T00:60:00Z"
    ExpectInvalid "2026-09-12T00:00:60Z"
    ExpectInvalid "2026-09-12T00:00:00+07:60"
    ExpectInvalid "2026-09-12T00:00:00+24:00"
    ExpectInvalid "2026-09-12T00:00:00+0:700"
    ExpectInvalid "2026-09-12T00:00:00.Z"
    ExpectInvalid "2026-09-12T00:00:00junk"
    ExpectInvalid "2026-09-12junk"
    ExpectInvalid "12/09/2026"
    ExpectInvalid ""
    ExpectInvalid Null
    ExpectInvalid Empty
    InvoiceDateRegression = "PASS: " & count & " cases"
    Exit Function
Failed:
    InvoiceDateRegression = "FAIL: " & stage & " | " & Err.Number & " " & Err.Description
End Function
Public Function RelatedNoticeDateRegression() As String
    On Error GoTo Failed
    Dim direction As Long, ws As Worksheet, suffix As String
    For direction = 1 To 2
        If direction = 1 Then suffix = "Mua" Else suffix = "Ban"
        Set ws = ThisWorkbook.Sheets("TongHopHD_" & suffix)
        WriteRelatedInformationResponse "{""mtthdtbssrs"":[{""ten"":""FIXTURE"",""ngay"":""2026-09-12T00:30:00+08:00""}]}", direction, 300
        If InStr(CStr(ws.Cells(300, 64).Value2), "11/09/2026") = 0 Then Err.Raise 5, , "Offset notice date"
        WriteRelatedInformationResponse "{""mtthdtbssrs"":[{""ten"":""FIXTURE"",""ngay"":""2026-09-11T17:00:00Z""}]}", direction, 300
        If InStr(CStr(ws.Cells(300, 64).Value2), "12/09/2026") = 0 Then Err.Raise 5, , "UTC notice date"
    Next direction
    RelatedNoticeDateRegression = "PASS"
    Exit Function
Failed:
    RelatedNoticeDateRegression = "FAIL: " & Err.Description
End Function
'@)
    foreach ($entry in @('InvoiceDateRegression', 'RelatedNoticeDateRegression')) {
        $actual = [string]$excel.Run("'$($workbook.Name)'!$entry")
        $results[$entry] = $actual
        Write-Output "$entry=$actual"
        if ($actual -notlike 'PASS*') { throw $actual }
    }
    $results | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath (Join-Path (Split-Path -Parent $BuiltWorkbook) 'invoice-date-runtime-result.json')
    Write-Output 'INVOICE-DATE RUNTIME TEST PASSED'
} finally {
    Release-ComObject $module
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $workbook; Release-ComObject $excel
}
