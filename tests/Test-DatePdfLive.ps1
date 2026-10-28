[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BuiltWorkbook,
    [Parameter(Mandatory)][string]$Username,
    [Parameter(Mandatory)][string]$Password,
    [switch]$WithSearchCaptcha,
    [switch]$PdfSmoke,
    [switch]$ApiPdfSmoke,
    [switch]$ProbeQuerySold
)
# Opt-in, read-only API checks. Credentials/token exist only in an unsaved
# workbook instance; only status codes and counts leave this process.
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'build\ExcelBuild.Common.psm1') -Force -DisableNameChecking
$excel = $null; $book = $null; $module = $null
$liveRoot = Join-Path $env:TEMP ('hddt-datepdf-live-' + [guid]::NewGuid().ToString('N'))
if ($PdfSmoke -or $ApiPdfSmoke) { New-Item -ItemType Directory -Path $liveRoot | Out-Null }
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $BuiltWorkbook).Path, 0, $true)
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'modDatePdfLive'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public LiveUser As String, LivePass As String
Public UseSearchCaptcha As Boolean
Public SmokePdf As Boolean, SmokeFolder As String
Public SmokeApiPdf As Boolean
Public ProbeOne As Boolean
Public Function LiveReadOnlyCheck() As String
    Dim http As Object, payload As Object, body As String, key As String, captcha As String
    Dim token As String, source As Variant, direction As Variant, days As Variant, endpoint As String
    Dim result As clsGdtRequestResult, report As String, limit As Long
    Dim sample As Object
    On Error GoTo Failed
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 15000, 15000, 15000, 30000
    http.Open "GET", BASE_URL & "/captcha", False
    http.setRequestHeader "Request-Id", CreateRequestID(): http.send
    If http.Status <> 200 Then LiveReadOnlyCheck = "captcha HTTP " & http.Status: Exit Function
    body = http.responseText
    Set payload = ParseGdtAuthObject(body)
    key = GdtRequiredAuthText(payload, "key"): captcha = detectSVGCaptcha(body)
    If Len(captcha) <> 6 Then LiveReadOnlyCheck = "captcha recognition unavailable": Exit Function
    http.Open "POST", BASE_URL & "/security-taxpayer/authenticate", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Request-Id", CreateRequestID()
    http.send BuildGdtLoginJson(LiveUser, LivePass, captcha, key)
    LivePass = vbNullString
    If http.Status <> 200 Then LiveReadOnlyCheck = "login HTTP " & http.Status: Exit Function
    Set payload = ParseGdtAuthObject(http.responseText)
    token = GdtRequiredAuthText(payload, "token")
    report = "login OK"
    ResetGdtRequestSession: ResetGdtOperationControl
    For Each source In Array("query", "sco-query")
        For Each direction In Array("sold", "purchase")
            For Each days In Array(5, 10, 30)
                endpoint = "https://hoadondientu.gdt.gov.vn/api/" & source & "/invoices/" & direction & _
                    "?sort=tdlap:desc&size=1&search=tdlap=ge=01/10/2026T00:00:00;tdlap=le=" & _
                    Format$(CLng(days), "00") & "/10/2026T23:59:59"
                If UseSearchCaptcha Then
                    Set result = ExecuteGdtListRequest(endpoint, token, 1)
                Else
                    Set result = ExecuteGdtRequest("GET", endpoint, token, , , , False, 1)
                End If
                limit = GdtReportedDayLimit(result.StatusCode, result.ErrorMessage)
                report = report & vbLf & source & "/" & direction & " " & days & "d HTTP " & result.StatusCode
                If limit > 0 Then report = report & " limit=" & limit
                If result.Success Then
                    Set payload = JsonConverter.ParseJSON(result.ResponseText)
                    report = report & " rows=" & payload("datas").Count
                    If sample Is Nothing And source = "query" And direction = "sold" Then
                        If payload("datas").Count > 0 Then Set sample = payload("datas")(1)
                    End If
                    If SmokeApiPdf And days = 5 And payload("datas").Count > 0 Then _
                        report = report & vbLf & LiveApiPdfSmoke(payload("datas")(1), token, CStr(source), CStr(direction))
                ElseIf result.StatusCode = 400 And limit = 0 Then
                    report = report & " other-validation-error"
                End If
                WaitGdtMilliseconds 800
                If ProbeOne Then GoTo ListsDone
            Next days
        Next direction
    Next source
ListsDone:
    If SmokePdf And Not sample Is Nothing Then report = report & vbLf & LivePdfSmoke(sample, token)
    LiveReadOnlyCheck = report
    Exit Function
Failed:
    LivePass = vbNullString
    LiveReadOnlyCheck = report & vbLf & "check failed; VBA error " & Err.Number
End Function
Private Function LiveApiPdfSmoke(ByVal sample As Object, ByVal token As String, ByVal source As String, ByVal direction As String) As String
    Dim result As clsGdtRequestResult, endpoint As String, identity As String, path As String, fso As Object, data As Object, key As Variant
    On Error GoTo Failed
    endpoint = "https://hoadondientu.gdt.gov.vn/api/" & source & "/invoices/detail?nbmst=" & sample("nbmst") & _
        "&khhdon=" & sample("khhdon") & "&shdon=" & sample("shdon") & "&khmshdon=" & sample("khmshdon")
    Set result = ExecuteGdtRequest("GET", endpoint, token, , , , False, 1)
    If Not result.Success Then LiveApiPdfSmoke = "API PDF " & source & "/" & direction & ": detail HTTP " & result.StatusCode: Exit Function
    identity = GdtPdfIdentity(CStr(sample("nbmst")), CStr(sample("khmshdon")), CStr(sample("khhdon")), CStr(sample("shdon")))
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(SmokeFolder, GdtInvoiceFileBase(CStr(sample("nbmst")), CStr(sample("khmshdon")), CStr(sample("khhdon")), CStr(sample("shdon"))))
    If fso.FileExists(path & ".xml") Then Err.Raise 5
    GdtExportDetailPdf result.ResponseText, path & ".pdf", JsonConverter.ConvertToJson(sample), identity
    If Not GdtIsPdfFile(path & ".pdf") Then Err.Raise 5
    LiveApiPdfSmoke = "API PDF " & source & "/" & direction & ": PASS without XML"
    Exit Function
Failed:
    LiveApiPdfSmoke = "API PDF " & source & "/" & direction & ": failed; VBA error " & Err.Number
    If InStr(Err.Description, "Missing API PDF") = 1 Then LiveApiPdfSmoke = LiveApiPdfSmoke & "; " & Err.Description
    On Error Resume Next
    Set data = JsonConverter.ParseJSON(result.ResponseText)
    For Each key In Array("tgtcthue", "tgtthue", "tgtttbso")
        LiveApiPdfSmoke = LiveApiPdfSmoke & "; " & key & " present(detail/list)=" & GdtJsonHasValue(data,CStr(key)) & "/" & GdtJsonHasValue(sample,CStr(key))
    Next key
End Function
Private Function LivePdfSmoke(ByVal sample As Object, ByVal token As String) As String
    Dim result As clsGdtRequestResult, endpoint As String, stream As Object, base As String, fso As Object
    Dim zipPath As String, xmlPath As String, pdfPath As String, identity As String
    On Error GoTo Failed
    endpoint = "https://hoadondientu.gdt.gov.vn/api/query/invoices/export-xml?nbmst=" & sample("nbmst") & _
        "&khhdon=" & sample("khhdon") & "&shdon=" & sample("shdon") & "&khmshdon=" & sample("khmshdon")
    Set result = ExecuteGdtBinaryRequest(endpoint, token, 1, "XML")
    If Not result.Success Then LivePdfSmoke = "PDF smoke: XML HTTP " & result.StatusCode: Exit Function
    base = GdtInvoiceFileBase(CStr(sample("nbmst")), CStr(sample("khmshdon")), CStr(sample("khhdon")), CStr(sample("shdon")))
    identity = GdtPdfIdentity(CStr(sample("nbmst")), CStr(sample("khmshdon")), CStr(sample("khhdon")), CStr(sample("shdon")))
    Set fso = CreateObject("Scripting.FileSystemObject")
    zipPath = fso.BuildPath(SmokeFolder, base & ".zip")
    xmlPath = fso.BuildPath(SmokeFolder, base & ".xml"): pdfPath = fso.BuildPath(SmokeFolder, base & ".pdf")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write result.ResponseBody: stream.SaveToFile zipPath, 2: stream.Close
    Unzip zipPath, SmokeFolder
    GdtExportInvoicePdf xmlPath, pdfPath, identity
    If Not GdtIsPdfFile(pdfPath) Then Err.Raise 5
    LivePdfSmoke = "PDF smoke: XML downloaded, extracted and rendered successfully"
    Exit Function
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    LivePdfSmoke = "PDF smoke failed; VBA error " & Err.Number
End Function
'@)
    $userLiteral = $Username.Replace('"','""')
    $passLiteral = $Password.Replace('"','""')
    $folderLiteral = $liveRoot.Replace('"','""')
    $module.CodeModule.AddFromString("Public Function LiveEntry() As String`r`nLiveUser = `"$userLiteral`"`r`nLivePass = `"$passLiteral`"`r`nUseSearchCaptcha = $($WithSearchCaptcha.IsPresent)`r`nSmokePdf = $($PdfSmoke.IsPresent)`r`nSmokeApiPdf = $($ApiPdfSmoke.IsPresent)`r`nProbeOne = $($ProbeQuerySold.IsPresent)`r`nSmokeFolder = `"$folderLiteral`"`r`nLiveEntry = LiveReadOnlyCheck()`r`nEnd Function")
    $result = [string]$excel.Run("'$($book.Name)'!LiveEntry")
    Write-Output $result
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $module; Release-ComObject $book; Release-ComObject $excel
    if (($PdfSmoke -or $ApiPdfSmoke) -and (Test-Path -LiteralPath $liveRoot)) {
        $resolved = [IO.Path]::GetFullPath($liveRoot)
        $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
        if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved) -match '^hddt-datepdf-live-[a-f0-9]{32}$') {
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
}
