[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BuiltWorkbook,
    [switch]$Audit
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$fixtureRoot = Join-Path $env:TEMP ('hddt-extended-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$xmlDir = Join-Path $fixtureRoot 'xml'
New-Item -ItemType Directory -Path $xmlDir | Out-Null
[IO.File]::WriteAllText((Join-Path $xmlDir '00-no-items.xml'), '<HDon><DLHDon Id="BAD"><NDHDon><NBan><Ten>SHOULD-NOT-LEAK</Ten></NBan></NDHDon></DLHDon></HDon>')
[IO.File]::WriteAllText((Join-Path $xmlDir '01-valid.xml'), '<HDon><DLHDon Id="VALID"><TTChung><SHDon>1</SHDon></TTChung><NDHDon><DSHHDVu><HHDVu><STT>1</STT><ThTien>100</ThTien><TSuat>10</TSuat></HHDVu></DSHHDVu></NDHDon></DLHDon></HDon>')
[IO.File]::WriteAllText((Join-Path $xmlDir '02-valid.XML'), '<HDon><DLHDon Id="UPPER"><TTChung><SHDon>2</SHDon></TTChung><NDHDon><DSHHDVu><HHDVu><STT>1</STT><ThTien>100</ThTien><TSuat>10</TSuat></HHDVu></DSHHDVu></NDHDon></DLHDon></HDon>')
$excel = $null; $workbook = $null; $module = $null; $formCode = $null; $server = $null
$results = [ordered]@{}
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $macroPrefix = "'" + $workbook.Name + "'!"
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'modExtendedReviewTest'
    $baseline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-DetailInvoiceRuntime.ps1') -Raw -Encoding UTF8
    $detail = [regex]::Match($baseline, '(?s)Private Function RunDetailInvoiceFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    $summary = [regex]::Match($baseline, '(?s)Private Function RunSummaryDateFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    if (-not $detail -or -not $summary) { throw 'Missing synthetic invoice schemas' }
    $module.CodeModule.AddFromString("Public Function ExtendedDetailJson() As String`r`n Dim sampleJson As String, sellerTaxId As String, buyerTaxId As String`r`n sellerTaxId = `"0000000001`": buyerTaxId = `"0000000002`"`r`n$detail`r`n ExtendedDetailJson = sampleJson`r`nEnd Function`r`n" +
        "Public Function ExtendedSummaryJson() As String`r`n Dim sampleJson As String`r`n$summary`r`n ExtendedSummaryJson = sampleJson`r`nEnd Function`r`n")
    $module.CodeModule.AddFromString(@'
Option Explicit
Public ExtendedBaseUrl As String
Public ExtendedXmlFolder As String
Private Sub InitializeLookups()
    LinkTraCuu
    tenCotTraCuu
    arrTrangThai = Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = Sheets("LinkTraCuu").Range("O2:O11").Value
End Sub
Public Function ExtendedSummaryRollback() As String
    On Error GoTo Failed
    InitializeLookups
    Dim page As Object, second As Object, r As Long, seq As Long, failure As Long
    Dim ws As Worksheet
    Set ws = Sheets("TongHopHD_Mua")
    ws.Range("A450:BL452").ClearContents
    ws.Cells(449, 5).Value2 = "KEEP"
    ws.Cells(450, 64).Value2 = "KEEP-TARGET"
    Set page = JsonConverter.ParseJSON(ExtendedSummaryJson())
    Set second = JsonConverter.ParseJSON(ExtendedSummaryJson())("datas")(1)
    second("tdlap") = "2026-02-30"
    page("datas").Add second
    r = 450: seq = 1
    On Error Resume Next
    ghiExcel_TongHop JsonConverter.ConvertToJson(page), r, 1, seq
    failure = Err.Number
    On Error GoTo Failed
    If failure = 0 Then Err.Raise 5, , "invalid invoice was accepted"
    If r <> 450 Or seq <> 1 Or Application.CountA(ws.Range("A450:BL451")) <> 1 Then Err.Raise 5, , "page left partial rows/cursor: " & r & "/" & seq & "; count=" & Application.CountA(ws.Range("A450:BL451"))
    If ws.Cells(450, 64).Value2 <> "KEEP-TARGET" Then Err.Raise 5, , "rollback lost pre-existing target data"
    If ws.Cells(449, 5).Value2 <> "KEEP" Then Err.Raise 5, , "previous invoice changed"
    ExtendedSummaryRollback = "PASS"
    Exit Function
Failed:
    ExtendedSummaryRollback = "FAIL: " & Err.Description
End Function
Public Function ExtendedDetailMetadata() As String
    On Error GoTo Failed
    InitializeLookups
    Dim invoice As Object, item As Object, metadata As Object, entries As Collection, r As Long
    Dim ws As Worksheet
    Set ws = Sheets("ChiTietHD_Mua")
    ws.Range("A460:AH461").ClearContents
    Sheets("LinkTraCuu").Range("O14").Value = "fixture_tax"
    Sheets("LinkTraCuu").Range("O15").Value = "fixture_total"
    Set invoice = JsonConverter.ParseJSON(ExtendedDetailJson())
    invoice("tgtthue") = 70
    For Each item In invoice("hdhhdvu")
        item("tthue") = 35: item("thtcthue") = 235
        Set entries = New Collection
        Set metadata = JsonConverter.ParseJSON("{""ttruong"":""fixture_tax"",""dlieu"":null}")
        entries.Add metadata
        Set metadata = JsonConverter.ParseJSON("{""ttruong"":""fixture_total""}")
        entries.Add metadata
        Set item("ttkhac") = entries
    Next item
    r = 460
    ghiExcel_ChiTiet JsonConverter.ConvertToJson(invoice), r, "mua"
    If ws.Cells(460, 28).Value2 <> 35 Or ws.Cells(460, 29).Value2 <> 235 Then Err.Raise 5, , "null/missing metadata replaced direct tax/total"
    If ws.Cells(460, 31).Value2 <> "Tien thue ok" Then Err.Raise 5, , "wrong tax reconciliation"
    ExtendedDetailMetadata = "PASS"
    Exit Function
Failed:
    ExtendedDetailMetadata = "FAIL: " & Err.Description
End Function
Public Function ExtendedErrorIdentity() As String
    On Error GoTo Failed
    Dim ws As Worksheet
    Set ws = EnsureGdtErrorReportSheet()
    ws.Rows("2:100").ClearContents
    GdtReportAccount = "fixture-account"
    UpsertGdtErrorReport "mua", "query", "0000000001", "1", "FIXTURE", "42", Empty, "DETAIL", "/fixture", 500, "fixture", 1, 0, "Failed"
    If CStr(ws.Cells(2, 5).Value2) <> "0000000001" Then Err.Raise 5, , "tax ID lost leading zeros: " & ws.Cells(2, 5).Value2
    UpsertGdtErrorReport "mua", "query", "0000000001", "1", "FIXTURE", "42", Empty, "DETAIL", "/fixture", 500, "fixture", 2, 0, "Failed"
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row <> 2 Then Err.Raise 5, , "duplicate error rows"
    MarkGdtRetrySuccess "mua", "query", "0000000001", "1", "FIXTURE", "42", "DETAIL", 3
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row <> 1 Then Err.Raise 5, , "error not cleared"
    GdtReportAccount = ""
    ExtendedErrorIdentity = "PASS"
    Exit Function
Failed:
    GdtReportAccount = ""
    ExtendedErrorIdentity = "FAIL: " & Err.Description
End Function
Public Function ExtendedXmlIsolation() As String
    On Error GoTo Failed
    Dim ws As Worksheet
    Set ws = Sheets("ChiTietHD_Mua_XML")
    ws.Range("A3:AI100").ClearContents
    parseXML ExtendedXmlFolder, True, False
    If ws.Cells(3, 1).Value2 <> "VALID" Or ws.Cells(3, 7).Value2 <> "" Then Err.Raise 5, , "invalid XML invoice leaked common fields: id=" & ws.Cells(3, 1).Value2 & "; seller=" & ws.Cells(3, 7).Value2
    If ws.Cells(4, 1).Value2 <> "UPPER" Then Err.Raise 5, , "uppercase XML file skipped"
    If ws.Cells(5, 1).Value2 <> "" Then Err.Raise 5, , "invalid file produced a row"
    ExtendedXmlIsolation = "PASS"
    Exit Function
Failed:
    ExtendedXmlIsolation = "FAIL: " & Err.Description
End Function
Public Function ExtendedRetryAfter() As String
    If CalculateGdtRetrySeconds(1, 120, 0) < 120 Then
        ExtendedRetryAfter = "FAIL: server Retry-After 120 was shortened"
    Else
        ExtendedRetryAfter = "PASS"
    End If
End Function
Public Function ExtendedFilenameIdentity() As String
    If GdtInvoiceFileBase("<A", "1", "FIXTURE", "42") = GdtInvoiceFileBase(ChrW(&H3CA), "1", "FIXTURE", "42") Then
        ExtendedFilenameIdentity = "FAIL: variable-width escaping aliases distinct identifiers"
    Else
        ExtendedFilenameIdentity = "PASS"
    End If
End Function
Public Function ExtendedTransportRetry() As String
    On Error GoTo Failed
    ResetGdtOperationControl
    ResetGdtRequestSession
    Dim response As clsGdtRequestResult
    Set response = ExecuteGdtRequest("GET", ExtendedBaseUrl & "/drop", "", , , , , 2)
    If response.Success Or response.Attempts <> 2 Or Not response.ShouldQueueFinalRetry Then Err.Raise 5, , "transport retries not exhausted safely"
    ExtendedTransportRetry = "PASS"
    Exit Function
Failed:
    ExtendedTransportRetry = "FAIL: " & Err.Description
End Function
'@)
    $formCode = $workbook.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $formCode.AddFromString(@'
Public Function ExtendedListRetry(ByVal endpoint As String) As String
    On Error GoTo Failed
    ResetGdtOperationControl
    ResetGdtRequestSession
    LinkTraCuu
    tenCotTraCuu
    arrTrangThai = Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = Sheets("LinkTraCuu").Range("O2:O11").Value
    Me.txtSleep.Value = "0"
    Set mFinalRetryQueue = New Collection
    mFinalRetryCooldownDone = True
    Dim queued As New clsGdtRetryItem, buffer As Variant, count As Long, r As Long, seq As Long
    queued.Direction = CurrentDirectionName(): queued.ApiSource = "query"
    queued.ResponseKind = "LIST": queued.Stage = UniConvert("Laasy danh sasch")
    queued.Endpoint = endpoint
    mFinalRetryQueue.Add queued
    ReDim buffer(0 To 10, 0 To 7)
    r = 470: seq = 1
    Sheets("TongHopHD_Mua").Range("A470:BL480").ClearContents
    ProcessQueuedListRetries buffer, count, r, seq, 1
    If count <> 1 Or r <> 471 Then Err.Raise 5, , "retried current state repeated invoice: " & count & "; " & Sheets("BaoCao_LoiTaiHD").Cells(2, 13).Value2
    ExtendedListRetry = "PASS"
    Exit Function
Failed:
    ExtendedListRetry = "FAIL: " & Err.Description
End Function
Public Function ExtendedRelationRetry(ByVal baseUrl As String, ByVal mode As Long) As String
    On Error GoTo Failed
    ResetGdtOperationControl
    ResetGdtRequestSession
    Set mFinalRetryQueue = New Collection
    mFinalRetryCooldownDone = True
    Dim queued As New clsGdtRetryItem, ws As Worksheet
    queued.Direction = CurrentDirectionName(): queued.ApiSource = "query"
    queued.SellerTaxCode = "seller-fixture": queued.TemplateCode = "1"
    queued.InvoiceSeries = "FIXTURE": queued.InvoiceNumber = "42": queued.TargetRow = 480
    If mode = 2 Then
        queued.ResponseKind = "RELATIVE": queued.Stage = "RELATIVE"
        queued.Endpoint = baseUrl & "/relative?invalid=yes"
    Else
        queued.ResponseKind = "RELATED": queued.Stage = "RELATED"
        If mode = 0 Then queued.Endpoint = baseUrl & "/bad-related" Else queued.Endpoint = baseUrl & "/error-related"
    End If
    Set ws = EnsureGdtErrorReportSheet()
    ws.Rows("2:100").ClearContents
    UpsertGdtErrorReport queued.Direction, "query", queued.SellerTaxCode, "1", "FIXTURE", "42", Empty, queued.Stage, queued.Endpoint, 500, "fixture", 1, 0, "Failed"
    mFinalRetryQueue.Add queued
    ProcessQueuedRelationRetries 1
    If ws.Cells(2, 10).Value2 <> queued.Stage Then Err.Raise 5, , "HTTP 200 with invalid related response cleared original error; mode=" & mode
    If mode = 2 Then queued.Endpoint = baseUrl & "/relative?valid=yes" Else queued.Endpoint = baseUrl & "/empty-related"
    ProcessQueuedRelationRetries 1
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row <> 1 Then Err.Raise 5, , "valid related retry left stale parse/download errors"
    ExtendedRelationRetry = "PASS"
    Exit Function
Failed:
    ExtendedRelationRetry = "FAIL: " & Err.Description
End Function
Public Function ExtendedDateInput() As String
    On Error GoTo Failed
    Me.txtTuNgay.Value = "01/01/" & Year(Date)
    Me.txtDenNgay.Value = "31/02/2026"
    txtDenNgay_AfterUpdate
    If Me.txtDenNgay.Value <> "31/02/2026" Or Not Me.lblSaiNgay.Visible Then Err.Raise 5, , "invalid end date silently became today"
    If lietKeThoiGian() Then Err.Raise 5, , "invalid date created a search period"
    Me.txtTuNgay.Value = "31/02/2026"
    txtTuNgay_AfterUpdate
    If Me.txtTuNgay.Value <> "31/02/2026" Or Not Me.lblSaiNgay.Visible Then Err.Raise 5, , "invalid start date silently became today"
    Me.txtDenNgay.Value = ""
    txtDenNgay_AfterUpdate
    If Me.txtDenNgay.Value <> "" Then Err.Raise 5, , "cleared date silently became today"
    Me.txtTuNgay.Value = "12/09/2026"
    txtTuNgay_AfterUpdate
    Me.txtDenNgay.Value = "12/09/2026"
    txtDenNgay_AfterUpdate
    If Not lietKeThoiGian() Or Me.lblSaiNgay.Visible Then Err.Raise 5, , "valid date pair failed"
    ExtendedDateInput = "PASS"
    Exit Function
Failed:
    ExtendedDateInput = "FAIL: " & Err.Description
End Function
'@)
    $module.CodeModule.AddFromString(@'
Public Function ExtendedPagination() As String
    Dim form As New frmTaiHoaDon
    ExtendedPagination = form.ExtendedListRetry(ExtendedBaseUrl & "/purchase?sort=tdlap:desc&size=50&state=repeat&search=fixture")
    Unload form
End Function
Public Function ExtendedRelatedRecovery() As String
    Dim form As New frmTaiHoaDon, mode As Long, result As String
    For mode = 0 To 2
        result = form.ExtendedRelationRetry(ExtendedBaseUrl, mode)
        If result <> "PASS" Then Exit For
    Next mode
    ExtendedRelatedRecovery = result
    Unload form
End Function
Public Function ExtendedInputDates() As String
    Dim form As New frmTaiHoaDon
    ExtendedInputDates = form.ExtendedDateInput()
    Unload form
End Function
'@)

    $summaryJson = [string]$excel.Run("${macroPrefix}ExtendedSummaryJson")
    $summaryJson = $summaryJson.Substring(0, $summaryJson.Length - 1) + ',"state":"repeat"}'
    $probe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $probe.Start(); $port = ([Net.IPEndPoint]$probe.LocalEndpoint).Port; $probe.Stop()
    $ready = Join-Path $fixtureRoot 'ready'
    $server = Start-Job -ArgumentList $port,$ready,$summaryJson -ScriptBlock {
        param($port,$ready,$summaryJson)
        $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $port)
        $listener.Start()
        [IO.File]::WriteAllText($ready, 'ready')
        try {
            while ($true) {
                if (-not $listener.Pending()) { Start-Sleep -Milliseconds 5; continue }
                $client = $listener.AcceptTcpClient()
                try {
                    $reader = New-Object IO.StreamReader($client.GetStream())
                    $request = $reader.ReadLine()
                    $authorization = ''
                    do {
                        $header = $reader.ReadLine()
                        if ($header -match '^Authorization:\s*(.+)$') { $authorization = $Matches[1] }
                    } while ($null -ne $header -and $header.Length -gt 0)
                    $path = ($request -split ' ')[1]
                    if ($path -eq '/drop') { continue }
                    $responseText = $summaryJson
                    if ($path -eq '/token') {
                        $responseText = if ($authorization -eq 'Bearer fixture-original') { '{"ok":true}' } else { '{"ok":false}' }
                    } elseif ($path -eq '/bad-related') {
                        $responseText = '{bad json'
                    } elseif ($path -eq '/error-related') {
                        $responseText = '{"error":"fixture"}'
                    } elseif ($path -eq '/empty-related') {
                        $responseText = '{"hdtbssrses":[]}'
                    } elseif ($path -eq '/relative?invalid=yes') {
                        $responseText = '{}'
                    } elseif ($path -eq '/relative?valid=yes') {
                        $responseText = '[]'
                    }
                    $body = [Text.Encoding]::UTF8.GetBytes($responseText)
                    $headers = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 Fixture`r`nContent-Type: application/json`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n")
                    $stream = $client.GetStream()
                    $stream.Write($headers, 0, $headers.Length); $stream.Write($body, 0, $body.Length)
                } finally { $client.Close() }
            }
        } finally { $listener.Stop() }
    }
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while (-not (Test-Path -LiteralPath $ready)) {
        if ([DateTime]::UtcNow -gt $deadline) { throw 'Fixture server failed to start' }
        Start-Sleep -Milliseconds 50
    }
    $module.CodeModule.AddFromString("Public Function ExtendedConfigure() As Boolean`r`n ExtendedBaseUrl = `"http://127.0.0.1:$port`"`r`n ExtendedXmlFolder = `"$xmlDir`"`r`n ExtendedConfigure = True`r`nEnd Function`r`n")
    $entries = @('ExtendedErrorIdentity', 'ExtendedSummaryRollback', 'ExtendedDetailMetadata', 'ExtendedXmlIsolation', 'ExtendedRetryAfter', 'ExtendedTransportRetry', 'ExtendedPagination', 'ExtendedFilenameIdentity', 'ExtendedRelatedRecovery', 'ExtendedInputDates')
    if (-not $Audit) {
        $module.CodeModule.AddFromString(@'
Public Function ExtendedFrozenToken() As String
    On Error GoTo Failed
    ResetGdtOperationControl
    ResetGdtRequestSession
    Dim parsed As Object
    cToken = "fixture-original"
    Set parsed = JsonConverter.ParseJSON(ApiGet(ExtendedBaseUrl & "/token"))
    If Not parsed("ok") Then Err.Raise 5, , "legacy ApiGet call lost its global token"
    cToken = "fixture-changed"
    Set parsed = JsonConverter.ParseJSON(ApiGet(ExtendedBaseUrl & "/token", "fixture", "fixture-original"))
    If Not parsed("ok") Then Err.Raise 5, , "global token replaced operation snapshot"
    ExtendedFrozenToken = "PASS"
    Exit Function
Failed:
    ExtendedFrozenToken = "FAIL: " & Err.Description
End Function
'@)
        $entries += 'ExtendedFrozenToken'
    }
    # Editing VBA resets its globals; configure only after all injections.
    [void]$excel.Run("${macroPrefix}ExtendedConfigure")
    foreach ($entry in $entries) {
        $actual = [string]$excel.Run("${macroPrefix}$entry")
        $results[$entry] = $actual
        Write-Output "$entry=$actual"
    }
    $reportName = if ($Audit) { 'extended-review-before.json' } else { 'extended-review-runtime-result.json' }
    $report = Join-Path (Split-Path -Parent $BuiltWorkbook) $reportName
    $results | ConvertTo-Json | Set-Content -LiteralPath $report -Encoding UTF8
    if (-not $Audit -and @($results.Values | Where-Object { $_ -ne 'PASS' }).Count -gt 0) { throw "Extended regression failed: $report" }
} finally {
    Release-ComObject $formCode; Release-ComObject $module
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $workbook; Release-ComObject $excel
    if ($null -ne $server) { Stop-Job $server; Remove-Job $server -Force }
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -match '^hddt-extended-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
Write-Output $(if ($Audit) { 'EXTENDED REVIEW AUDIT COMPLETED' } else { 'EXTENDED REVIEW RUNTIME TEST PASSED' })
