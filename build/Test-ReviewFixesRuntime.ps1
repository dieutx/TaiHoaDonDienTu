[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$fixtureRoot = Join-Path $env:TEMP ('hddt-fixes-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$excel = $null; $workbook = $null; $server = $null
$results = [ordered]@{}

function Assert-Result([string]$Name, [object]$Actual) {
    $results[$Name] = [string]$Actual
    Write-Output "$Name=$Actual"
    if ([string]$Actual -ne 'PASS') { throw "$Name failed: $Actual" }
}

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression
    $zipPath = Join-Path $fixtureRoot 'fixture.zip'
    $zip = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in @('invoice.xml', 'nested/invoice.xml', 'invoice.html')) {
            $entry = $zip.CreateEntry($name)
            $writer = New-Object IO.StreamWriter($entry.Open())
            try { $writer.Write($(if ($name.EndsWith('.xml')) { '<HDon><DLHDon/></HDon>' } else { '<html>fixture</html>' })) }
            finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }
    $xmlDir = Join-Path $fixtureRoot 'xml'
    $outDir = Join-Path $fixtureRoot 'output'
    $unrelated = Join-Path $fixtureRoot 'Temporary Directory unrelated'
    New-Item -ItemType Directory -Path $xmlDir,$outDir,$unrelated | Out-Null
    [IO.File]::WriteAllText((Join-Path $unrelated 'keep.txt'), 'keep')
    $lines = '<HHDVu><STT>1</STT><ThTien>100</ThTien><TSuat>10</TSuat><TTKhac><TTin><TTruong>tax</TTruong><DLieu>20</DLieu></TTin></TTKhac></HHDVu>' +
        '<HHDVu><STT>2</STT><ThTien>100</ThTien><TSuat>10</TSuat><TTKhac><TTin><TTruong>tax</TTruong><DLieu>30</DLieu></TTin></TTKhac></HHDVu>' +
        '<HHDVu><STT>3</STT><ThTien>100</ThTien><TSuat>10</TSuat></HHDVu>' +
        '<HHDVu><STT>4</STT><ThTien>100</ThTien><TSuat>10</TSuat><TTKhac/></HHDVu>' +
        '<HHDVu><STT>5</STT><ThTien>100</ThTien><TSuat>10</TSuat><TTKhac><TTin><TTruong>tax</TTruong><DLieu>0</DLieu></TTin></TTKhac></HHDVu>' +
        '<HHDVu><STT>6</STT></HHDVu>'
    [IO.File]::WriteAllText((Join-Path $xmlDir 'fixture.xml'), '<HDon><DLHDon Id="fixture"><TTChung><SHDon>1</SHDon></TTChung><NDHDon><DSHHDVu>' + $lines + '</DSHHDVu></NDHDon></DLHDon></HDon>')

    $probe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $probe.Start(); $port = ([Net.IPEndPoint]$probe.LocalEndpoint).Port; $probe.Stop()
    $ready = Join-Path $fixtureRoot 'ready'
    $server = Start-Job -ArgumentList $port,$ready -ScriptBlock {
        param($port,$ready)
        $ErrorActionPreference = 'Stop'
        $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $port)
        $listener.Start()
        [IO.File]::WriteAllText($ready, 'ready')
        $pending = New-Object System.Collections.ArrayList
        $counts = @{}; $peak = 0
        try {
            while ($true) {
                if ($listener.Pending()) {
                    $client = $listener.AcceptTcpClient()
                    $client.ReceiveTimeout = 2000
                    $reader = New-Object IO.StreamReader($client.GetStream())
                    $request = $reader.ReadLine()
                    do { $header = $reader.ReadLine() } while ($null -ne $header -and $header.Length -gt 0)
                    $path = ($request -split ' ')[1]
                    if (-not $counts.ContainsKey($path)) { $counts[$path] = 0 }
                    $counts[$path]++
                    $status = 200; $body = 'PK-fixture'; $extra = ''; $delay = 0
                    if ($path.StartsWith('/slow')) { $delay = 1000 }
                    elseif ($path -eq '/rate' -and $counts[$path] -eq 1) { $status = 429; $body = 'rate limited'; $extra = "Retry-After: 1`r`n" }
                    elseif ($path -eq '/auth') { $status = 401; $body = 'login required' }
                    elseif ($path -eq '/bad') { $body = '{bad json' }
                    elseif ($path -eq '/fail') { $client.Close(); continue }
                    elseif ($path -eq '/stats') { $body = "$peak|$($counts['/auth'])|$($counts['/blocked'])" }
                    [void]$pending.Add(@{ Client = $client; At = [DateTime]::UtcNow.AddMilliseconds($delay); Status = $status; Body = $body; Extra = $extra })
                    $slowCount = @($pending | Where-Object { $_.At -gt [DateTime]::UtcNow }).Count
                    if ($slowCount -gt $peak) { $peak = $slowCount }
                }
                for ($i = $pending.Count - 1; $i -ge 0; $i--) {
                    $item = $pending[$i]
                    if ($item.At -gt [DateTime]::UtcNow) { continue }
                    try {
                        $stream = $item.Client.GetStream()
                        $bytes = [Text.Encoding]::UTF8.GetBytes($item.Body)
                        $headers = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $($item.Status) Fixture`r`n$($item.Extra)Content-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n")
                        $stream.Write($headers, 0, $headers.Length); $stream.Write($bytes, 0, $bytes.Length)
                    } catch {} finally { $item.Client.Close(); $pending.RemoveAt($i) }
                }
                Start-Sleep -Milliseconds 5
            }
        } finally {
            foreach ($item in $pending) { $item.Client.Close() }
            $listener.Stop()
        }
    }
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while (-not (Test-Path -LiteralPath $ready)) {
        if ([DateTime]::UtcNow -gt $deadline) { throw 'Fixture server did not start' }
        Start-Sleep -Milliseconds 50
    }

    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $macroPrefix = "'" + $workbook.Name + "'!"
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'modReviewFixesTest'
    # Reuse the complete fake invoice schemas from the existing regression.
    $baseline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-DetailInvoiceRuntime.ps1') -Raw -Encoding UTF8
    $detail = [regex]::Match($baseline, '(?s)Private Function RunDetailInvoiceFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    $summary = [regex]::Match($baseline, '(?s)Private Function RunSummaryDateFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    if (-not $detail -or -not $summary) { throw 'Missing regression fixtures' }
    $fixtureFunctions = "Public Function ReviewDetailJson() As String`r`n    Dim sampleJson As String, sellerTaxId As String, buyerTaxId As String`r`n    sellerTaxId = `"0000000001`": buyerTaxId = `"0000000002`"`r`n$detail`r`n    ReviewDetailJson = sampleJson`r`nEnd Function`r`n" +
        "Public Function ReviewSummaryJson() As String`r`n    Dim sampleJson As String`r`n$summary`r`n    ReviewSummaryJson = sampleJson`r`nEnd Function`r`n"
    $module.CodeModule.AddFromString(@'
Option Explicit
Public ReviewBaseUrl As String
Private stage As String
Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    stage = label
    If Not condition Then Err.Raise 5, , label
End Sub

Public Function ReviewCore() As String
    On Error GoTo Failed
    Dim value As Object, response As String, r As Long, seq As Long, mode As Long
    Dim ws As Worksheet, invoice As Object, items As Collection, item As Variant
    Call LinkTraCuu
    Call tenCotTraCuu
    arrTrangThai = Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = Sheets("LinkTraCuu").Range("O2:O11").Value
    For mode = 0 To 2
        Set value = JsonConverter.ParseJSON(ReviewSummaryJson())
        If mode = 1 Then
            value("datas")(1)("thttltsuat") = Null
            value("datas")(1)("thttlphi") = Null
            value("datas")(1)("cttkhac") = Null
            value("datas")(1)("ttkhac") = Null
        ElseIf mode = 2 Then
            value("datas")(1).Remove "thttltsuat"
            value("datas")(1).Remove "thttlphi"
            value("datas")(1).Remove "cttkhac"
            value("datas")(1).Remove "ttkhac"
        End If
        r = 300 + mode: seq = 1
        ghiExcel_TongHop JsonConverter.ConvertToJson(value), r, 1, seq
        Check r = 301 + mode, "summary optional arrays " & mode
    Next mode
    Set ws = Sheets("ChiTietHD_Mua")
    For mode = 0 To 2
        Set invoice = JsonConverter.ParseJSON(ReviewDetailJson())
        Set items = invoice("hdhhdvu")
        invoice("cttkhac") = Null: invoice("ttkhac") = Null
        If mode = 1 Then
            For Each item In items
                item("tthue") = Null: item("thtcthue") = Null: item("ttkhac") = Null
            Next item
        ElseIf mode = 2 Then
            For Each item In items
                item("tthue") = 0: item("thtcthue") = 200
            Next item
            invoice("tgtthue") = 0
        End If
        r = 310 + mode * 2
        ghiExcel_ChiTiet JsonConverter.ConvertToJson(invoice), r, "mua"
        Check r = 312 + mode * 2, "detail cursor " & mode
        Check ws.Cells(r - 2, 31).Value2 = "Tien thue ok", "detail tax sum " & mode
        Check ws.Cells(r - 2, 28).Value2 = IIf(mode = 2, 0, 20), "detail missing vs explicit zero tax " & mode
    Next mode
    r = 320
    On Error Resume Next
    ghiExcel_ChiTiet "{bad json", r, "mua"
    mode = Err.Number
    Err.Clear
    On Error GoTo Failed
    Check mode <> 0 And r = 320, "malformed detail must raise without advancing"
    Set value = JsonConverter.ParseJSON(BuildGdtLoginJson("fixture", "quote" & Chr$(34) & "slash\", "123", "key"))
    Check value("password") = "quote" & Chr$(34) & "slash\", "login JSON escaping"
    Check Not GdtJsonHasError(JsonConverter.ParseJSON("{""gchu"":""error-free fixture""}")), "substring error is not an API error"
    Check GdtJsonHasError(JsonConverter.ParseJSON("{""error"":""failed""}")), "API error object"
    Check GdtInvoiceFileBase("0000000001", "1", "TEST", "1") <> GdtInvoiceFileBase("0000000002", "1", "TEST", "1"), "seller file collision"
    Check GdtInvoiceFileBase("a_b", "1", "TEST", "1") <> GdtInvoiceFileBase("a~5Fb", "1", "TEST", "1"), "escaped filename collision"
    Set ws = EnsureGdtErrorReportSheet()
    ws.Rows("2:100").ClearContents
    GdtReportAccount = "account-a"
    UpsertGdtErrorReport "mua", "query", "", "", "", "", Empty, "LIST", "period-1", 500, "fixture", 1, 0, "Failed"
    UpsertGdtErrorReport "mua", "query", "", "", "", "", Empty, "LIST", "period-2", 500, "fixture", 1, 0, "Failed"
    GdtReportAccount = "account-b"
    UpsertGdtErrorReport "mua", "query", "", "", "", "", Empty, "LIST", "period-1", 500, "fixture", 1, 0, "Failed"
    Check ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 4, "three distinct LIST tasks"
    GdtReportAccount = "account-a"
    MarkGdtRetrySuccess "mua", "query", "", "", "", "", "LIST", 1, "", "period-1"
    Check ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 3, "clear only one LIST task"
    Check ws.Cells(2, 11).Value2 = "period-2" And ws.Cells(3, 18).Value2 = "account-b", "preserve other period and account"
    GdtReportAccount = ""
    ReviewCore = "PASS"
    Exit Function
Failed:
    ReviewCore = stage & ": " & Err.Number & " " & Err.Description
End Function

Public Function ReviewThrottle() As String
    On Error GoTo Failed
    Dim throttle As New clsGdtXmlThrottle
    Call ResetGdtOperationControl
    Call ResetGdtRequestSession
    throttle.Configure 4, 800
    Check throttle.CanStart(0, 100), "initial slot"
    throttle.Started 100
    Check Not throttle.CanStart(0, 100.4), "request spacing"
    Check Not throttle.CanStart(4, 101), "concurrency cap"
    throttle.RateLimited 20, 101
    Check throttle.CurrentConcurrency = 3 And throttle.CooldownUntil = 121, "429 cooldown and reduction"
    throttle.RateLimited 1, 102
    Check throttle.CurrentConcurrency = 3 And throttle.CooldownUntil = 121, "shared cooldown never shortened and drop once"
    Check Not throttle.CanStart(0, 120), "no dispatch during cooldown"
    throttle.Recover 122
    Check throttle.CurrentIntervalMs < 1800 And throttle.CurrentConcurrency = 3, "recover spacing first"
    Dim n As Long
    For n = 1 To 20
        throttle.Recover 122 + n * 10
    Next n
    Check throttle.CurrentIntervalMs = 800 And throttle.CurrentConcurrency = 4, "full recovery"
    SetGdtPaused True
    Check Not throttle.CanStart(0, 500), "pause dispatch"
    SetGdtPaused False
    RequestGdtStop
    Check Not throttle.CanStart(0, 500), "stop dispatch"
    ResetGdtOperationControl
    ReviewThrottle = "PASS"
    Exit Function
Failed:
    ReviewThrottle = stage & ": " & Err.Description
End Function

Public Function ReviewXml(ByVal folder As String, ByVal zipPath As String, ByVal destination As String) As String
    On Error GoTo Failed
    Dim ws As Worksheet, r As Long
    Set ws = Sheets("ChiTietHD_Mua_XML")
    ws.Range("A3:AI100").ClearContents
    Sheets("LinkTraCuu").Range("O14").Value = "tax"
    parseXML folder, True, False
    Check ws.Cells(3, 26).Value2 = 20 And ws.Cells(4, 26).Value2 = 30, "line-relative XML metadata"
    For r = 5 To 6
        Check ws.Cells(r, 26).Value2 = 10 And ws.Cells(r, 27).Value2 = 110, "XML fallback once " & r
    Next r
    Check ws.Cells(7, 26).Value2 = 0 And ws.Cells(7, 27).Value2 = 100, "explicit XML zero tax"
    Check ws.Cells(8, 26).Value2 = 0 And ws.Cells(8, 27).Value2 = 0, "no tax carried to next XML line"
    Unzip zipPath, destination
    Unzip zipPath, destination
    ReviewXml = "PASS"
    Exit Function
Failed:
    ReviewXml = stage & ": " & Err.Description
End Function

Public Function RunHttpFixture() As String
    On Error GoTo Failed
    Dim throttle As New clsGdtXmlThrottle, tasks As New Collection
    Dim task As clsGdtXmlTask, r As Long, active As Long, completed As Long, start As Double
    Call ResetGdtOperationControl
    Call ResetGdtRequestSession
    throttle.Configure 4, 50
    For r = 1 To 4
        Set task = New clsGdtXmlTask
        task.Configure r, ReviewBaseUrl & "/slow/" & r, "", 1
        tasks.Add task
    Next r
    start = GdtClockSeconds()
    Do While completed < 4
        active = 0: completed = 0
        For Each task In tasks
            If task.Running Then active = active + 1
        Next task
        For Each task In tasks
            If task.Running Then active = active - 1
            task.Tick throttle, active
                    If task.Running Then active = active + 1
            If task.Finished Then
                Check task.Result.Success, "parallel request status"
                Check task.Result.ResponseText = "PK-fixture", "parallel response text retained"
                Check UBound(task.Result.ResponseBody) >= 1, "parallel binary response retained"
                completed = completed + 1
            End If
        Next task
        If GdtClockSeconds() - start > 8 Then Err.Raise 5, , "parallel request timeout"
        GdtPumpWait
    Loop
    Set task = New clsGdtXmlTask
    throttle.Configure 4, 50
    task.Configure 0, ReviewBaseUrl & "/rate", "", 2
    start = GdtClockSeconds()
    Do Until task.Finished
        task.Tick throttle, IIf(task.Running, 0, 0)
        If GdtClockSeconds() - start > 10 Then Err.Raise 5, , "429 request timeout"
        GdtPumpWait
    Loop
    Check task.Result.Success And task.Result.Attempts = 2, "429 retry succeeds"
    Check Len(task.Result.ErrorMessage) = 0, "successful retry clears transient error"
    Check GdtClockSeconds() - start >= 1 And throttle.CurrentConcurrency = 3, "429 respected by real requests"
    Set task = New clsGdtXmlTask
    throttle.Configure 4, 0
    task.Configure 0, ReviewBaseUrl & "/slow/pause", "", 1
    SetGdtPaused True
    task.Tick throttle, 0
    Check task.Result.Attempts = 0, "pause before launch"
    SetGdtPaused False
    task.Tick throttle, 0
    SetGdtPaused True
    start = GdtClockSeconds()
    Do Until task.Finished
        task.Tick throttle, 0
        If GdtClockSeconds() - start > 5 Then Err.Raise 5, , "pause drain timeout"
        GdtPumpWait
    Loop
    Check task.Result.Success, "drain in-flight request while paused"
    SetGdtPaused False
    Set task = New clsGdtXmlTask
    task.Configure 0, ReviewBaseUrl & "/slow/stop", "", 1
    task.Tick throttle, 0
    RequestGdtStop
    task.Tick throttle, 0
    Check task.Finished And Not task.Result.Success And Not task.Running, "stop aborts active XML"
    ResetGdtOperationControl
    Set task = New clsGdtXmlTask
    task.Configure 0, ReviewBaseUrl & "/auth", "", 1
    start = GdtClockSeconds()
    Do Until task.Finished
        task.Tick throttle, 0
        If GdtClockSeconds() - start > 5 Then Err.Raise 5, , "auth timeout"
        GdtPumpWait
    Loop
    Check GdtAuthenticationFailed And task.Result.AuthenticationFailure, "401 stops session"
    Dim result As clsGdtRequestResult
    Set result = ExecuteGdtRequest("GET", ReviewBaseUrl & "/blocked", "", , , , , 1)
    Check result.Attempts = 0 And result.AuthenticationFailure, "central auth guard before send"
    Set task = New clsGdtXmlTask
    task.Configure 0, ReviewBaseUrl & "/blocked", "", 1
    task.Tick throttle, 0
    Check task.Result.Attempts = 0 And task.Finished, "async auth guard before send"
    ResetGdtRequestSession
    Set task = New clsGdtXmlTask
    task.Configure 0, ReviewBaseUrl & "/fail", "", 2, 1
    start = GdtClockSeconds()
    Do Until task.Finished
        task.Tick throttle, 0
        If GdtClockSeconds() - start > 10 Then Err.Raise 5, , "transport retry timeout"
        GdtPumpWait
    Loop
    Check Not task.Result.Success And task.Result.Attempts = 2 And task.Result.ShouldQueueFinalRetry, "transport failures exhausted"
    RunHttpFixture = "PASS"
    Exit Function
Failed:
    Dim failure As String
    failure = Err.Description
    If Not task Is Nothing Then task.Cancel
    Call ResetGdtOperationControl
    Call ResetGdtRequestSession
    RunHttpFixture = stage & ": " & failure
End Function



Public Function ReviewTaskSmoke() As String
    Dim task As New clsGdtXmlTask
    Dim throttle As New clsGdtXmlThrottle
    task.Configure 0, "http://127.0.0.1:1/fixture", "", 1
    throttle.Configure 4, 0
    RequestGdtStop
    task.Tick throttle, 0
    ResetGdtOperationControl
    ReviewTaskSmoke = "PASS"
End Function
'@)
    $formCode = $workbook.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $formCode.AddFromString(@'
Public Function ReviewXmlSave(ByVal zipPath As String, ByVal destination As String) As String
    On Error GoTo Failed
    Dim stream As Object, body As Variant, ws As Worksheet, originalFolder As Variant
    originalFolder = oSaveUnzipFolder
    oSaveUnzipFolder = destination
    Me.txtXMLFolderPath.Value = destination
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.LoadFromFile zipPath
    body = stream.Read: stream.Close
    If Not TrySaveXml(body, "query", "/fixture", "0000000001", "1", "FIXTURE", "42", Empty) Then Err.Raise 5, , "save first seller"
    If Not TrySaveXml(body, "query", "/fixture", "0000000001", "1", "FIXTURE", "42", Empty) Then Err.Raise 5, , "repeat same invoice"
    If Not TrySaveXml(body, "sco-query", "/fixture", "0000000002", "1", "FIXTURE", "42", Empty) Then Err.Raise 5, , "save second seller"
    Set ws = EnsureGdtErrorReportSheet()
    ws.Rows("2:100").ClearContents
    Dim invalidBody(0 To 2) As Byte
    invalidBody(0) = 80: invalidBody(1) = 75: invalidBody(2) = 0
    If TrySaveXml(invalidBody, "query", "/fixture", "0000000001", "1", "FIXTURE", "42", Empty) Then Err.Raise 5, , "invalid replacement ZIP reported success"
    If FileLen(destination & "\0000000001_1_FIXTURE_42.zip") <> UBound(body) - LBound(body) + 1 Then Err.Raise 5, , "invalid replacement destroyed existing ZIP"
    ws.Rows("2:100").ClearContents
    If TrySaveXml(invalidBody, "query", "/fixture", "0000000003", "1", "FIXTURE", "42", Empty) Then Err.Raise 5, , "invalid ZIP reported success"
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row <> 2 Then Err.Raise 5, , "save failure missing error report"
    oSaveUnzipFolder = originalFolder
    ReviewXmlSave = "PASS"
    Exit Function
Failed:
    oSaveUnzipFolder = originalFolder
    ReviewXmlSave = Err.Description
End Function
'@)
    $module.CodeModule.AddFromString($fixtureFunctions)
    $formCode.AddFromString(@'
Public Function ReviewForm(ByVal badEndpoint As String) As String
    On Error GoTo Failed
    Dim parsed As Object, originalEnabled As Boolean, closeCancel As Integer
    Dim queued As clsGdtRetryItem, row As Long, ws As Worksheet
    originalEnabled = Me.txtXMLFolderPath.Enabled
    cToken = "fixture-session"
    SetDownloadControlState True
    If Me.cboDonVi.Enabled Or Me.optMua.Enabled Or Me.cmdTaiHoaDon.Enabled Then Err.Raise 5, , "mutable download controls"
    cToken = "changed-session"
    If RunToken() <> "fixture-session" Then Err.Raise 5, , "token snapshot"
    cmdTaiHoaDon_Click
    If Not mDownloadRunning Then Err.Raise 5, , "reentrant download"
    UserForm_QueryClose closeCancel, 0
    If closeCancel <> 1 Or Not GdtStopRequested Then Err.Raise 5, , "close while downloading"
    SetDownloadControlState False
    If Me.txtXMLFolderPath.Enabled <> originalEnabled Or Not Me.cmdTaiHoaDon.Enabled Then Err.Raise 5, , "restore enabled states"
    Call ResetGdtOperationControl
    Call ResetGdtRequestSession
    If Not TryParseGdtJson("{""datas"":[],""gchu"":""error-free fixture""}", parsed, "query", "/purchase?") Then Err.Raise 5, , "substring error"
    If TryParseGdtJson("{""message"":""bad schema""}", parsed, "query", "/purchase?") Then Err.Raise 5, , "list schema"
    Set mFinalRetryQueue = New Collection
    mFinalRetryCooldownDone = True
    Set queued = New clsGdtRetryItem
    queued.Direction = CurrentDirectionName()
    queued.ApiSource = "query": queued.SellerTaxCode = "0000000001"
    queued.TemplateCode = "1": queued.InvoiceSeries = "FIXTURE": queued.InvoiceNumber = "1"
    queued.Stage = UniConvert("Laasy chi tieest"): queued.ResponseKind = "DETAIL"
    queued.Endpoint = badEndpoint
    mFinalRetryQueue.Add queued
    UpsertGdtErrorReport queued.Direction, "query", queued.SellerTaxCode, "1", "FIXTURE", "1", Empty, queued.Stage, badEndpoint, 500, "fixture", 1, 0, "Failed"
    row = 330
    ProcessQueuedDetailRetries row
    If row <> 330 Then Err.Raise 5, , "malformed retry advanced row"
    Set ws = EnsureGdtErrorReportSheet()
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row < 2 Then Err.Raise 5, , "malformed retry cleared errors"
    ReviewForm = "PASS"
    Exit Function
Failed:
    ReviewForm = Err.Description
    SetDownloadControlState False
    Call ResetGdtOperationControl
    Call ResetGdtRequestSession
End Function
'@)
    $module.CodeModule.AddFromString("Public Function ReviewXmlSaveEntry() As String`r`n Dim form As New frmTaiHoaDon`r`n ReviewXmlSaveEntry = form.ReviewXmlSave(`"$zipPath`", `"$outDir`")`r`n Unload form`r`nEnd Function`r`n")
    $module.CodeModule.AddFromString(@'
Public Function ReviewRunForm(ByVal endpoint As String) As String
    Dim form As New frmTaiHoaDon
    ReviewRunForm = form.ReviewForm(endpoint)
    Unload form
End Function
'@)
    # Use no-argument COM entrypoints. VBA passes the fixture paths internally;
    # this avoids Windows PowerShell 5.1's Excel.Run optional-argument binder.
    $entrypoints = "Public Function ReviewXmlEntry() As String`r`n ReviewXmlEntry = ReviewXml(`"$xmlDir`", `"$zipPath`", `"$outDir`")`r`nEnd Function`r`n" +
        "Public Function RunHttpFixtureEntry() As String`r`n ReviewBaseUrl = `"http://127.0.0.1:$port`": RunHttpFixtureEntry = RunHttpFixture()`r`nEnd Function`r`n" +
        "Public Function ReviewFormEntry() As String`r`n ReviewFormEntry = ReviewRunForm(`"http://127.0.0.1:$port/bad`")`r`nEnd Function`r`n"
    $module.CodeModule.AddFromString($entrypoints)
    Write-Output 'Compiling fixture VBA...'
    $excel.VBE.ActiveVBProject = $workbook.VBProject
    $excel.VBE.CommandBars.FindControl(1, 578).Execute()
    Write-Output 'Fixture VBA compiled'
    Assert-Result 'AsyncHttpPauseStopAuthRetry' ($excel.Run("${macroPrefix}RunHttpFixtureEntry"))
    Assert-Result 'ParserTaxLoginFilesReport' ($excel.Run("${macroPrefix}ReviewCore"))
    Assert-Result 'ThrottlePolicy' ($excel.Run("${macroPrefix}ReviewThrottle"))
    Assert-Result 'XmlParseAndExtraction' ($excel.Run("${macroPrefix}ReviewXmlEntry"))
    if (@(Get-ChildItem -LiteralPath $outDir -Filter '*.xml').Count -ne 2 -or
        -not (Test-Path -LiteralPath (Join-Path $outDir 'fixture.html')) -or
        -not (Test-Path -LiteralPath (Join-Path $unrelated 'keep.txt')) -or
        (Test-Path -LiteralPath (Join-Path $fixtureRoot 'invoice.xml')) -or
        @(Get-ChildItem -LiteralPath $outDir -Directory).Count -ne 0) { throw 'ZIP isolation/cleanup failed' }
    $stats = (Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$port/stats").Content
    if ($stats -is [byte[]]) { $stats = [Text.Encoding]::UTF8.GetString($stats) }
    if ($stats -ne '4|1|') { throw "Expected peak 4, one auth request, no blocked requests: $stats" }
    $results['ServerPeakConcurrency'] = 4
    Assert-Result 'FormFreezeSchemaMalformedRetry' ($excel.Run("${macroPrefix}ReviewFormEntry"))
    Assert-Result 'XmlSaveRepeatCollisionAndFailure' ($excel.Run("${macroPrefix}ReviewXmlSaveEntry"))
    foreach ($seller in @('0000000001', '0000000002')) {
        $fileBase = "${seller}_1_FIXTURE_42"
        foreach ($extension in @('.zip', '.xml', '_2.xml', '.html')) {
            if (-not (Test-Path -LiteralPath (Join-Path $outDir ($fileBase + $extension)))) {
                throw "Missing saved invoice fixture: $fileBase$extension"
            }
        }
    }
    if ((Test-Path -LiteralPath (Join-Path $outDir '0000000003_1_FIXTURE_42.xml')) -or
        @(Get-ChildItem -LiteralPath $outDir -Directory).Count -ne 0) { throw 'Failed archive created XML or left staging folders' }
    $report = Join-Path (Split-Path -Parent $BuiltWorkbook) 'review-fixes-runtime-result.json'
    $results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Output "REVIEW-FIXES RUNTIME TEST PASSED: $report"
} catch {
    Write-Output $_.InvocationInfo.PositionMessage
    throw
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $module; Release-ComObject $formCode; Release-ComObject $workbook; Release-ComObject $excel
    if ($null -ne $server) { Stop-Job $server; Remove-Job $server -Force }
    # This path is a freshly created GUID child of the OS temporary directory.
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -match '^hddt-fixes-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
