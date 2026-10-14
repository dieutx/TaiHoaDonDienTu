[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$excel = $null; $book = $null; $module = $null; $code = $null; $formCode = $null; $taskCode = $null; $server = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'modDownloaderParityTests'
    $code = $module.CodeModule
    $code.AddFromString(@'
Option Explicit
Public Function DownloaderThrottleFixture() As String
    On Error GoTo Failed
    Dim throttle As clsGdtXmlThrottle, base As Variant, step As Long
    For Each base In Array(0, 300, 600, 5000, 5001, 20000, 60000)
        Set throttle = New clsGdtXmlThrottle
        throttle.Configure 4, CLng(base)
        throttle.Started 100
        For step = 1 To 3
            throttle.RateLimited 1, 100 + step
            If throttle.CurrentIntervalMs < CLng(base) Then Err.Raise 5, , "429 shortened configured interval " & base
        Next step
        For step = 1 To 40
            throttle.Recover 200 + step * 11
        Next step
        If throttle.CurrentIntervalMs <> CLng(base) Or throttle.CurrentConcurrency <> 4 Then Err.Raise 5, , "Recovery did not return to base"
    Next base
    DownloaderThrottleFixture = "PASS"
    Exit Function
Failed:
    DownloaderThrottleFixture = "FAIL: " & Err.Description
End Function
Public Function DownloaderIdentityFixture() As String
    On Error GoTo Failed
    If GdtInvoiceIdentity("seller-a", "1", "same-series", "1") = GdtInvoiceIdentity("seller-b", "1", "same-series", "1") Then Err.Raise 5, , "Seller identities collided"
    If GdtInvoiceIdentity("seller-a", "1", "same-series", "1") = GdtInvoiceIdentity("seller-a", "2", "same-series", "1") Then Err.Raise 5, , "Template identities collided"
    ResetGdtOperationControl
    GdtStopRequested = True
    Dim started As Double
    started = GdtClockSeconds()
    WaitGdtSeconds 600
    If GdtClockSeconds() - started > 1 Then Err.Raise 5, , "Retry ignored safe stop"
    ResetGdtOperationControl
    DownloaderIdentityFixture = "PASS"
    Exit Function
Failed:
    ResetGdtOperationControl
    DownloaderIdentityFixture = "FAIL: " & Err.Description
End Function
'@)
    foreach ($name in @('DownloaderThrottleFixture', 'DownloaderIdentityFixture')) {
        $result = [string]$excel.Run("${prefix}$name")
        Write-Output "$name=$result"
        if ($result -ne 'PASS') { throw $result }
    }
    $baseline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-DetailInvoiceRuntime.ps1') -Raw -Encoding UTF8
    Write-Output 'Preparing synthetic detail JSON...'
    $fragment = [regex]::Match($baseline, '(?s)Private Function RunDetailInvoiceFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    if (-not $fragment) { throw 'Missing synthetic detail fixture.' }
    $code.AddFromString("Public Function DownloaderDetailJson() As String`r`n Dim sampleJson As String, sellerTaxId As String, buyerTaxId As String`r`n sellerTaxId = `"0000000001`": buyerTaxId = `"0000000002`"`r`n$fragment`r`n DownloaderDetailJson = sampleJson`r`nEnd Function`r`n")
    $template = [string]$excel.Run("${prefix}DownloaderDetailJson")
    Write-Output 'Detail JSON ready.'
    $json = $template | ConvertFrom-Json
    $json.nbmst = '0000000001'; $json.khmshdon = '1'; $json.khhdon = 'C26TEST'; $json.shdon = 42
    $json.nmmst = '0000000002'
    $template = $json | ConvertTo-Json -Depth 20 -Compress
    $probe = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
    $probe.Start(); $port = ([Net.IPEndPoint]$probe.LocalEndpoint).Port; $probe.Stop()
    $server = Start-Job -ArgumentList $port,$template -ScriptBlock {
        param($port, $template)
        $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $port)
        $listener.Start()
        $hits = @{}
        try {
            while ($true) {
                if (-not $listener.Pending()) { Start-Sleep -Milliseconds 5; continue }
                $client = $listener.AcceptTcpClient()
                $client.ReceiveTimeout = 2000
                try {
                    $reader = New-Object IO.StreamReader($client.GetStream())
                    $line = $reader.ReadLine()
                    $path = ($line -split ' ')[1]
                    $headers = @{}
                    do {
                        $line = $reader.ReadLine()
                        if ($line -match '^([^:]+):\s*(.*)$') { $headers[$Matches[1]] = $Matches[2] }
                    } while ($null -ne $line -and $line.Length -gt 0)
                    $status = 200; $body = $template
                    if ($path -eq '/ready') { $body = 'ready' }
                    elseif ($path -eq '/stats') { $body = $hits | ConvertTo-Json -Compress }
                    else {
                        if ($path -notmatch '^/([^/]+)/api/query/invoices/(detail|export-xml)\?') { throw ('Unexpected endpoint: ' + $path) }
                        $mode = $Matches[1]; $kind = $Matches[2]
                        if ($kind -eq 'export-xml') { $mode += '-xml' }
                        $hits[$mode] = [int]$hits[$mode] + 1
                        if ($kind -eq 'export-xml') { $status = 504; $body = 'Synthetic XML gateway timeout' }
                        elseif ($headers['Authorization'] -ne 'Bearer synthetic-session' -or -not $headers['Request-Id'] -or
                            $headers['End-Point'] -ne '/tra-cuu/tra-cuu-hoa-don' -or -not $headers['Action']) {
                            $status = 400; $body = 'Invalid fixture headers'
                        }
                        elseif ($mode -eq 'invalid') { $body = '{bad JSON' }
                        elseif ($mode -eq 'mismatch') { $body = $template.Replace('C26TEST', 'WRONG-SERIES') }
                        elseif ($mode -eq 'failure') { $status = 400; $body = 'Synthetic detail error' }
                        elseif ($mode -eq 'auth') { $status = 401; $body = 'Synthetic auth error' }
                    }
                    $bytes = [Text.Encoding]::UTF8.GetBytes($body)
                    $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $status Fixture`r`nContent-Type: application/json; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n")
                    $stream = $client.GetStream()
                    $stream.Write($head, 0, $head.Length); $stream.Write($bytes, 0, $bytes.Length)
                }
                catch { Write-Output $_.Exception.Message }
                finally { $client.Close() }
            }
        }
        finally { $listener.Stop() }
    }
    $deadline = [datetime]::UtcNow.AddSeconds(10)
    while ($true) {
        try { $null = Invoke-WebRequest "http://127.0.0.1:$port/ready" -UseBasicParsing -TimeoutSec 1; break }
        catch { if ([datetime]::UtcNow -gt $deadline) { throw 'Local fixture server did not start.' }; Start-Sleep -Milliseconds 50 }
    }
    $formCode = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    Write-Output 'Local HTTP ready; instrumenting form in memory...'
    $text = $formCode.Lines(1, $formCode.CountOfLines)
    $text = $text.Replace('https://hoadondientu.gdt.gov.vn/api/', 'http://127.0.0.1:' + $port + '/" & mDownloaderFixtureMode & "/api/')
    $formCode.DeleteLines(1, $formCode.CountOfLines)
    $formCode.AddFromString($text)
    $formCode.AddFromString(@'
Private mDownloaderFixtureMode As String
Public Function DownloaderRecoveryFixture(ByVal mode As String, ByVal purchase As Boolean, ByVal status As Long, ByVal source As Long) As String
    On Error GoTo Failed
    EndGdtResume
    ResetGdtOperationControl
    ResetGdtRequestSession
    EnsureRuntimeControls
    mRunPurchase = purchase: mRunToken = "synthetic-session": mDownloadRunning = True
    mDownloaderFixtureMode = mode
    Me.txtSleep.Value = "1"
    Call LinkTraCuu: Call tenCotTraCuu
    arrTrangThai = ThisWorkbook.Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = ThisWorkbook.Sheets("LinkTraCuu").Range("O2:O11").Value
    Set mFinalRetryQueue = New Collection
    Dim ws As Worksheet, report As Worksheet, xmlWs As Worksheet, task As clsGdtXmlTask, result As clsGdtRequestResult
    Dim rows As Long, lastReport As Long
    Set ws = ThisWorkbook.Worksheets(IIf(purchase, "ChiTietHD_Mua", "ChiTietHD_Ban"))
    Set xmlWs = ThisWorkbook.Worksheets(IIf(purchase, "ChiTietHD_Mua_XML", "ChiTietHD_Ban_XML"))
    Set report = EnsureGdtErrorReportSheet()
    ws.Range("A3:AJ200").ClearContents: xmlWs.Range("A3:AF200").ClearContents
    report.Range("A2:S200").ClearContents
    GdtReportAccount = "SYNTHETIC-ACCOUNT": GdtReportSession = "SYNTHETIC-SESSION"
    ReDim arrHDChiTiet(1, 7)
    arrHDChiTiet(0, 0) = "0000000001": arrHDChiTiet(0, 1) = "C26TEST"
    arrHDChiTiet(0, 2) = "42": arrHDChiTiet(0, 3) = "1"
    arrHDChiTiet(0, 4) = source: arrHDChiTiet(0, 5) = DateSerial(2026, 1, 1)
    arrHDChiTiet(0, 6) = 3: arrHDChiTiet(0, 7) = 1
    If mode = "reuse" Then
        Dim summaryWs As Worksheet, page As Object, invoice As Object, invoices As Collection
        Dim summaryRow As Long, seq As Long
        Set summaryWs = ThisWorkbook.Worksheets(IIf(purchase, "TongHopHD_Mua", "TongHopHD_Ban"))
        summaryWs.Range("A3:BN200").ClearContents
        BeginGdtResume "SYNTHETIC-ACCOUNT", IIf(purchase, 1, 2), True
        Set invoice = JsonConverter.ParseJSON(DownloaderDetailJson())
        invoice("nbmst") = "0000000001": invoice("khmshdon") = "1"
        invoice("khhdon") = "C26TEST": invoice("shdon") = 42
        Set page = CreateObject("Scripting.Dictionary"): Set invoices = New Collection
        invoices.Add invoice: page.Add "datas", invoices
        summaryRow = 3: seq = 1
        If Not WriteSummaryPage(JsonConverter.ConvertToJson(page), summaryRow, IIf(purchase, 1, 2), seq, _
            "query", "synthetic-list", DateSerial(2026, 1, 1), page) Then Err.Raise 5, , "Resume setup failed"
    End If
    Set task = New clsGdtXmlTask
    task.Configure 0, "http://example.invalid/export-xml", "synthetic-session"
    Set result = task.Result
    result.StatusCode = status: result.Attempts = 5
    result.ErrorMessage = "Synthetic XML failure": result.ShouldQueueFinalRetry = True
    If mode = "stop" Then GdtStopRequested = True
    If mode = "scheduler" Then
        taiXML_zip 1
    Else
        HandleXmlResult task
    End If
    If mode = "reuse" Then HandleXmlResult task
    rows = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row - 2
    lastReport = report.Cells(report.Rows.Count, 1).End(xlUp).Row
    If mode = "purchase" Or mode = "sold" Or mode = "sco" Or mode = "no-xml" Or mode = "reuse" Or mode = "scheduler" Then
        If rows <> 2 Then Err.Raise 5, , "Fallback did not write expected goods"
        If CStr(ws.Cells(3, 8).Value2) <> "0000000001" Then Err.Raise 5, , "Seller lost leading zero"
        If CStr(report.Cells(lastReport, 16).Value2) <> "Recovered from API detail (no XML)" Then Err.Raise 5, , "Recovery result missing"
        If report.Cells(lastReport, 12).Value2 <> status Or report.Cells(lastReport, 14).Value2 <> IIf(mode = "scheduler", 1, 5) Then Err.Raise 5, , "Original XML metadata lost"
        If mFinalRetryQueue.Count <> 0 Then Err.Raise 5, , "Recovered data queued another XML request"
    Else
        If rows <> 0 Then Err.Raise 5, , "Failed/stopped recovery wrote goods"
        If mode <> "stop" And mode <> "other" Then
            If InStr(CStr(report.Cells(lastReport, 13).Value2), "fallback failed") = 0 Then Err.Raise 5, , "Fallback failure not reported"
        End If
    End If
    If Application.CountA(xmlWs.Range("A3:AF200")) <> 0 Then Err.Raise 5, , "API detail invented XML data"
    mDownloadRunning = False: mRunToken = ""
    EndGdtResume
    ResetGdtOperationControl
    DownloaderRecoveryFixture = "PASS"
    Exit Function
Failed:
    mDownloadRunning = False: mRunToken = ""
    EndGdtResume
    ResetGdtOperationControl
    DownloaderRecoveryFixture = "FAIL: " & Err.Description
End Function
'@)
    # Only this in-memory fixture lowers retries so the localhost 504 case finishes quickly.
    $taskCode = $book.VBProject.VBComponents.Item('clsGdtXmlTask').CodeModule
    $text = $taskCode.Lines(1, $taskCode.CountOfLines).Replace('mMaxAttempts = maxAttempts', 'mMaxAttempts = 1')
    $taskCode.DeleteLines(1, $taskCode.CountOfLines)
    $taskCode.AddFromString($text)
    $code.AddFromString(@'
Public Function DownloaderRunRecovery(ByVal mode As String, ByVal purchase As Boolean, ByVal status As Long, ByVal source As Long) As String
    DownloaderRunRecovery = frmTaiHoaDon.DownloaderRecoveryFixture(mode, purchase, status, source)
End Function
'@)
    foreach ($case in @(
        @('purchase', $true, 504, 1), @('sold', $false, 504, 1), @('sco', $true, 504, 2),
        @('no-xml', $true, 500, 1), @('reuse', $true, 504, 1), @('scheduler', $true, 504, 1), @('invalid', $true, 504, 1), @('mismatch', $true, 504, 1),
        @('failure', $true, 504, 1), @('auth', $true, 504, 1), @('stop', $true, 504, 1), @('other', $true, 502, 1)
    )) {
        Write-Output ('Running recovery case: ' + $case[0])
        $result = [string]$excel.Run("${prefix}DownloaderRunRecovery", [string]$case[0], [bool]$case[1], [int]$case[2], [int]$case[3])
        Write-Output "$($case[0])=$result"
        if ($result -ne 'PASS') { throw $result }
    }
    $stats = (Invoke-WebRequest "http://127.0.0.1:$port/stats" -UseBasicParsing).Content | ConvertFrom-Json
    foreach ($name in @('purchase','sold','sco','no-xml','reuse','scheduler','scheduler-xml','invalid','mismatch','failure','auth')) {
        if ($stats.$name -ne 1) { throw "Expected one detail request for $name." }
    }
    if ($null -ne $stats.PSObject.Properties['stop'] -or $null -ne $stats.PSObject.Properties['other']) { throw 'Stop/other status sent an unwanted detail request.' }
}
finally {
    if ($null -ne $server) { Stop-Job $server; Remove-Job $server -Force }
    Release-ComObject $taskCode; Release-ComObject $formCode; Release-ComObject $code; Release-ComObject $module
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Fixture modified the workbook artifact.' }
Write-Output 'DOWNLOADER PARITY RUNTIME TEST PASSED'
