[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$fixtureRoot = Join-Path $env:TEMP ('hddt-json-parallel-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$excel = $null; $workbook = $null; $module = $null; $formCode = $null; $server = $null
$results = [ordered]@{}
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'" + $workbook.Name + "'!"
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'modParallelJsonFixture'
    $baseline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-DetailInvoiceRuntime.ps1') -Raw -Encoding UTF8
    $detail = [regex]::Match($baseline, '(?s)Private Function RunDetailInvoiceFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    if (-not $detail) { throw 'Missing synthetic invoice schema.' }
    $module.CodeModule.AddFromString("Public Function ParallelDetailTemplate() As String`r`n Dim sampleJson As String, sellerTaxId As String, buyerTaxId As String`r`n sellerTaxId = `"0000000001`": buyerTaxId = `"0000000002`"`r`n$detail`r`n ParallelDetailTemplate = sampleJson`r`nEnd Function`r`n")
    $template = [string]$excel.Run("${prefix}ParallelDetailTemplate")
    if ($template.Length -lt 100 -or -not $template.StartsWith('{')) { throw 'Detail fixture returned an empty/invalid template.' }
    $probe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $probe.Start(); $port = ([Net.IPEndPoint]$probe.LocalEndpoint).Port; $probe.Stop()
    $ready = Join-Path $fixtureRoot 'ready'
    $server = Start-Job -ArgumentList $port,$ready,$template -ScriptBlock {
        param($port,$ready,$template)
        $ErrorActionPreference = 'Stop'
        $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $port)
        $listener.Start()
        [IO.File]::WriteAllText($ready, 'ready')
        $pending = New-Object Collections.ArrayList
        $counts = @{}; $peaks = @{}; $ids = @{}; $lastStart = @{}; $minIntervals = @{}
        $badHeaders = 0; $duplicateIds = 0; $badHeaderKinds = @{}; $referers = @{}
        try {
            while ($true) {
                while ($listener.Pending()) {
                    $client = $listener.AcceptTcpClient()
                    $client.ReceiveTimeout = 2000
                    $reader = New-Object IO.StreamReader($client.GetStream())
                    $request = $reader.ReadLine()
                    $headers = @{}
                    do {
                        $header = $reader.ReadLine()
                        if ($header -match '^([^:]+):\s*(.*)$') { $headers[$Matches[1]] = $Matches[2] }
                    } while ($null -ne $header -and $header.Length -gt 0)
                    $path = ($request -split ' ')[1]
                    if ($path -eq '/stats') {
                        $body = @{ Counts=$counts; Peaks=$peaks; MinIntervalsMs=$minIntervals; BadHeaders=$badHeaders; BadHeaderKinds=$badHeaderKinds; Referers=$referers; DuplicateIds=$duplicateIds } | ConvertTo-Json -Depth 6 -Compress
                        [void]$pending.Add(@{ Client=$client; At=[DateTime]::UtcNow; Status=200; Body=$body; Extra=''; Mode='stats' })
                        continue
                    }
                    if ($path -notmatch '^/([^/]+)/api/(query|sco-query)/invoices/(detail|relative|related)\?') { throw "Unexpected fixture endpoint: $path" }
                    $routeMatch = [regex]::Match($path, '^/([^/]+)/api/(query|sco-query)/invoices/(detail|relative|related)\?')
                    $mode = $routeMatch.Groups[1].Value; $kind = $routeMatch.Groups[3].Value
                    $started = [DateTime]::UtcNow
                    if ($lastStart.ContainsKey($mode)) {
                        $interval = ($started - $lastStart[$mode]).TotalMilliseconds
                        if (-not $minIntervals.ContainsKey($mode) -or $interval -lt $minIntervals[$mode]) { $minIntervals[$mode] = $interval }
                    }
                    $lastStart[$mode] = $started
                    $number = [int][regex]::Match($path, '[?&]shdon=(\d+)').Groups[1].Value
                    $key = "$mode/$kind/$number"
                    if (-not $counts.ContainsKey($key)) { $counts[$key] = 0 }
                    $counts[$key]++
                    $requestId = [string]$headers['Request-Id']
                    if ($ids.ContainsKey($requestId)) { $duplicateIds++ } else { $ids[$requestId] = $true }
                    if ($headers['Authorization'] -ne 'Bearer fixture-snapshot' -or $headers['Content-Type'] -ne 'application/json' -or
                        $headers['Accept'] -ne 'application/json, text/plain, */*' -or -not $requestId) { $badHeaders++ }
                    if ($kind -ne 'detail' -and (-not $headers['Action'] -or $headers['End-Point'] -ne '/tra-cuu/tra-cuu-hoa-don' -or
                        $headers['Accept-Language'] -ne 'vi')) { $badHeaders++ }
                    if ($kind -ne 'detail') { $referers[$mode] = [string]$headers['Referer'] }
                    $expectedHeaders = @{ Authorization='Bearer fixture-snapshot'; 'Content-Type'='application/json'; Accept='application/json, text/plain, */*' }
                    if ($kind -ne 'detail') { $expectedHeaders['End-Point']='/tra-cuu/tra-cuu-hoa-don'; $expectedHeaders['Accept-Language']='vi' }
                    foreach ($name in $expectedHeaders.Keys) {
                        if ($headers[$name] -ne $expectedHeaders[$name]) { $badHeaderKinds[$name] = [int]$badHeaderKinds[$name] + 1 }
                    }
                    $status = 200; $extra = ''; $delay = if ($number -eq 1) { 1800 } else { 800 }
                    $body = switch ($kind) {
                        'detail' { $template.Replace('"shdon":42', '"shdon":' + $number) }
                        'relative' { '[{"khmshdon":"1","khhdon":"REL-' + $number + '","shdon":' + (9000 + $number) + '}]' }
                        'related' { '{"notice":"NOTICE-' + $number + '"}' }
                    }
                    if ($mode -eq 'mixed' -and $number -eq 2 -and $counts[$key] -eq 1) { $status=429; $body='rate'; $extra="Retry-After: 1`r`n"; $delay=100 }
                    if ($mode -eq 'mixed' -and $number -eq 3 -and $counts[$key] -eq 1) { $status=504; $body='gateway timeout'; $delay=100 }
                    if ($mode -eq 'mixed' -and $number -eq 4 -and $counts[$key] -eq 1) { $status=503; $body='service unavailable'; $delay=100 }
                    if ($mode -eq 'failure' -and $number -eq 2 -and $counts[$key] -le 2) { $status=500; $body='fixture failure'; $delay=100 }
                    if ($mode -eq 'invalid' -and $number -eq 2) { $body='{bad json'; $delay=100 }
                    if ($mode -eq 'auth' -and $number -eq 1) { $status=403; $body='login required'; $delay=150 }
                    if ($mode -eq 'buffer') { $delay = if ($number -eq 1) { 2200 } else { 100 } }
                    [void]$pending.Add(@{ Client=$client; At=[DateTime]::UtcNow.AddMilliseconds($delay); Status=$status; Body=$body; Extra=$extra; Mode=$mode })
                    $active = @($pending | Where-Object { $_.Mode -eq $mode -and $_.At -gt [DateTime]::UtcNow }).Count
                    if (-not $peaks.ContainsKey($mode) -or $active -gt $peaks[$mode]) { $peaks[$mode] = $active }
                }
                for ($i=$pending.Count-1; $i -ge 0; $i--) {
                    $item = $pending[$i]
                    if ($item.At -gt [DateTime]::UtcNow) { continue }
                    try {
                        $bytes = [Text.Encoding]::UTF8.GetBytes($item.Body)
                        $responseHeaders = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $($item.Status) Fixture`r`n$($item.Extra)Content-Type: application/json`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n")
                        $stream = $item.Client.GetStream()
                        $stream.Write($responseHeaders,0,$responseHeaders.Length); $stream.Write($bytes,0,$bytes.Length)
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
        if ([DateTime]::UtcNow -gt $deadline) { throw 'Fixture server did not start.' }
        Start-Sleep -Milliseconds 50
    }
    $module.CodeModule.AddFromString(@'
Option Explicit
Public ParallelBaseUrl As String
Public ParallelMode As String
Public ParallelStarted As Double
Public ParallelPaused As Boolean
Public ParallelResumed As Boolean
Public ParallelPendingPeak As Long
Public Sub ParallelControlHook()
    Dim elapsed As Double
    elapsed = GdtClockSeconds() - ParallelStarted
    If ParallelMode = "pause" Then
        If elapsed >= 0.15 And Not ParallelPaused Then
            SetGdtPaused True
            ParallelPaused = True
        End If
        If elapsed >= 1.1 And Not ParallelResumed Then
            SetGdtPaused False
            ParallelResumed = True
        End If
    ElseIf ParallelMode = "stop" And elapsed >= 0.15 Then
        RequestGdtStop
    End If
    cToken = "fixture-global-changed"
End Sub
Public Function ParallelSerial() As Double
    Dim index As Long, started As Double, result As clsGdtRequestResult
    ResetGdtRequestSession
    ResetGdtOperationControl
    started = GdtClockSeconds()
    For index = 1 To 8
        Set result = ExecuteGdtRequest("GET", ParallelBaseUrl & "/serial/api/query/invoices/detail?shdon=" & index, _
            "fixture-snapshot", , , , , 1)
        If Not result.Success Then Err.Raise 5, , "serial fixture failed"
    Next index
    ParallelSerial = GdtClockSeconds() - started
End Function
Public Function ParallelHeaderBaseline() As Boolean
    Dim response As clsGdtRequestResult
    ResetGdtRequestSession: ResetGdtOperationControl
    Set response = ExecuteGdtRequest("GET", ParallelBaseUrl & "/baseline/api/query/invoices/related?shdon=42", _
        "fixture-snapshot", , , , False, 1, "fixture", "fixture-action")
    ParallelHeaderBaseline = response.Success
End Function
'@)
    $module.CodeModule.AddFromString("Public Function ParallelConfigure() As Boolean`r`n ParallelBaseUrl = `"http://127.0.0.1:$port`"`r`n ParallelConfigure = True`r`nEnd Function`r`n")
    [void]$excel.Run("${prefix}ParallelConfigure")
    # Only this in-memory fixture changes retry count and routes known endpoints
    # to localhost. The delivered workbook is closed without saving these edits.
    $retryCode = $workbook.VBProject.VBComponents.Item('modGdtRetry').CodeModule
    for ($line=1; $line -le $retryCode.CountOfLines; $line++) {
        if ($retryCode.Lines($line,1) -match '^Public Const GDT_MAX_RETRIES ') { $retryCode.ReplaceLine($line, 'Public Const GDT_MAX_RETRIES As Long = 2'); break }
    }
    $pumpCode = $workbook.VBProject.VBComponents.Item('modGdtXmlScheduler').CodeModule
    for ($line=1; $line -le $pumpCode.CountOfLines; $line++) {
        if ($pumpCode.Lines($line,1) -match '^Public Sub GdtPumpWait') { $pumpCode.InsertLines($line+1, '    ParallelControlHook'); break }
    }
    $formCode = $workbook.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $redirected = $false
    for ($line=1; $line -le $formCode.CountOfLines; $line++) {
        if ($formCode.Lines($line,1) -match '^\s*BuildRelationEndpoint = ') {
            $formCode.ReplaceLine($line, '    BuildRelationEndpoint = ParallelBaseUrl & "/" & ParallelMode & "/api/" & apiSource & _')
            $redirected = $true; break
        }
    }
    if (-not $redirected) { throw 'Missing endpoint builder.' }
    for ($line=1; $line -le $formCode.CountOfLines; $line++) {
        if ($formCode.Lines($line,1).Trim() -eq 'tasks.Add task') {
            $formCode.InsertLines($line+1, '                If tasks.Count > ParallelPendingPeak Then ParallelPendingPeak = tasks.Count')
            break
        }
    }
    $formCode.AddFromString(@'
Public Function ParallelFixture(ByVal mode As String, ByVal invoiceType As Long) As String
    Dim buffer As Variant, index As Long, expected As Long, detailRow As Long, invoiceCount As Long
    Dim summary As Worksheet, detail As Worksheet, report As Worksheet, elapsed As Double
    On Error GoTo Failed
    ParallelMode = mode: ParallelPaused = False: ParallelResumed = False: ParallelPendingPeak = 0
    invoiceCount = IIf(mode = "buffer", 24, 8)
    ResetGdtRequestSession: ResetGdtOperationControl
    mRunToken = "fixture-snapshot": mRunPurchase = (invoiceType = 1): mDownloadRunning = True
    Me.optMua.Value = mRunPurchase: Me.txtSleep.Value = IIf(mode = "default", "600", "1")
    Set mFinalRetryQueue = New Collection
    mFinalRetryCooldownDone = True
    mInvoiceWorkTotal = invoiceCount * 3: mInvoiceWorkDone = 0
    Set summary = ThisWorkbook.Sheets(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban"))
    Set detail = ThisWorkbook.Sheets(IIf(invoiceType = 1, "ChiTietHD_Mua", "ChiTietHD_Ban"))
    Set report = ThisWorkbook.Sheets("BaoCao_LoiTaiHD")
    summary.Range("BE200:BL207").ClearContents
    detail.Range("A200:AH260").ClearContents
    report.Range("A2:Q1000").ClearContents
    ReDim buffer(0 To invoiceCount - 1, 0 To 7)
    For index = 0 To invoiceCount - 1
        buffer(index, 0) = "0000000001": buffer(index, 1) = "C26TST"
        buffer(index, 2) = CStr(index + 1): buffer(index, 3) = "1"
        buffer(index, 4) = 1 + (index Mod 2): buffer(index, 5) = DateSerial(2026, 9, 12)
        buffer(index, 6) = 200 + index
        Select Case index Mod 4
            Case 0: buffer(index, 7) = 2
            Case 1: buffer(index, 7) = 6
            Case 2: buffer(index, 7) = 1
            Case 3: buffer(index, 7) = 5
        End Select
    Next index
    detailRow = 200
    ParallelStarted = GdtClockSeconds()
    If Left$(mode, 8) = "relation" Then
        ProcessRelatedInvoiceApis buffer, 8, invoiceType
        If mInvoiceWorkDone <> 10 Then Err.Raise 5, , "related progress count"
        For index = 0 To 7
            If index Mod 4 = 2 Then
                If Len(CStr(summary.Cells(200 + index, 64).Value2)) > 0 Then Err.Raise 5, , "status 1 was fetched"
            Else
                If InStr(CStr(summary.Cells(200 + index, 64).Value2), "NOTICE-" & (index + 1)) = 0 Then Err.Raise 5, , "related wrong summary row"
                If index Mod 4 <> 1 Then
                    If InStr(CStr(summary.Cells(200 + index, 57).Value2), "REL-" & (index + 1)) = 0 Then Err.Raise 5, , "relative wrong summary row"
                Else
                    If InStr(CStr(summary.Cells(200 + index, 57).Value2), "REL-") > 0 Then Err.Raise 5, , "status 6 fetched relative"
                End If
            End If
        Next index
    Else
        ProcessJsonInvoiceRequests buffer, invoiceCount, invoiceType, True, detailRow
        If ParallelPendingPeak > 12 Then Err.Raise 5, , "unbounded response queue"
        If mode = "buffer" And ParallelPendingPeak <> 12 Then Err.Raise 5, , "buffer boundary was not exercised"
        If mode = "auth" Then
            If Not GdtAuthenticationFailed Then Err.Raise 5, , "auth did not stop session"
        ElseIf mode = "stop" Then
            If Not GdtStopRequested Then Err.Raise 5, , "stop was ignored"
        Else
            If mInvoiceWorkDone <> invoiceCount Then Err.Raise 5, , "detail progress count"
            expected = 200
            For index = 1 To invoiceCount
                If Not ((mode = "failure" Or mode = "invalid") And index = 2) Then
                    If CLng(detail.Cells(expected, 3).Value2) <> index Or CLng(detail.Cells(expected + 1, 3).Value2) <> index Then _
                        Err.Raise 5, , "detail order: expected " & index & " at row " & expected & "; got " & detail.Cells(expected, 3).Value2 & "/" & detail.Cells(expected + 1, 3).Value2 & "; cursor " & detailRow
                    If CStr(detail.Cells(expected, 8).Value2) <> "0000000001" Then Err.Raise 5, , "lost leading zero tax ID"
                    If CDbl(detail.Cells(expected, 4).Value2) <> CDbl(DateSerial(2026, 9, 12)) Then Err.Raise 5, , "wrong UTC+7 date"
                    expected = expected + 2
                End If
            Next index
            If detailRow <> expected Then Err.Raise 5, , "wrong detail cursor"
            If mode = "failure" Then
                If mFinalRetryQueue.Count <> 1 Then Err.Raise 5, , "exhausted failure not queued exactly once"
                ProcessQueuedDetailRetries detailRow
                If detailRow <> 216 Or CLng(detail.Cells(214, 3).Value2) <> 2 Then Err.Raise 5, , "final detail retry did not recover"
                If Len(CStr(report.Cells(2, 1).Value2)) > 0 Then Err.Raise 5, , "successful recovery kept error row"
            ElseIf mode = "invalid" Then
                If mFinalRetryQueue.Count <> 0 Or Len(CStr(report.Cells(2, 1).Value2)) = 0 Then Err.Raise 5, , "malformed JSON report/queue"
            Else
                If mFinalRetryQueue.Count <> 0 Then Err.Raise 5, , "successful detail was queued"
            End If
            If mode = "pause" And (Not ParallelPaused Or Not ParallelResumed) Then Err.Raise 5, , "pause/resume hook did not run"
        End If
    End If
    elapsed = GdtClockSeconds() - ParallelStarted
    ParallelFixture = "PASS|" & Format$(elapsed, "0.000")
CleanUp:
    mDownloadRunning = False
    ResetGdtOperationControl: ResetGdtRequestSession
    Exit Function
Failed:
    ParallelFixture = "FAIL: " & Err.Description
    If Not report Is Nothing Then ParallelFixture = ParallelFixture & "; first report: " & CStr(report.Cells(2, 13).Value2)
    Resume CleanUp
End Function
'@)
    foreach ($scenario in @(@('normal1',1),@('normal2',2),@('default',1),@('buffer',1),@('relation1',1),@('relation2',2),@('mixed',1),@('failure',1),@('invalid',1),@('pause',1),@('stop',1),@('auth',1))) {
        $mode = [string]$scenario[0]; $invoiceType = [int]$scenario[1]
        $entry = 'Parallel_' + $mode
        $module.CodeModule.AddFromString("Public Function $entry() As String`r`n Dim form As New frmTaiHoaDon`r`n $entry = form.ParallelFixture(`"$mode`", $invoiceType)`r`n Unload form`r`nEnd Function`r`n")
        # Editing a VBA module resets its globals; initialize after injection.
        [void]$excel.Run("${prefix}ParallelConfigure")
        if ($mode -eq 'normal1' -and -not [bool]$excel.Run("${prefix}ParallelHeaderBaseline")) { throw 'Synchronous header baseline failed.' }
        $actual = [string]$excel.Run("$prefix$entry")
        Write-Output "$mode=$actual"
        $results[$mode] = $actual
        if (-not $actual.StartsWith('PASS|')) { throw "$mode failed: $actual" }
        $headerStats = Invoke-RestMethod -Uri "http://127.0.0.1:$port/stats" -TimeoutSec 5
        if ($headerStats.BadHeaders -ne 0 -or $headerStats.DuplicateIds -ne 0) {
            throw ('Header regression in ' + $mode + ': ' + ($headerStats.BadHeaderKinds | ConvertTo-Json -Compress) + '; duplicate IDs=' + $headerStats.DuplicateIds)
        }
    }
    $serialSeconds = [double]$excel.Run("${prefix}ParallelSerial")
    $parallelSeconds = [double]::Parse((([string]$results['normal1'] -split '\|')[1]).Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture)
    $defaultSeconds = [double]::Parse((([string]$results['default'] -split '\|')[1]).Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture)
    if ($parallelSeconds -ge $serialSeconds * 0.85) { throw "No measured speedup: parallel=$parallelSeconds, serial=$serialSeconds" }
    $stats = Invoke-RestMethod -Uri "http://127.0.0.1:$port/stats" -TimeoutSec 5
    if ($stats.BadHeaders -ne 0 -or $stats.DuplicateIds -ne 0) { throw ('JSON header/request-id regression: ' + ($stats | ConvertTo-Json -Depth 6 -Compress)) }
    foreach ($mode in @('normal1','normal2','relation1','relation2')) {
        if ($stats.Peaks.$mode -lt 2 -or $stats.Peaks.$mode -gt 4) { throw "Concurrency was not bounded/parallel in $mode" }
    }
    if ($stats.Peaks.default -lt 2 -or $stats.Peaks.default -gt 4 -or $stats.MinIntervalsMs.default -lt 550) { throw 'Default form spacing/concurrency was not respected.' }
    foreach ($mode in @('stop','auth')) {
        $total = ($stats.Counts.PSObject.Properties | Where-Object Name -Like "$mode/*" | Measure-Object Value -Sum).Sum
        if ($total -gt 4) { throw "$mode launched new work after cancellation: $total" }
    }
    foreach ($mode in @('relation1','relation2')) {
        # Compare with the existing synchronous transport. WinHTTP can filter
        # an HTTPS Referer on these plain-HTTP localhost fixture requests.
        if ([string]$stats.Referers.$mode -cne [string]$stats.Referers.baseline) { throw 'Related Referer differs from synchronous baseline.' }
        foreach ($number in @(2,3,6,7)) {
            if ($stats.Counts.PSObject.Properties.Name -contains "$mode/relative/$number") { throw 'Unexpected relative request for status 1/6.' }
        }
    }
    $results['MeasuredHttp'] = [ordered]@{ SerialSeconds=$serialSeconds; ParallelSeconds=$parallelSeconds; TestIntervalMs=1; Speedup=($serialSeconds/$parallelSeconds); DefaultSeconds=$defaultSeconds; DefaultIntervalMs=600; DefaultSpeedup=($serialSeconds/$defaultSeconds); Peaks=$stats.Peaks; MinIntervalsMs=$stats.MinIntervalsMs; BadHeaders=$stats.BadHeaders; DuplicateIds=$stats.DuplicateIds }
    $results | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path (Split-Path -Parent $BuiltWorkbook) 'parallel-json-runtime-result.json')
    Write-Output ('MEASURED SPEEDUP: ' + [Math]::Round($serialSeconds/$parallelSeconds,2) + 'x (localhost fixture, not live GDT)')
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $formCode; Release-ComObject $module; Release-ComObject $workbook; Release-ComObject $excel
    if ($null -ne $server) { Stop-Job $server -ErrorAction SilentlyContinue; Remove-Job $server -Force -ErrorAction SilentlyContinue }
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -match '^hddt-json-parallel-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
Write-Output 'PARALLEL JSON RUNTIME TEST PASSED'
