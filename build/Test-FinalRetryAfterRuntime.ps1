[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'build/ExcelBuild.Common.psm1') -Force -DisableNameChecking
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$probe = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
$probe.Start(); $port = ([Net.IPEndPoint]$probe.LocalEndpoint).Port; $probe.Stop()
$server = Start-Job -ArgumentList $port -ScriptBlock {
    param($port)
    $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $port)
    $listener.Start()
    try {
        while ($true) {
            if (-not $listener.Pending()) { Start-Sleep -Milliseconds 5; continue }
            $client = $listener.AcceptTcpClient()
            try {
                $stream = $client.GetStream()
                $reader = New-Object IO.StreamReader($stream)
                $first = $reader.ReadLine()
                do { $line = $reader.ReadLine() } while ($null -ne $line -and $line.Length -gt 0)
                if ($first -match 'GET /ready ') {
                    $response = "HTTP/1.1 200 OK`r`nContent-Length: 5`r`nConnection: close`r`n`r`nready"
                } else {
                    Write-Output ([DateTime]::UtcNow.ToString('O'))
                    $response = "HTTP/1.1 429 Too Many Requests`r`nRetry-After: 5`r`nContent-Length: 7`r`nConnection: close`r`n`r`nlimited"
                }
                $bytes = [Text.Encoding]::ASCII.GetBytes($response)
                $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()
            } finally { $client.Close() }
        }
    } finally { $listener.Stop() }
}
$excel = $null; $book = $null; $formCode = $null; $taskCode = $null; $retryCode = $null; $code = $null; $module = $null
try {
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        try { $null = Invoke-WebRequest "http://127.0.0.1:$port/ready" -UseBasicParsing -TimeoutSec 1; break } catch {
            if ([DateTime]::UtcNow -gt $deadline) { throw }
            Start-Sleep -Milliseconds 100
        }
    } while ($true)
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $formCode = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $text = $formCode.Lines(1, $formCode.CountOfLines).Replace('https://hoadondientu.gdt.gov.vn/api/', "http://127.0.0.1:$port/api/")
    $formCode.DeleteLines(1, $formCode.CountOfLines); $formCode.AddFromString($text)
    $taskCode = $book.VBProject.VBComponents.Item('clsGdtXmlTask').CodeModule
    $text = $taskCode.Lines(1, $taskCode.CountOfLines).Replace('mMaxAttempts = maxAttempts', 'mMaxAttempts = 1')
    $taskCode.DeleteLines(1, $taskCode.CountOfLines); $taskCode.AddFromString($text)
    $retryCode = $book.VBProject.VBComponents.Item('modGdtRetry').CodeModule
    $text = $retryCode.Lines(1, $retryCode.CountOfLines).Replace('GDT_FINAL_RETRY_COOLDOWN_SECONDS As Long = 30', 'GDT_FINAL_RETRY_COOLDOWN_SECONDS As Long = 1')
    $retryCode.DeleteLines(1, $retryCode.CountOfLines); $retryCode.AddFromString($text)
    $formCode.AddFromString(@'
Public Function AuditRetryAfter() As String
    On Error GoTo Failed
    ResetGdtOperationControl: ResetGdtRequestSession: EndGdtResume
    EnsureRuntimeControls
    mRunPurchase = True: mRunToken = "synthetic": mDownloadRunning = True
    GdtReportAccount = "AUDIT": GdtReportSession = "RETRY-AFTER"
    Set mFinalRetryQueue = New Collection: mFinalRetryCooldownDone = False
    ReDim arrHDChiTiet(1, 7)
    arrHDChiTiet(0, 0) = "0000000001": arrHDChiTiet(0, 1) = "C26AUDIT"
    arrHDChiTiet(0, 2) = "42": arrHDChiTiet(0, 3) = "1"
    arrHDChiTiet(0, 4) = 1: arrHDChiTiet(0, 5) = DateSerial(2026, 1, 1)
    Me.txtSleep.Value = "0"
    mInvoiceWorkDone = 0: mInvoiceWorkTotal = 1
    taiXML_zip 1
    AuditRetryAfter = "DONE"
    mDownloadRunning = False: mRunToken = ""
    Exit Function
Failed:
    AuditRetryAfter = "ERROR: " & Err.Number & ": " & Err.Description
    mDownloadRunning = False: mRunToken = ""
End Function
'@)
    $module = $book.VBProject.VBComponents.Add(1); $module.Name = 'modAuditRetryAfter'
    $code = $module.CodeModule
    $code.AddFromString("Public Function RunAuditRetryAfter() As String`r`nRunAuditRetryAfter = frmTaiHoaDon.AuditRetryAfter()`r`nEnd Function")
    $result = [string]$excel.Run("${prefix}RunAuditRetryAfter")
    if ($result -ne 'DONE') { throw $result }
    $hits = @(Receive-Job $server)
    if ($hits.Count -ne 2) { throw "Expected two requests, got $($hits.Count)" }
    $delta = ([DateTime]::Parse($hits[1]) - [DateTime]::Parse($hits[0])).TotalSeconds
    Write-Output ("RETRY-AFTER: header=5 seconds; fixture final cooldown=1 second; actual request gap=" + $delta.ToString('0.000') + " seconds")
    if ($delta -lt 5) { throw 'Final retry ignored Retry-After' }
} finally {
    Stop-Job $server; Remove-Job $server -Force
    Release-ComObject $formCode; Release-ComObject $taskCode; Release-ComObject $retryCode; Release-ComObject $code; Release-ComObject $module
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Audit modified workbook artifact' }
Write-Output 'FINAL RETRY-AFTER RUNTIME TEST PASSED; WORKBOOK HASH UNCHANGED'
