[CmdletBinding()]
param(
    [string]$BuiltWorkbook = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\TaiHoaDonDienTu_v6.7.3.xlsm')
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'build\ExcelBuild.Common.psm1') -Force -DisableNameChecking
$root = Split-Path -Parent $PSScriptRoot
$testWorkbook = Join-Path $root 'dist\UpdateRuntime-test-only.xlsm'
$fixture = Join-Path $root 'dist\UpdateRuntime-response-test-only.json'
$portProbe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,0)
$portProbe.Start()
$port = ([Net.IPEndPoint]$portProbe.LocalEndpoint).Port
$portProbe.Stop()
$server = $null
$excel = $null
$workbook = $null

function Set-CodeLine([object]$Module, [string]$Pattern, [string]$Replacement) {
    for ($line = 1; $line -le $Module.CountOfLines; $line++) {
        if ($Module.Lines($line, 1) -match $Pattern) {
            $Module.ReplaceLine($line, $Replacement)
            return
        }
    }
    throw "Missing VBA line: $Pattern"
}

function Read-OpenResult([object]$Excel, [string]$Path) {
    $opened = $null
    try {
        $Excel.EnableEvents = $true
        $opened = $Excel.Workbooks.Open($Path, 0, $true)
        return [string]$Excel.Run("'$($opened.Name)'!CodexReadUpdateForm")
    } finally {
        if ($null -ne $opened) { $opened.Close($false); Release-ComObject $opened }
    }
}

try {
    if (-not (Test-Path -LiteralPath $BuiltWorkbook)) { throw "Missing workbook: $BuiltWorkbook" }
    Copy-Item -LiteralPath $BuiltWorkbook -Destination $testWorkbook -Force

    # This disposable copy uses a local metadata fixture and captures the
    # prepared form before Show, avoiding a hidden Excel modal UI during COM.
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($testWorkbook, 0, $false)
    $versionCode = $workbook.VBProject.VBComponents.Item('modVersion').CodeModule
    $versionCode.DeleteLines(5, 2)
    $versionCode.InsertLines(5, "Public Const UPDATE_URL As String = `"http://127.0.0.1:$port/update.json`"")
    $updateCode = $workbook.VBProject.VBComponents.Item('modUpdate').CodeModule
    Set-CodeLine $updateCode '^\s*notice\.Show vbModal' '    CodexCaptureUpdate notice'
    $testModule = $workbook.VBProject.VBComponents.Add(1)
    $testModule.Name = 'modCodexUpdateRuntime'
    $testModule.CodeModule.AddFromString(@'
Public CapturedUpdate As String
Public Sub CodexCaptureUpdate(ByVal frm As Object)
    CapturedUpdate = frm.Controls("lblCurrent").Caption & "|" & _
        frm.Controls("lblNew").Caption & "|" & _
        frm.Controls("lblDate").Caption & "|" & _
        frm.Controls("lblNotes").Caption
End Sub
Public Function CodexReadUpdateForm() As String
    CodexReadUpdateForm = CapturedUpdate
    CapturedUpdate = ""
End Function
'@)
    $workbook.Save()
    Release-ComObject $versionCode; Release-ComObject $updateCode; Release-ComObject $testModule
    $workbook.Close($false); Release-ComObject $workbook; $workbook = $null
    $excel.Quit(); Release-ComObject $excel; $excel = $null

    $server = Start-Job -ArgumentList $port,$fixture -ScriptBlock {
        param($listenPort,$responsePath)
        $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,[int]$listenPort)
        $listener.Start()
        try {
            while ($true) {
                $client = $listener.AcceptTcpClient()
                try {
                    $stream = $client.GetStream()
                    $reader = New-Object IO.StreamReader($stream)
                    $requestLine = $reader.ReadLine()
                    do { $line = $reader.ReadLine() } while ($null -ne $line -and $line.Length -gt 0)
                    $shutdown = $requestLine -match '^GET /shutdown '
                    $body = [IO.File]::ReadAllText($responsePath,[Text.Encoding]::UTF8)
                    $bytes = [Text.Encoding]::UTF8.GetBytes($body)
                    $header = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: application/json; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n")
                    $stream.Write($header,0,$header.Length)
                    $stream.Write($bytes,0,$bytes.Length)
                } finally { $client.Close() }
                if ($shutdown) { break }
            }
        } finally { $listener.Stop() }
    }

    [IO.File]::WriteAllText($fixture, '{"version":"6.7.4","releaseDate":"2026-09-24","downloadUrl":"https://github.com/dieutx/TaiHoaDonDienTu/releases/download/v6.7.4/TaiHoaDonDienTu_v6.7.4.xlsm","releaseNote":["Test update"]}', [Text.Encoding]::UTF8)
    Start-Sleep -Milliseconds 700
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $excel.EnableEvents = $false

    $newResult = Read-OpenResult $excel $testWorkbook
    if ($newResult -ne '6.7.3|6.7.4|24/09/2026|? Test update' -and $newResult -ne "6.7.3|6.7.4|24/09/2026|$([char]10003) Test update") { throw "New-version test failed: $newResult" }
    Write-Output "NEW_VERSION=PASS $newResult"

    [IO.File]::WriteAllText($fixture, '{"version":"6.7.3","downloadUrl":"https://github.com/dieutx/TaiHoaDonDienTu/releases/download/v6.7.3/TaiHoaDonDienTu_v6.7.3.xlsm"}', [Text.Encoding]::UTF8)
    $sameResult = Read-OpenResult $excel $testWorkbook
    if ($sameResult) { throw "Same-version test failed: $sameResult" }
    Write-Output 'SAME_VERSION=PASS no form'

    [IO.File]::WriteAllText($fixture, '{bad json', [Text.Encoding]::UTF8)
    $badResult = Read-OpenResult $excel $testWorkbook
    if ($badResult) { throw "Malformed-JSON test failed: $badResult" }
    Write-Output 'BAD_JSON=PASS no form'

    $shutdownRequest = [Net.WebRequest]::Create("http://127.0.0.1:$port/shutdown")
    $shutdownResponse = $shutdownRequest.GetResponse()
    $shutdownResponse.Close()
    Wait-Job $server -Timeout 5 | Out-Null
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $offlineResult = Read-OpenResult $excel $testWorkbook
    $watch.Stop()
    if ($offlineResult) { throw "Offline test failed: $offlineResult" }
    Write-Output "OFFLINE=PASS no form; $($watch.ElapsedMilliseconds) ms"
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $workbook; Release-ComObject $excel
    if ($null -ne $server) {
        if ($server.State -eq 'Running') {
            try { $response = [Net.WebRequest]::Create("http://127.0.0.1:$port/shutdown").GetResponse(); $response.Close() } catch {}
            Wait-Job $server -Timeout 5 | Out-Null
        }
        Remove-Job $server -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $testWorkbook) { Remove-Item -LiteralPath $testWorkbook -Force }
    if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Force }
}
