[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'build/ExcelBuild.Common.psm1') -Force -DisableNameChecking
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$excel = $null; $book = $null; $code = $null; $formCode = $null; $module = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'modProgressClock20261014'
    $code = $module.CodeModule
    $code.AddFromString(@'
Option Explicit
Public Function RunClockAudit() As String
    RunClockAudit = frmTaiHoaDon.ClockAudit()
End Function
'@)
    $formCode = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $formCode.AddFromString(@'

' The clock audit drives the real form members: it exercises ResetProgressUI,
' BeginInvoiceProgress, the pump-driven TickElapsedTime and the dashboard
' counters the redesigned progress area shows.
Public Function ClockAudit() As String
    Dim first As String, second As String, progress As String, logText As String
    Dim done As String, remaining As String, phase As String, eta As String
    On Error GoTo Failed
    EnsureRuntimeControls
    ResetGdtOperationControl
    ResetGdtRequestSession
    ResetProgressUI
    mDownloadRunning = True
    mRunStarted = GdtClockSeconds()
    BeginInvoiceProgress 3, 0
    first = Me.lblStatus.caption
    WaitGdtSeconds 3
    second = Me.lblStatus.caption
    AdvanceInvoiceWork UniConvert("DDax xuwr lys chi tieest ") & 1 & "/" & mInvoiceCount
    progress = Me.Controls("lblProgressMessage").caption
    logText = Me.Controls("txtSimpleLog").Value
    done = Me.Controls("lblMetricDone").caption
    remaining = Me.Controls("lblMetricRemaining").caption
    phase = Me.Controls("lblDashPhase2").caption
    eta = Me.Controls("lblMetricEta").caption
    mDownloadRunning = False
    If first <> "00:00:00" Then Err.Raise 5, , "Clock did not reset: " & first
    If second = first Then Err.Raise 5, , "Clock did not tick while waiting: " & first
    If CLng(Right$(second, 2)) < 2 Then Err.Raise 5, , "Clock seconds did not advance: " & second
    If Me.lblStatus.Left <> Me.Controls("lblMetricEta").Left - 182 Then Err.Raise 5, , "Clock metric column shifted"
    If Me.lblStatus.Top <> Me.Controls("lblMetricEta").Top Then Err.Raise 5, , "Clock metric row shifted"
    If Me.lblStatus.Height <> Me.Controls("lblMetricEta").Height Then Err.Raise 5, , "Clock metric height differs"
    If Me.lblStatus.TextAlign <> 1 Then Err.Raise 5, , "Clock metric is not left aligned"
    If done <> "0/3" Then Err.Raise 5, , "Processed metric wrong: " & done
    If remaining <> "3" Then Err.Raise 5, , "Remaining metric wrong: " & remaining
    If InStr(phase, "0/3 (0,0%)") = 0 Then Err.Raise 5, , "Phase percent wrong: " & phase
    If Len(eta) = 0 Then Err.Raise 5, , "ETA placeholder missing"
    If InStr(logText, UniConvert("Toorng soos hosa ddown caafn xuwr lys: ") & 3) = 0 Then Err.Raise 5, , "Total not logged"
    If InStr(progress, "1/3") = 0 Then Err.Raise 5, , "Detail progress total missing: " & progress
    ClockAudit = "PASS|" & first & "|" & second & "|" & done & "|" & remaining & "|" & phase
    Exit Function
Failed:
    mDownloadRunning = False
    ClockAudit = "FAIL: " & Err.Description
End Function
'@)
    $result = [string]$excel.Run("${prefix}RunClockAudit")
    Write-Output $result
    if ($result -notlike 'PASS*') { throw $result }
} finally {
    Release-ComObject $formCode; Release-ComObject $code; Release-ComObject $module
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Clock audit modified workbook artifact' }
Write-Output 'PROGRESS CLOCK RUNTIME TEST PASSED; WORKBOOK HASH UNCHANGED'
