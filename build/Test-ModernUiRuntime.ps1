[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BuiltWorkbook,
    [string]$PreviewPath,
    [ValidateSet('idle', 'listing', 'running', 'finished')][string]$PreviewState = 'finished'
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$excel = $null; $book = $null; $formCode = $null; $module = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $formCode = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $formCode.AddFromString(@'

Public Function ModernUiAudit() As String
    Dim table As Object, j As Long, beforeControls As Long, expandedHeight As Single
    On Error GoTo Failed
    ResetGdtOperationControl
    beforeControls = Me.Controls.Count
    EnsureRuntimeControls
    EnsureRuntimeControls
    If Me.Controls.Count <> beforeControls Then Err.Raise 5, , "Repeated initialization added controls"
    ResetProgressUI
    If Not Me.lblStatus.Visible Or Me.lblStatus.Caption <> "00:00:00" Then Err.Raise 5, , "Idle clock hidden or stale"
    If Me.InsideWidth < 790 Or Me.InsideHeight < 700 Then Err.Raise 5, , "Form content clipped"
    If Me.chkXmlZip.Parent.Name <> "fraLoaiDuLieu" Then Err.Raise 5, , "XML selection parent changed"
    If Me.chkXmlZip.Left + Me.chkXmlZip.Width > Me.fraLoaiDuLieu.InsideWidth + 1 Then Err.Raise 5, , "XML selection clipped by parent"
    If Me.fraLoaiDuLieu.Left + Me.fraLoaiDuLieu.Width > Me.txtXMLFolderPath.Left Then Err.Raise 5, , "XML selection overlaps folder"
    If Me.Label21.Left + Me.Label21.Width > UiCtl("cmdPauseResume").Left - 6 Then Err.Raise 5, , "XML notice overlaps actions"
    If UiCtl("cmdStopDownload").Top + UiCtl("cmdStopDownload").Height > UiCtl("lblAdvancedTitle").Top And _
        UiCtl("lblAdvancedTitle").Left + UiCtl("lblAdvancedTitle").Width > UiCtl("cmdStopDownload").Left Then Err.Raise 5, , "Actions overlap advanced options"
    If UiPercentText(75, 862) <> "8,7%" Or UiPercentText(900, 862) <> "100,0%" Then Err.Raise 5, , "Percentage rounding"
    If UiEtaText(7, 862, 120) <> "" Or UiEtaText(75, 862, 9) <> "" Then Err.Raise 5, , "Premature ETA"
    If UiEtaText(75, 862, 120) <> GdtElapsedText(1259) Then Err.Raise 5, , "ETA calculation"
    Set table = UiCtl("lstLog")
    AppendLogRow "FIXTURE-001", "Downloaded", "0,2s", "x2"
    If table.ListCount <> 1 Or table.ColumnCount <> 5 Then Err.Raise 5, , "Log dimensions"
    If Len(CStr(table.List(0, 0))) <> 8 Or table.List(0, 1) <> "FIXTURE-001" Or _
        table.List(0, 2) <> "Downloaded" Or table.List(0, 3) <> "0,2s" Or table.List(0, 4) <> "x2" Then Err.Raise 5, , "Log column data"
    For j = 1 To 405
        AppendLogRow CStr(j), "fixture", "", ""
    Next j
    If table.ListCount <> 400 Or table.List(399, 1) <> "405" Then Err.Raise 5, , "Log retention"
    Me.Show 0
    DoEvents
    expandedHeight = table.Height
    If Abs(table.Width - 728) > 1 Or table.Height < 72 Or table.Height > 92 Or Not table.IntegralHeight Then Err.Raise 5, , "Expanded log must show complete rows: " & table.Height
    If table.Top + table.Height > UiCtl("lblLogDetail").Top Then Err.Raise 5, , "Log overlaps selected detail"
    If UiCtl("lblLogDetail").Width <> 728 Or UiCtl("lblLogDetail").Height <> 36 Then Err.Raise 5, , "Selected log detail clipped"
    If UiCtl("lblDashPhase1").Width <> 302 Or UiCtl("lblDashPhase2").Width <> 298 Then Err.Raise 5, , "Phase labels clipped"
    Me.Hide
    ToggleTechnicalLog
    If table.Visible Or Not UiCtl("txtSimpleLog").Visible Then Err.Raise 5, , "Technical log toggle"
    ToggleTechnicalLog
    If Not table.Visible Or UiCtl("txtSimpleLog").Visible Then Err.Raise 5, , "Structured log restore"
    table.ListIndex = 10
    Dim selectedText As String
    selectedText = CStr(table.List(10, 1))
    AppendLogRow "new row", "new status", "0,2s", ""
    If CStr(table.List(table.ListIndex, 1)) <> selectedText Then Err.Raise 5, , "Incoming log stole selected row"
    If InStr(UiCtl("lblLogDetail").Caption, selectedText) = 0 Then Err.Raise 5, , "Selected row detail not updated"
    ResetProgressUI
    mDownloadRunning = True
    mRunStarted = GdtClockSeconds() - 120
    ReDim arrDate(1 To 1, 1 To 2)
    arrDate(1, 1) = DateSerial(2026, 10, 14): arrDate(1, 2) = DateSerial(2026, 10, 15)
    LastGdtStatus = 200: LastGdtAttempts = 1
    For j = 1 To 12
        ReportListPageStart 1, 1, "sco-query", j
        ReportListPageComplete 1, 1, "sco-query", j, 50, j * 50, Timer - 0.2, True
    Next j
    ReportListPageStart 1, 1, "query", 1
    ReportListPageComplete 1, 1, "query", 1, 20, 20, Timer - 0.2, False
    ReportListPageStart 1, 1, "sco-query", 13
    If UiCtl("lblMetricDone").Caption <> "620" Or UiCtl("lblMetricRemaining").Caption <> "13" Then Err.Raise 5, , "Live list counts"
    If InStr(UiCtl("lblDashPhase1").Caption, "trang 13") = 0 Then Err.Raise 5, , "Current list page missing"
    If UiCtl("lblDashPhase2").Caption <> UniConvert("2. Chi tieest: chowf danh sasch") Then Err.Raise 5, , "Detail should wait for list"
    If InStr(UiCtl("lblProgressPercent").Caption, "%") > 0 Or UiCtl("lblProgressFill").Width <> 0 Then Err.Raise 5, , "Unknown list total displayed as percent"
    If UiCtl("lblMetricEta").Caption <> "-" Or UiCtl("lblSummaryCard").Visible Then Err.Raise 5, , "Premature ETA/summary during pagination"
    If Len(CStr(table.List(0, 1))) > 45 Then Err.Raise 5, , "List page label not compact"
    mRunFinished = True: mDownloadRunning = False
    SetRunState UI_STATE_STOPPED
    If UiCtl("lblSumFound").Caption <> "620" Or UiCtl("lblProgressPercent").Caption <> UniConvert("Danh sasch chuwa hoafn taast") Then Err.Raise 5, , "Stopped pagination lost received counts"
    mRunFinished = False: mDownloadRunning = True
    RefreshSummary
    BeginInvoiceProgress 862, 0
    If UiCtl("lblMetricDone").Caption <> "0/862" Or UiCtl("lblMetricRemaining").Caption <> "862" Then Err.Raise 5, , "List to detail transition"
    mDetailDone = 75
    mDetailStartedAt = GdtClockSeconds() - 15
    RefreshDashboard
    If UiCtl("lblMetricEta").Caption = GdtElapsedText(1259) Then Err.Raise 5, , "ETA includes list time"
    If UiCtl("lblMetricDone").Caption <> "75/862" Or UiCtl("lblMetricRemaining").Caption <> "787" Then Err.Raise 5, , "Running counts"
    If InStr(UiCtl("lblDashPhase2").Caption, "75/862 (8,7%)") = 0 Then Err.Raise 5, , "Running phase"
    TogglePauseDownload
    If Not GdtPauseRequested Or mRunState <> UI_STATE_PAUSED Then Err.Raise 5, , "Pause state"
    TogglePauseDownload
    If GdtPauseRequested Or mRunState <> UI_STATE_RUNNING Then Err.Raise 5, , "Resume state"
    RequestStopDownload
    If Not GdtStopRequested Or UiCtl("cmdStopDownload").Enabled Then Err.Raise 5, , "Stop state"
    ResetGdtOperationControl
    mDownloadRunning = False
    mRunFinishedAt = mRunStarted + 120
    mRunFinished = True
    mXmlMissing = 9: mFailedCount = 5
    SetRunState UI_STATE_FAILED
    RefreshDashboard
    RefreshSummary
    If UiCtl("lblSumFound").Caption <> "862" Or UiCtl("lblSumDownloaded").Caption <> "75" Or _
        UiCtl("lblSumNoXml").Caption <> "9" Or UiCtl("lblSumFailed").Caption <> "5" Then Err.Raise 5, , "Summary counts"
    If Not UiCtl("cmdRetryFailed").Enabled Then Err.Raise 5, , "Retry guidance unavailable"
    If Not UiCtl("lblSummaryCard").Visible Or table.Height < 30 Or table.Height > 60 Then Err.Raise 5, , "Finished summary missing"
    ToggleDownloadSummary
    If UiCtl("lblSummaryCard").Visible Or Abs(table.Height - expandedHeight) > 1 Then Err.Raise 5, , "Finished log cannot expand"
    ToggleDownloadSummary
    If Not UiCtl("lblSummaryCard").Visible Then Err.Raise 5, , "Finished summary cannot restore"
    If UiCtl("lblMetricEta").Caption <> "-" Or UiCtl("lblSumElapsed").Caption <> "00:02:00" Then Err.Raise 5, , "Finished clock/ETA"
    ResetProgressUI
    If UiCtl("lblSumFound").Caption <> "-" Or UiCtl("cmdRetryFailed").Enabled Then Err.Raise 5, , "New session retains old summary"
    ModernUiAudit = "PASS: initialization, live pagination, phase transition, layout, percentages, ETA, selected log, retention, pause/resume/stop, summary toggle/reset" & _
        "; log width=" & table.Width & "; height=" & table.Height & "; columns=" & table.ColumnWidths
    Exit Function
Failed:
    mDownloadRunning = False
    ResetGdtOperationControl
    ModernUiAudit = "FAIL: " & Err.Description
End Function

Public Sub ModernUiPreview(ByVal stateName As String)
    Dim accounts(0 To 0, 0 To 3) As Variant
    Dim page As Long
    accounts(0, 0) = "FIXTURE": accounts(0, 2) = "Synthetic test company"
    Me.cboDonVi.List = accounts: Me.cboDonVi.ListIndex = 0
    Me.txtXMLFolderPath.Value = "C:\HoaDon\"
    Me.txtTuNgay.Value = "14/10/2026": Me.txtDenNgay.Value = "15/10/2026"
    Me.Caption = "HDDT Modern UI Preview"
    ResetGdtOperationControl
    ResetProgressUI
    ' Match the real workflow: the form is already visible when requests run.
    Me.Show 0
    DoEvents
    If stateName = "listing" Then
        SetDownloadControlState True
        SetRunState UI_STATE_RUNNING
        mRunStarted = GdtClockSeconds() - 144
        ReDim arrDate(1 To 1, 1 To 2)
        arrDate(1, 1) = DateSerial(2026, 10, 14): arrDate(1, 2) = DateSerial(2026, 10, 15)
        LastGdtStatus = 200: LastGdtAttempts = 1
        For page = 1 To 12
            ReportListPageStart 1, 1, "sco-query", page
            ReportListPageComplete 1, 1, "sco-query", page, 50, page * 50, Timer - 1.2, True
        Next page
        ReportListPageStart 1, 1, "sco-query", 13
    ElseIf stateName <> "idle" Then
        SetDownloadControlState True
        mRunStarted = GdtClockSeconds() - 120
        BeginInvoiceProgress 862, 0
        mDetailDone = 75
        mDetailStartedAt = mRunStarted + 30
        SetProgress 18, UniConvert("DDax xuwr lys chi tieest ") & "75/862"
        SetRunState UI_STATE_RUNNING
        AppendLogRow "FIXTURE-001 (73/862)", UniConvert("DDax tari chi tieest"), "0,2s", ""
        AppendLogRow "FIXTURE-002 (74/862)", UniConvert("DDax tari chi tieest"), "1,4s", "x2"
        AppendLogRow "FIXTURE-003 (75/862)", UniConvert("DDax tari chi tieest"), "0,1s", ""
    End If
    If stateName = "finished" Then
        SetDownloadControlState False
        mRunFinishedAt = mRunStarted + 120
        mRunFinished = True
        mXmlMissing = 9: mFailedCount = 5
        SetRunState UI_STATE_FAILED
    End If
    RefreshDashboard
    RefreshSummary
    DoEvents
End Sub
'@)
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'modModernUiAudit'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function RunModernUiAudit() As String
    RunModernUiAudit = frmTaiHoaDon.ModernUiAudit()
End Function
Public Sub ShowModernUiPreview(ByVal stateName As String)
    frmTaiHoaDon.ModernUiPreview stateName
End Sub
Public Sub CloseModernUiAudit()
    Unload frmTaiHoaDon
End Sub
'@)
    $prefix = "'$($book.Name)'!"
    $result = [string]$excel.Run("${prefix}RunModernUiAudit")
    Write-Output $result
    if ($result -notlike 'PASS:*') { throw $result }
    if (-not [string]::IsNullOrWhiteSpace($PreviewPath)) {
        Add-Type -AssemblyName System.Drawing
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public class ModernUiRender {
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string cls,string title);
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr handle,out Rect bounds);
 [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr handle,IntPtr bounds,IntPtr region,uint flags);
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr handle,IntPtr target,uint flags);
}
'@
        $null = $excel.Run("${prefix}ShowModernUiPreview", $PreviewState)
        $handle = [ModernUiRender]::FindWindow('ThunderDFrame', 'HDDT Modern UI Preview')
        if ($handle -eq [IntPtr]::Zero) { throw 'Preview window not found' }
        # Repaint child windows after their enabled states change; otherwise
        # PrintWindow can capture empty cached frame/label surfaces.
        $null = [ModernUiRender]::RedrawWindow($handle, [IntPtr]::Zero, [IntPtr]::Zero, 389)
        Start-Sleep -Milliseconds 200
        $bounds = New-Object ModernUiRender+Rect
        $null = [ModernUiRender]::GetWindowRect($handle, [ref]$bounds)
        $bitmap = New-Object Drawing.Bitmap(($bounds.Right-$bounds.Left), ($bounds.Bottom-$bounds.Top))
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $dc = $graphics.GetHdc()
            try { if (-not [ModernUiRender]::PrintWindow($handle, $dc, 0)) { throw 'Preview rendering failed' } }
            finally { $graphics.ReleaseHdc($dc) }
            $bitmap.Save([IO.Path]::GetFullPath($PreviewPath), [Drawing.Imaging.ImageFormat]::Png)
        } finally { $graphics.Dispose(); $bitmap.Dispose() }
        Write-Output "PREVIEW: $PreviewPath"
    }
    $null = $excel.Run("${prefix}CloseModernUiAudit")
} finally {
    Release-ComObject $formCode; Release-ComObject $module
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Modern UI audit modified workbook artifact' }
Write-Output 'MODERN UI RUNTIME TEST PASSED; WORKBOOK HASH UNCHANGED'
