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
    $module.Name = 'modAudit20261014'
    $code = $module.CodeModule
    $baseline = Get-Content -LiteralPath (Join-Path $repo 'build/Test-DetailInvoiceRuntime.ps1') -Raw -Encoding UTF8
    $fragment = [regex]::Match($baseline, '(?s)Private Function RunDetailInvoiceFixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow').Groups[1].Value
    if (-not $fragment) { throw 'Missing synthetic fixture' }
    $code.AddFromString("Public Function AuditDetailJson() As String`r`n Dim sampleJson As String, sellerTaxId As String, buyerTaxId As String`r`n sellerTaxId = `"0000000001`": buyerTaxId = `"0000000002`"`r`n$fragment`r`n AuditDetailJson = sampleJson`r`nEnd Function`r`n")
    $code.AddFromString(@'
Option Explicit
Public AuditPayload As String
Public AuditListEndpoints As String
Public Function AuditEmptyList(ByVal endpoint As String) As String
    AuditListEndpoints = AuditListEndpoints & endpoint & vbLf
    LastGdtStatus = 200: LastGdtAttempts = 1: LastGdtShouldQueue = False
    AuditEmptyList = "{""datas"":[],""state"":null}"
End Function
Public Function AuditFakeDetail() As clsGdtRequestResult
    Dim result As New clsGdtRequestResult
    result.Success = True: result.StatusCode = 200: result.Attempts = 1
    result.ResponseText = AuditPayload
    Set AuditFakeDetail = result
End Function
Public Function RunAudit(ByVal mode As String, ByVal invoiceType As Long) As String
    RunAudit = frmTaiHoaDon.AuditCase(mode, invoiceType)
End Function
Public Function RunSourceAudit(ByVal invoiceType As Long, ByVal selection As Long) As String
    RunSourceAudit = frmTaiHoaDon.AuditSourceCase(invoiceType, selection)
End Function
Public Function RunSourceClearAudit(ByVal invoiceType As Long, ByVal selection As Long) As String
    RunSourceClearAudit = frmTaiHoaDon.AuditSourceClearCase(invoiceType, selection)
End Function
'@)
    $formCode = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $text = $formCode.Lines(1, $formCode.CountOfLines)
    $pattern = 'Set detailResult = ExecuteGdtRequest\("GET", endpoint, RunToken\(\), vbNullString, _\r?\n\s*"application/json", "application/json, text/plain, \*/\*", False, GDT_MAX_RETRIES, "DETAIL", action\)'
    if (-not [regex]::IsMatch($text, $pattern)) { throw 'Fallback request seam not found' }
    $text = [regex]::Replace($text, $pattern, 'Set detailResult = AuditFakeDetail()')
    $text = [regex]::Replace($text, 'res = ApiGet\(url2, ListPageLabel\([^\r\n]+', 'res = AuditEmptyList(url2)')
    $text = $text.Replace('ret = MsgBoxUni(sMsg, vbYesNo + vbInformation, UniConvert("Thoong baso"))', 'ret = vbYes')
    $text = $text.Replace('MsgBoxUni UniConvert("Cos looxi phast sinh. Vui lofng kieerm tra BaoCao_LoiTaiHD."), vbExclamation, UniConvert("Thoong baso")', "' Audit suppresses the modal error notice only.")
    $formCode.DeleteLines(1, $formCode.CountOfLines)
    $formCode.AddFromString($text)
    $formCode.AddFromString(@'
Public Function AuditCase(ByVal mode As String, ByVal invoiceType As Long) As String
    On Error GoTo Failed
    EndGdtResume
    ResetGdtOperationControl
    ResetGdtRequestSession
    EnsureRuntimeControls
    mRunPurchase = (invoiceType = 1): mDownloadRunning = True: mRunToken = "synthetic"
    Set mFinalRetryQueue = New Collection
    GdtReportAccount = "AUDIT-ACCOUNT": GdtReportSession = "AUDIT-SESSION"
    Call LinkTraCuu: Call tenCotTraCuu
    arrTrangThai = ThisWorkbook.Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = ThisWorkbook.Sheets("LinkTraCuu").Range("O2:O11").Value
    Dim ws As Worksheet, summary As Worksheet, report As Worksheet, state As Worksheet
    Set ws = ThisWorkbook.Sheets(IIf(mRunPurchase, "ChiTietHD_Mua", "ChiTietHD_Ban"))
    Set summary = ThisWorkbook.Sheets(IIf(mRunPurchase, "TongHopHD_Mua", "TongHopHD_Ban"))
    Set report = EnsureGdtErrorReportSheet()
    ws.Range("A3:AJ200").ClearContents
    summary.Range("A3:BN200").ClearContents
    report.Range("A2:S200").ClearContents
    For Each state In ThisWorkbook.Worksheets
        If state.Name = "TrangThaiTaiHD" Then state.Range("A2:F200").ClearContents
    Next state
    Dim invoice As Object, page As Object, invoices As New Collection
    Set invoice = JsonConverter.ParseJSON(AuditDetailJson())
    invoice("nbmst") = "0000000001": invoice("khmshdon") = "1"
    invoice("khhdon") = "C26AUDIT": invoice("shdon") = 42
    If mode = "related-text" Then
        invoice("tthai") = 2: invoice("gchdgoc") = "=1+1": invoice("shdgoc") = "00042"
    End If
    AuditPayload = JsonConverter.ConvertToJson(invoice)
    Set page = CreateObject("Scripting.Dictionary")
    invoices.Add invoice: page.Add "datas", invoices
    BeginGdtResume GdtReportAccount, invoiceType, (mode <> "fresh-duplicate")
    Dim sr As Long, seq As Long, dr As Long, first As Long
    sr = 3: seq = 1: dr = 3
    If Not WriteSummaryPage("", sr, invoiceType, seq, "query", "synthetic-list", Empty, page) Then Err.Raise 5, , "Summary setup failed"
    ReDim arrHDChiTiet(1, 7)
    arrHDChiTiet(0, 0) = "0000000001": arrHDChiTiet(0, 1) = "C26AUDIT"
    arrHDChiTiet(0, 2) = "42": arrHDChiTiet(0, 3) = "1"
    arrHDChiTiet(0, 4) = 1: arrHDChiTiet(0, 5) = DateSerial(2026, 1, 1)
    arrHDChiTiet(0, 6) = 3: arrHDChiTiet(0, 7) = 2
    Dim task As clsGdtXmlTask, errorsBefore As Long, errorsAfter As Long, initialRows As Long
    Set task = New clsGdtXmlTask
    task.Configure 0, "http://example.invalid/detail?nbmst=0000000001", "synthetic"
    task.ResponseKind = "DETAIL"
    task.Result.Success = True: task.Result.StatusCode = 200: task.Result.Attempts = 1
    task.Result.ResponseText = AuditPayload
    Select Case mode
    Case "fresh-duplicate"
        HandleJsonInvoiceResult task, arrHDChiTiet, invoiceType, dr
        initialRows = dr - 3
        Dim xmlTask As New clsGdtXmlTask
        xmlTask.Configure 0, "http://example.invalid/export-xml", "synthetic"
        xmlTask.Result.StatusCode = 504: xmlTask.Result.Attempts = 5
        HandleXmlResult xmlTask
        FinalizeGdtResumeDetails
        AuditCase = "initialRows=" & initialRows & "; finalRows=" & (ws.Cells(ws.Rows.Count, 3).End(xlUp).Row - 2)
    Case "parse-erased"
        Dim broken As New clsGdtRequestResult
        broken.Success = True: broken.StatusCode = 200: broken.Attempts = 1: broken.ResponseText = "{bad"
        ProcessOneRelationRequest arrHDChiTiet, 0, invoiceType, "query", "relative", broken
        errorsBefore = CountGdtSessionErrors(CurrentDirectionName())
        broken.ResponseText = "{}"
        ProcessOneRelationRequest arrHDChiTiet, 0, invoiceType, "query", "related", broken
        If CountGdtSessionErrors(CurrentDirectionName()) <> 1 Then Err.Raise 5, , "Related success cleared relative parse error"
        HandleJsonInvoiceResult task, arrHDChiTiet, invoiceType, dr
        errorsAfter = CountGdtSessionErrors(CurrentDirectionName())
        AuditCase = "errorsBeforeDetail=" & errorsBefore & "; errorsAfterDetail=" & errorsAfter & "; relativeCell=" & CStr(summary.Cells(3, 57).Value2)
    Case "mismatched-detail"
        invoice("nbmst") = "0000000999"
        task.Result.ResponseText = JsonConverter.ConvertToJson(invoice)
        HandleJsonInvoiceResult task, arrHDChiTiet, invoiceType, dr
        AuditCase = "requestedSeller=0000000001; writtenSeller=" & CStr(ws.Cells(3, 8).Value2) & "; resumable=" & CanResumeGdtStage(arrHDChiTiet, 0, "DETAIL") & "; errors=" & CountGdtSessionErrors(CurrentDirectionName())
    Case "error-formula"
        UpsertGdtErrorReport CurrentDirectionName(), "query", "", "", "", "", Empty, "Download", "synthetic", 0, "=1+1", 1, 0, "Failed"
        AuditCase = "HasFormula=" & report.Cells(2, 13).HasFormula & "; stored=" & CStr(report.Cells(2, 13).Value2)
    Case "related-text"
        AuditCase = "originalNoteHasFormula=" & summary.Cells(3, 63).HasFormula & "; storedNote=" & CStr(summary.Cells(3, 63).Value2) & "; originalNumber=" & CStr(summary.Cells(3, 61).Value2)
    Case "legacy-source"
        HandleJsonInvoiceResult task, arrHDChiTiet, invoiceType, dr
        ws.Range("AI3:AJ4").ClearContents
        GdtReportAccount = IIf(invoiceType = 1, "0000000002", "0000000001")
        ws.Cells(3, IIf(invoiceType = 1, 14, 8)).Resize(2, 1).Value2 = GdtReportAccount
        ' The account match alone cannot identify the source of a legacy row.
        EndGdtResume
        BeginGdtResume GdtReportAccount, invoiceType, False, "query"
        first = dr
        ghiExcel_ChiTiet AuditPayload, dr, IIf(invoiceType = 1, "mua", "ban")
        CommitGdtDetailRows arrHDChiTiet, 0, first, dr
        FinalizeGdtResumeDetails
        AuditCase = "legacyRowsRetained=" & (ws.Cells(ws.Rows.Count, 3).End(xlUp).Row - 2)
        If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row <> 6 Then Err.Raise 5, , AuditCase
    Case "reset-validation"
        mDownloadRunning = False: mDownloadEntering = False
        Dim accounts(0 To 0, 0 To 3) As Variant
        accounts(0, 0) = "AUDIT-ACCOUNT": accounts(0, 1) = "": accounts(0, 2) = "Synthetic": accounts(0, 3) = "synthetic"
        Me.cboDonVi.List = accounts: Me.cboDonVi.ListIndex = 0
        Me.optMua.Value = mRunPurchase: Me.chkXmlZip.Value = False
        Me.txtTuNgay.Value = "01/01/2026": Me.txtDenNgay.Value = "31/01/2026"
        ThisWorkbook.Sheets("TrangThaiTaiHD").Cells(1, 1).Value2 = "Unsupported schema"
        cmdTaiHoaDon_Click
        ThisWorkbook.Sheets("TrangThaiTaiHD").Cells(1, 1).Value2 = "Identity (v1)"
        AuditCase = "summaryAfterRejectedSchema=" & Application.CountA(summary.Range("A3:BN3")) & "; detailAfterRejectedSchema=" & Application.CountA(ws.Range("A3:AJ3")) & "; error=" & CStr(report.Cells(report.Rows.Count, 1).End(xlUp).Offset(0, 12).Value2)
    End Select
    Select Case mode
    Case "fresh-duplicate": If initialRows <> 2 Or ws.Cells(ws.Rows.Count, 3).End(xlUp).Row - 2 <> initialRows Then Err.Raise 5, , AuditCase
    Case "parse-erased": If errorsBefore <> 1 Or errorsAfter <> 1 Then Err.Raise 5, , AuditCase
    Case "mismatched-detail": If dr <> 3 Or CanResumeGdtStage(arrHDChiTiet, 0, "DETAIL") Or CountGdtSessionErrors(CurrentDirectionName()) <> 1 Then Err.Raise 5, , AuditCase
    Case "related-text": If summary.Cells(3, 63).HasFormula Or CStr(summary.Cells(3, 63).Value2) <> "=1+1" Or CStr(summary.Cells(3, 61).Value2) <> "00042" Then Err.Raise 5, , AuditCase
    Case "error-formula": If report.Cells(2, 13).HasFormula Or CStr(report.Cells(2, 13).Value2) <> "=1+1" Then Err.Raise 5, , AuditCase
    Case "reset-validation": If Application.CountA(summary.Range("A3:BN3")) = 0 Then Err.Raise 5, , AuditCase
    End Select
    AuditCase = "PASS: " & AuditCase
    EndGdtResume
    mDownloadRunning = False: mRunToken = ""
    Exit Function
Failed:
    AuditCase = "ERROR " & Err.Number & ": " & Err.Description
    EndGdtResume
    mDownloadRunning = False: mRunToken = ""
End Function
Public Function AuditSourceCase(ByVal invoiceType As Long, ByVal selection As Long) As String
    On Error GoTo Failed
    EndGdtResume: ResetGdtOperationControl: ResetGdtRequestSession
    EnsureRuntimeControls
    mDownloadRunning = False
    Me.optMua.Value = (invoiceType = 1)
    Me.Controls("cboInvoiceSource").ListIndex = selection
    Me.txtTuNgay.Value = "01/01/2026": Me.txtDenNgay.Value = "31/01/2026"
    Me.chkCT.Value = False: Me.chkXmlZip.Value = False
    AuditListEndpoints = ""
    Set mFinalRetryQueue = New Collection
    SetDownloadControlState True
    If Me.Controls("cboInvoiceSource").Enabled Then Err.Raise 5, , "Source is editable during download"
    ' Changing the disabled control in code must not change the captured scope.
    Me.Controls("cboInvoiceSource").ListIndex = (selection + 1) Mod 3
    taiHoaDon_Total
    Dim hasQuery As Boolean, hasSco As Boolean, endpoint As String
    hasQuery = InStr(AuditListEndpoints, "/api/query/") > 0
    hasSco = InStr(AuditListEndpoints, "/api/sco-query/") > 0
    endpoint = IIf(invoiceType = 1, "/invoices/purchase?", "/invoices/sold?")
    If hasQuery <> (selection <> 2) Or hasSco <> (selection <> 1) Then Err.Raise 5, , "Wrong source requests: " & AuditListEndpoints
    If InStr(AuditListEndpoints, endpoint) = 0 Then Err.Raise 5, , "Wrong purchase/sales endpoint"
    SetDownloadControlState False
    If Not Me.Controls("cboInvoiceSource").Enabled Then Err.Raise 5, , "Source not restored"
    Dim pause As Object, stopButton As Object, source As Object
    Set pause = Me.Controls("cmdPauseResume"): Set stopButton = Me.Controls("cmdStopDownload")
    Set source = Me.Controls("cboInvoiceSource")
    If Abs(pause.Left - Me.cmdTaiHoaDon.Left) > 0.1 Or Abs(stopButton.Left + stopButton.Width - Me.cmdTaiHoaDon.Left - Me.cmdTaiHoaDon.Width) > 0.1 Then Err.Raise 5, , "Action column edges"
    If pause.Top <> stopButton.Top Or pause.Width <> stopButton.Width Or pause.Height <> stopButton.Height Then Err.Raise 5, , "Pause/stop alignment"
    If pause.Top < Me.cmdChiTietXML.Top + Me.cmdChiTietXML.Height + 6 Or pause.Top + pause.Height > Me.Controls("lblProgressMessage").Top Then Err.Raise 5, , "Action row overlap"
    If Me.Label21.Left + Me.Label21.Width > pause.Left - 6 Then Err.Raise 5, , "Attribution overlaps actions"
    If source.Left <> Me.cboTTHD.Left Or source.Width <> Me.cboTTHD.Width Or source.Top <> Me.cboDonVi.Top Then Err.Raise 5, , "Source grid alignment"
    If source.Font.Size <> Me.cboTTHD.Font.Size Or pause.Font.Name <> Me.cmdTaiHoaDon.Font.Name Or stopButton.Font.Size <> pause.Font.Size Then Err.Raise 5, , "Control font consistency"
    AuditSourceCase = "PASS"
    Exit Function
Failed:
    AuditSourceCase = "FAIL: " & Err.Description
    SetDownloadControlState False
End Function
Public Function AuditSourceClearCase(ByVal invoiceType As Long, ByVal selection As Long) As String
    On Error GoTo Failed
    mRunSource = IIf(selection = 1, "query", "sco-query")
    GdtReportAccount = "AUDIT-ACCOUNT"
    Dim ws As Worksheet, sheetName As Variant, col As Long, r As Long
    For Each sheetName In Array(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban"), IIf(invoiceType = 1, "ChiTietHD_Mua", "ChiTietHD_Ban"))
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        col = IIf(Left$(CStr(sheetName), 7) = "TongHop", 66, 36)
        ws.Range("A3:BN200").ClearContents
        For r = 3 To 7
            ws.Cells(r, 3).Value2 = r
            ws.Cells(r, col - 1).Value2 = "AUDIT-ACCOUNT"
            ws.Cells(r, col).Value2 = mRunSource
        Next r
        ws.Cells(4, col).Value2 = IIf(selection = 1, "sco-query", "query")
        ws.Cells(6, col - 1).Value2 = "OTHER-ACCOUNT"
        ws.Cells(7, col - 1).ClearContents: ws.Cells(7, col).ClearContents
    Next sheetName
    ClearSelectedSourceRows invoiceType
    For Each sheetName In Array(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban"), IIf(invoiceType = 1, "ChiTietHD_Mua", "ChiTietHD_Ban"))
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row <> 5 Then Err.Raise 5, , "Source clear row count"
        If ws.Cells(3, 3).Value2 <> 4 Or ws.Cells(4, 3).Value2 <> 6 Or ws.Cells(5, 3).Value2 <> 7 Then Err.Raise 5, , "Source clear removed other scope/legacy"
    Next sheetName
    AuditSourceClearCase = "PASS"
    Exit Function
Failed:
    AuditSourceClearCase = "FAIL: " & Err.Description
End Function
'@)
    foreach ($direction in @(1, 2)) {
    foreach ($mode in @('fresh-duplicate','parse-erased','mismatched-detail','error-formula','related-text','legacy-source','reset-validation')) {
        $result = [string]$excel.Run("${prefix}RunAudit", [string]$mode, [int]$direction)
        Write-Output "$mode => $result"
        if ($result -notlike 'PASS:*') { throw $result }
    }
    }
    foreach ($direction in @(1, 2)) {
        foreach ($selection in @(0, 1, 2)) {
            $result = [string]$excel.Run("${prefix}RunSourceAudit", [int]$direction, [int]$selection)
            Write-Output "source-$direction-$selection=$result"
            if ($result -ne 'PASS') { throw $result }
        }
        foreach ($selection in @(1, 2)) {
            $result = [string]$excel.Run("${prefix}RunSourceClearAudit", [int]$direction, [int]$selection)
            Write-Output "source-clear-$direction-$selection=$result"
            if ($result -ne 'PASS') { throw $result }
        }
    }
} finally {
    Release-ComObject $formCode; Release-ComObject $code; Release-ComObject $module
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Audit modified workbook artifact' }
Write-Output 'REVIEW FIXES RUNTIME TEST PASSED; WORKBOOK HASH UNCHANGED'
