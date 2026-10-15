[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$testRoot = Join-Path $env:TEMP ('hddt-core-fixes-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$reopenPath = Join-Path $testRoot 'TaiHoaDonDienTu_v6.7.6.xlsm'
$utf8 = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $testRoot 'sample.xml'), '<HDon/>', $utf8)
[IO.File]::WriteAllText((Join-Path $testRoot 'sample.html'), '<html>fixture</html>', $utf8)
Compress-Archive -LiteralPath @((Join-Path $testRoot 'sample.xml'), (Join-Path $testRoot 'sample.html')) -DestinationPath (Join-Path $testRoot '0000000001_1_C26TST_42.zip')
Copy-Item -LiteralPath (Join-Path $testRoot 'sample.xml') -Destination (Join-Path $testRoot '0000000001_1_C26TST_42.xml')
Copy-Item -LiteralPath (Join-Path $testRoot 'sample.html') -Destination (Join-Path $testRoot '0000000001_1_C26TST_42.html')
$excel = $null; $book = $null; $component = $null; $code = $null; $formCode = $null
$results = [ordered]@{}
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $component = $book.VBProject.VBComponents.Add(1)
    $component.Name = 'modCoreFixTests'
    $code = $component.CodeModule
    $baseline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-DetailInvoiceRuntime.ps1') -Raw -Encoding UTF8
    foreach ($kind in @('SummaryDate','DetailInvoice')) {
        $fragment = [regex]::Match($baseline, "(?s)Private Function Run${kind}Fixture.*?(    sampleJson = .*?)\r?\n\r?\n    writeRow").Groups[1].Value
        if (-not $fragment) { throw "Missing $kind synthetic fixture" }
        $code.AddFromString("Public Function Core${kind}Json() As String`r`n Dim sampleJson As String, sellerTaxId As String, buyerTaxId As String`r`n sellerTaxId = `"0000000001`": buyerTaxId = `"0000000002`"`r`n$fragment`r`n Core${kind}Json = sampleJson`r`nEnd Function`r`n")
    }
    $code.AddFromString(@'
Option Explicit
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise 5, , message
End Sub
Private Sub Initialize(ByVal invoiceType As Long)
    EndGdtResume
    ResetGdtOperationControl
    ResetGdtRequestSession
    Call LinkTraCuu: Call tenCotTraCuu
    arrTrangThai = ThisWorkbook.Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = ThisWorkbook.Sheets("LinkTraCuu").Range("O2:O11").Value
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "TrangThaiTaiHD" Then ws.Range("A2:F2000").ClearContents
    Next ws
    Set ws = SummarySheet(invoiceType): ws.Range("A3:BN2000").ClearContents
    Set ws = DetailSheet(invoiceType): ws.Range("A3:AJ2000").ClearContents
End Sub
Private Function SummarySheet(ByVal invoiceType As Long) As Worksheet
    Set SummarySheet = ThisWorkbook.Sheets(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban"))
End Function
Private Function DetailSheet(ByVal invoiceType As Long) As Worksheet
    Set DetailSheet = ThisWorkbook.Sheets(IIf(invoiceType = 1, "ChiTietHD_Mua", "ChiTietHD_Ban"))
End Function
Private Sub WritePage(ByVal page As Object, ByVal source As String, ByVal invoiceType As Long, _
    ByRef nextRow As Long, ByRef buffer As Variant)
    Dim targets As Object, seq As Long, start As Long, item As Object, identity As String
    start = nextRow: seq = 1
    PrepareGdtResumePage page, source, targets
    ghiExcel_TongHop "", nextRow, invoiceType, seq, page, targets
    CommitGdtResumePage page, source, targets, start, invoiceType
    If page("datas").Count = 0 Then Exit Sub
    Set item = page("datas")(1)
    ReDim buffer(0 To 0, 0 To 7)
    buffer(0, 0) = SafeJsonText(item, "nbmst"): buffer(0, 1) = SafeJsonText(item, "khhdon")
    buffer(0, 2) = SafeJsonText(item, "shdon"): buffer(0, 3) = SafeJsonText(item, "khmshdon")
    buffer(0, 4) = IIf(source = "query", 1, 2): buffer(0, 5) = ISODateValue(item("tdlap"))
    identity = GdtInvoiceIdentity(CStr(buffer(0, 0)), CStr(buffer(0, 3)), CStr(buffer(0, 1)), CStr(buffer(0, 2)))
    buffer(0, 6) = targets(identity): buffer(0, 7) = GdtInvoiceStatus(SafeJsonText(item, "tthai"))
End Sub
Private Sub WriteDetail(ByRef buffer As Variant, ByVal invoiceType As Long, ByRef nextRow As Long)
    Dim start As Long
    start = nextRow
    ghiExcel_ChiTiet CoreDetailInvoiceJson(), nextRow, IIf(invoiceType = 1, "mua", "ban")
    CommitGdtDetailRows buffer, 0, start, nextRow
End Sub
Public Function CoreFixCase(ByVal mode As String, ByVal invoiceType As Long) As String
    On Error GoTo Failed
    Initialize invoiceType
    BeginGdtResume "CORE-ACCOUNT", invoiceType, True
    Dim page As Object, second As Object, targets As Object, buffer As Variant, r As Long, dr As Long, seq As Long
    Dim ws As Worksheet, detailWs As Worksheet, errorNumber As Long, first As Long
    Set ws = SummarySheet(invoiceType)
    Set detailWs = DetailSheet(invoiceType)
    Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
    r = 3: dr = 3
    Select Case mode
    Case "text"
        Dim literal As Variant
        For Each literal In Array("=1+1", "+1", "-1", "@name", "'quote", "00042")
            GdtWriteTextCell ws.Cells(3, 9), literal
            Check Not ws.Cells(3, 9).HasFormula And CStr(ws.Cells(3, 9).Value2) = CStr(literal), "literal text round-trip " & literal
        Next literal
        page("datas")(1)("nbten") = "=1+1"
        page("datas")(1)("shdon") = "00042"
        WritePage page, "query", invoiceType, r, buffer
        Check Not ws.Cells(3, 9).HasFormula And CStr(ws.Cells(3, 9).Value2) = "=1+1", "summary text formula"
        Check CStr(ws.Cells(3, 5).Value2) = "00042", "invoice leading zeros"
        Dim detail As Object
        Set detail = JsonConverter.ParseJSON(CoreDetailInvoiceJson())
        detail("nbten") = "=1+1": detail("hdhhdvu")(1)("ten") = "=1+1"
        ghiExcel_ChiTiet JsonConverter.ConvertToJson(detail), dr, IIf(invoiceType = 1, "mua", "ban")
        Check Not detailWs.Cells(3, 7).HasFormula And Not detailWs.Cells(4, 7).HasFormula, "detail common text filldown"
        Check CStr(detailWs.Cells(3, 19).Value2) = "=1+1" And Not detailWs.Cells(3, 19).HasFormula, "item text formula"
        Check IsNumeric(detailWs.Cells(3, 27).Value2), "amount no longer numeric"
    Case "unknown"
        page("datas")(1)("tthai") = 99
        page("datas")(1)("ttxly") = "new-code"
        WritePage page, "query", invoiceType, r, buffer
        Check r = 4 And InStr(CStr(ws.Cells(3, 52).Value2), "99") > 0, "unknown status dropped invoice"
        Check buffer(0, 7) = -1, "unknown status routed to related API"
    Case "numbers"
        page("datas")(1)("tgtthue") = "10.5"
        WritePage page, "query", invoiceType, r, buffer
        Check ws.Cells(3, 44).Value2 = 10.5 And VarType(ws.Cells(3, 44).Value2) = vbDouble, "JSON decimal/locale"
        Check GdtJsonNumber("1.25e2", "fixture") = 125, "JSON exponent"
    Case "bad-number"
        page("datas")(1)("tgtthue") = "=1+1"
        On Error Resume Next
        WritePage page, "query", invoiceType, r, buffer
        errorNumber = Err.Number
        On Error GoTo Failed
        Check errorNumber <> 0 And r = 3 And Application.CountA(ws.Range("A3:BL3")) = 0, "numeric formula accepted or partial row"
    Case "metadata-conflict"
        EndGdtResume
        ws.Cells(2, 65).Value2 = "USER COLUMN"
        On Error Resume Next
        BeginGdtResume "CORE-ACCOUNT", invoiceType, True
        errorNumber = Err.Number
        On Error GoTo Failed
        ws.Cells(2, 65).Value2 = "Account scope"
        Check errorNumber <> 0 And Application.CountA(ws.Range("A3:BN3")) = 0, "reserved metadata overwrote user column"
    Case "legacy"
        EndGdtResume
        seq = 1
        ghiExcel_TongHop CoreSummaryDateJson(), r, invoiceType, seq
        ghiExcel_ChiTiet CoreDetailInvoiceJson(), dr, IIf(invoiceType = 1, "mua", "ban")
        BeginGdtResume IIf(invoiceType = 1, "0000000002", "0000000001"), invoiceType, True
        WritePage page, "query", invoiceType, r, buffer
        Check r = 4 And buffer(0, 6) = 3 And Not CanResumeGdtStage(buffer, 0, "DETAIL"), "legacy ownership/adoption"
        WriteDetail buffer, invoiceType, dr
        FinalizeGdtResumeDetails
        Check detailWs.Cells(detailWs.Rows.Count, 3).End(xlUp).Row = 4, "legacy successful replacement duplicated detail"
    Case "clear"
        WritePage page, "query", invoiceType, r, buffer
        WriteDetail buffer, invoiceType, dr
        Call EndGdtResume
        ws.Range("A3:BN2000").ClearContents: detailWs.Range("A3:AJ2000").ClearContents
        BeginGdtResume "CORE-ACCOUNT", invoiceType, False
        Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
        r = 3: dr = 3
        WritePage page, "query", invoiceType, r, buffer
        Check Not CanResumeGdtStage(buffer, 0, "DETAIL"), "fresh download reused old checkpoint"
        WriteDetail buffer, invoiceType, dr
        FinalizeGdtResumeDetails
        Check dr = 5 And CStr(detailWs.Cells(3, 3).Value2) = "42", "clear-mode deleted new detail"
    Case "dedup"
        Set second = JsonConverter.ParseJSON(CoreSummaryDateJson())("datas")(1)
        page("datas").Add second
        WritePage page, "query", invoiceType, r, buffer
        Check r = 4 And page("datas").Count = 1, "same-page duplicate"
        Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
        WritePage page, "query", invoiceType, r, buffer
        Check r = 4 And page("datas").Count = 0 And GdtDuplicatesSkipped = 2, "cross-page duplicate"
    Case "resume", "changed", "missing", "scope", "backend", "replace", "failed-replace"
        WritePage page, "query", invoiceType, r, buffer
        WriteDetail buffer, invoiceType, dr
        Check dr = 5, "first detail block"
        EndGdtResume
        BeginGdtResume IIf(mode = "scope", "OTHER-ACCOUNT", "CORE-ACCOUNT"), invoiceType, True
        If mode = "missing" Then
            detailWs.Rows(4).Delete
            Call EndGdtResume: BeginGdtResume "CORE-ACCOUNT", invoiceType, True
        End If
        Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
        If mode = "changed" Or mode = "replace" Or mode = "failed-replace" Then page("datas")(1)("tgtthue") = 41
        WritePage page, IIf(mode = "backend", "sco-query", "query"), invoiceType, r, buffer
        If mode = "scope" Or mode = "backend" Then
            Check r = 5 And Not CanResumeGdtStage(buffer, 0, "DETAIL"), "account/source reused foreign state"
        Else
            Check r = 4 And buffer(0, 6) = 3, "resume appended summary or wrong target"
            Check CanResumeGdtStage(buffer, 0, "DETAIL") = (mode = "resume"), "stage proof/signature/missing rows"
        End If
        If mode = "replace" Then
            WriteDetail buffer, invoiceType, dr
            FinalizeGdtResumeDetails
            Check detailWs.Cells(detailWs.Rows.Count, 3).End(xlUp).Row = 4, "replacement duplicated detail"
            Call EndGdtResume: BeginGdtResume "CORE-ACCOUNT", invoiceType, True
            Set page = JsonConverter.ParseJSON(CoreSummaryDateJson()): page("datas")(1)("tgtthue") = 41
            WritePage page, "query", invoiceType, r, buffer
            Check CanResumeGdtStage(buffer, 0, "DETAIL"), "replacement state not reusable"
        ElseIf mode = "failed-replace" Then
            first = dr
            On Error Resume Next
            ghiExcel_ChiTiet "{bad", dr, IIf(invoiceType = 1, "mua", "ban")
            errorNumber = Err.Number
            On Error GoTo Failed
            FinalizeGdtResumeDetails
            Check errorNumber <> 0 And dr = first And CStr(detailWs.Cells(3, 3).Value2) = "42", "failed replacement destroyed old detail"
        End If
    Case "rollback"
        WritePage page, "query", invoiceType, r, buffer
        ws.Cells(3, 64).Value2 = "KEEP"
        GdtWriteTextCell ws.Cells(3, 9), "=1+1"
        ws.Cells(3, 63).Formula = "=1+1"
        ws.Cells(3, 15).NumberFormat = "0.00": ws.Cells(3, 15).Formula = "=1+1"
        ws.Cells(3, 62).Value2 = CVErr(xlErrNA)
        Call EndGdtResume: BeginGdtResume "CORE-ACCOUNT", invoiceType, True
        Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
        Set second = JsonConverter.ParseJSON(CoreSummaryDateJson())("datas")(1)
        second("shdon") = 43: second("tdlap") = "2026-02-30": page("datas").Add second
        PrepareGdtResumePage page, "query", targets
        seq = 1
        On Error Resume Next
        ghiExcel_TongHop "", r, invoiceType, seq, page, targets
        errorNumber = Err.Number
        On Error GoTo Failed
        Check errorNumber <> 0 And r = 4 And CStr(ws.Cells(3, 64).Value2) = "KEEP", "upsert transaction rollback"
        Check Application.CountA(ws.Range("A4:BL4")) = 0, "rollback leaked appended row"
        Check Not ws.Cells(3, 9).HasFormula And CStr(ws.Cells(3, 9).Value2) = "=1+1", "rollback converted literal text into formula"
        Check ws.Cells(3, 63).HasFormula And ws.Cells(3, 63).Value2 = 2, "rollback lost original formula"
        Check ws.Cells(3, 15).HasFormula And ws.Cells(3, 15).Value2 = 2 And ws.Cells(3, 15).NumberFormat = "0.00", "rollback lost formula format in text column"
        Check IsError(ws.Cells(3, 62).Value2), "rollback lost existing Excel error value"
    Case "sort"
        Set second = JsonConverter.ParseJSON(CoreSummaryDateJson())("datas")(1)
        second("shdon") = 43: second("tdlap") = "2026-09-10": page("datas").Add second
        WritePage page, "query", invoiceType, r, buffer
        WriteDetail buffer, invoiceType, dr
        FinalizeInvoiceSheets invoiceType
        Call EndGdtResume: BeginGdtResume "CORE-ACCOUNT", invoiceType, True
        Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
        WritePage page, "query", invoiceType, r, buffer
        Check r = 5 And buffer(0, 6) = 4 And CanResumeGdtStage(buffer, 0, "DETAIL"), "sort lost ownership/state/target row"
    End Select
    EndGdtResume
    CoreFixCase = "PASS"
    Exit Function
Failed:
    CoreFixCase = "FAIL: " & Err.Number & ": " & Err.Description
    EndGdtResume
End Function
Public Function CoreRetryDates() As String
    On Error GoTo Failed
    Dim utc As Date, raw As Variant
    utc = DateSerial(2026, 10, 13) + TimeSerial(12, 0, 0)
    Check ParseGdtRetryAfter("120", utc) = 120, "seconds header"
    Check ParseGdtRetryAfter(String$(20, "0") & "120", utc) = 120, "seconds leading zeros"
    Check ParseGdtRetryAfter("Wed, 14 Oct 2026 12:00:00 GMT", utc) = 86400, "GMT header"
    Check ParseGdtRetryAfter("Tue, 13 Oct 2026 12:00:01 GMT", utc) = 1, "rounding"
    Check ParseGdtRetryAfter("Mon, 12 Oct 2026 12:00:00 GMT", utc) = 0, "past date"
    Check ParseGdtRetryAfter("999999999999", utc) = -1, "overflow sentinel"
    For Each raw In Array("", "-1", "1.5", "bad", "Tue, 31 Feb 2026 12:00:00 GMT", "Tue, 13 Oct 2026 24:00:00 GMT")
        Check ParseGdtRetryAfter(CStr(raw), utc) = 0, "malformed header " & raw
    Next raw
    Check ISODateValue("2026-12-31T23:00:00Z") = DateSerial(2027, 1, 1), "UTC year rollover"
    Check ISODateValue("2024-02-29") = DateSerial(2024, 2, 29), "leap day"
    Check GdtElapsedText(90061) = "25:01:01" And GdtElapsedText(86401) = "24:00:01", "elapsed time wrapped after midnight/24 hours"
    CoreRetryDates = "PASS"
    Exit Function
Failed:
    CoreRetryDates = "FAIL: " & Err.Description
End Function
Public Function CoreLog() As String
    CoreLog = frmTaiHoaDon.CoreTestLog()
End Function
Public Function CoreXmlDestination(ByVal folder As String) As String
    CoreXmlDestination = frmTaiHoaDon.CoreTestFolder(folder)
End Function
Public Function CoreStateConflict() As String
    On Error GoTo Failed
    Dim ws As Worksheet, errorNumber As Long
    Set ws = ThisWorkbook.Sheets("TrangThaiTaiHD")
    ws.Cells(1, 1).Value2 = "USER SHEET"
    On Error Resume Next
    BeginGdtResume "CORE-ACCOUNT", 1, True
    errorNumber = Err.Number
    On Error GoTo Failed
    ws.Cells(1, 1).Value2 = "Identity (v1)"
    Check errorNumber <> 0, "unrecognized state schema accepted"
    EndGdtResume
    CoreStateConflict = "PASS"
    Exit Function
Failed:
    CoreStateConflict = "FAIL: " & Err.Description
End Function
Public Function CoreSchedulerSkip() As String
    On Error GoTo Failed
    Initialize 1
    BeginGdtResume "CORE-ACCOUNT", 1, True
    Dim page As Object, buffer As Variant, r As Long, dr As Long
    Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
    r = 3: dr = 3
    WritePage page, "query", 1, r, buffer
    WriteDetail buffer, 1, dr
    Check frmTaiHoaDon.CoreTestSkip(buffer), "completed detail scheduled another request"
    EndGdtResume
    CoreSchedulerSkip = "PASS"
    Exit Function
Failed:
    CoreSchedulerSkip = "FAIL: " & Err.Description
    EndGdtResume
End Function
Public Function CoreXmlFixture(ByVal folder As String, ByVal invoiceType As Long) As String
    On Error GoTo Failed
    Initialize invoiceType
    BeginGdtResume "CORE-ACCOUNT", invoiceType, True
    Dim page As Object, buffer As Variant, r As Long, fso As Object, stream As Object, base As String
    Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
    r = 3: WritePage page, "query", invoiceType, r, buffer
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = fso.BuildPath(folder, "0000000001_1_C26TST_42")
    MarkGdtStageComplete buffer, 0, "XML", 0, folder, folder
    Check CanResumeGdtStage(buffer, 0, "XML", folder, folder), "valid ZIP/XML/HTML proof"
    fso.DeleteFile base & ".html"
    Check Not CanResumeGdtStage(buffer, 0, "XML", folder, folder), "missing HTML skipped"
    Set stream = fso.CreateTextFile(base & ".html", True, False): stream.Write "<html>fixture</html>": stream.Close
    Set stream = fso.CreateTextFile(base & ".xml", True, False): stream.Write "<HDon??": stream.Close
    Check Not CanResumeGdtStage(buffer, 0, "XML", folder, folder), "invalid XML skipped"
    Set stream = fso.CreateTextFile(base & ".xml", True, False): stream.Write "<HDon/>": stream.Close
    Check Not CanResumeGdtStage(buffer, 0, "XML", folder & "other", folder), "folder change reused proof"
    Check CanResumeGdtStage(buffer, 0, "XML", folder, folder), "restored valid files not reusable"
    EndGdtResume
    CoreXmlFixture = "PASS"
    Exit Function
Failed:
    CoreXmlFixture = "FAIL: " & Err.Description
    EndGdtResume
End Function
Public Function CoreErrorCounts() As String
    On Error GoTo Failed
    Dim ws As Worksheet
    Set ws = EnsureGdtErrorReportSheet()
    ws.Range("A2:S2000").ClearContents
    GdtReportAccount = "CORE-ACCOUNT": GdtReportSession = "RUN-1"
    UpsertGdtErrorReport "purchase", "query", "0000000001", "1", "FIXTURE", "1", Empty, "DETAIL", "fixture", 500, "fixture", 5, 0, "Failed"
    GdtReportSession = "RUN-2"
    UpsertGdtErrorReport "purchase", "query", "0000000001", "1", "FIXTURE", "2", Empty, "DETAIL", "fixture", 500, "fixture", 5, 0, "Failed"
    GdtReportAccount = "OTHER-ACCOUNT"
    UpsertGdtErrorReport "purchase", "query", "0000000001", "1", "FIXTURE", "3", Empty, "DETAIL", "fixture", 500, "fixture", 5, 0, "Failed"
    GdtReportAccount = "CORE-ACCOUNT"
    Check CountGdtSessionErrors("purchase") = 1 And CountGdtSessionErrors("sales") = 0, "account/session/direction error counts"
    MarkGdtRetrySuccess "purchase", "query", "0000000001", "1", "FIXTURE", "2", "DETAIL", 6
    Check CountGdtSessionErrors("purchase") = 0 And ws.Cells(2, 1).Value2 = 1 And ws.Cells(3, 1).Value2 = 2, "retry cleanup/renumber"
    CoreErrorCounts = "PASS"
    Exit Function
Failed:
    CoreErrorCounts = "FAIL: " & Err.Description
End Function
Public Function CorePersist(ByVal prepare As Boolean) As String
    On Error GoTo Failed
    Dim page As Object, buffer As Variant, r As Long, dr As Long
    If prepare Then Initialize 1
    BeginGdtResume "CORE-ACCOUNT", 1, True
    Set page = JsonConverter.ParseJSON(CoreSummaryDateJson())
    r = IIf(prepare, 3, 4): dr = 3
    WritePage page, "query", 1, r, buffer
    If prepare Then WriteDetail buffer, 1, dr
    If Not prepare Then Check r = 4 And CanResumeGdtStage(buffer, 0, "DETAIL"), "saved/reopened state lost"
    EndGdtResume
    CorePersist = "PASS"
    Exit Function
Failed:
    CorePersist = "FAIL: " & Err.Description
End Function
'@)
    $formCode = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    # This fixture changes the XML checkbox programmatically. Keep its real
    # state updates, but print the informational prompt instead of blocking
    # unattended Excel on a modal dialog. The workbook is never saved.
    $xmlClickStart = $formCode.ProcStartLine('chkXmlZip_Click', 0)
    $xmlClickCount = $formCode.ProcCountLines('chkXmlZip_Click', 0)
    $xmlClickText = $formCode.Lines($xmlClickStart, $xmlClickCount)
    if ($xmlClickText -notmatch '(?m)^\s*MsgBoxUni ') { throw 'XML checkbox information prompt not found.' }
    $formCode.DeleteLines($xmlClickStart, $xmlClickCount)
    $formCode.InsertLines($xmlClickStart, ($xmlClickText -replace '(?m)^(\s*)MsgBoxUni ', '$1Debug.Print '))
    $formCode.AddFromString(@'
Public Function CoreTestLog() As String
    On Error GoTo Failed
    EnsureRuntimeControls
    Dim i As Long, text As String, parsed As Object, seen As Object, nextState As String
    Me.Controls("txtSimpleLog").Value = ""
    For i = 1 To 500
        AppendSimpleLog "DETAIL | HTTP 200 | attempts=1 | complete line " & i & String$(30, "x")
    Next i
    text = CStr(Me.Controls("txtSimpleLog").Value)
    If Len(text) > 12000 Or Not Left$(text, 10) Like "##/##/####" Or InStr(text, "500") = 0 Then Err.Raise 5, , "log boundary/timestamp/tail"
    Set parsed = JsonConverter.ParseJSON("{""state"":""repeat""}")
    Set seen = CreateObject("Scripting.Dictionary"): seen("repeat") = True
    mPaginationIncomplete = False
    If TryRegisterNextListState(parsed, seen, "query", nextState) Then Err.Raise 5, , "repeated state accepted"
    If Not mPaginationIncomplete Then Err.Raise 5, , "pagination warning not propagated"
    CoreTestLog = "PASS"
    Exit Function
Failed:
    CoreTestLog = "FAIL: " & Err.Description
End Function
Public Function CoreTestSkip(ByRef buffer As Variant) As Boolean
    EnsureRuntimeControls
    Dim index As Long, part As Long, task As clsGdtXmlTask
    mInvoiceWorkTotal = 1: mInvoiceWorkDone = 0
    Set task = NextJsonInvoiceTask(buffer, 1, 1, True, index, part)
    CoreTestSkip = (task Is Nothing And index = 1 And mInvoiceWorkDone = 1)
End Function
Public Function CoreTestFolder(ByVal folder As String) As String
    On Error GoTo Failed
    Dim errorNumber As Long
    Me.chkXmlZip.Value = True: Me.txtXMLFolderPath.Value = folder
    oSaveUnzipFolder = ""
    PrepareXmlDestination
    If CStr(oSaveUnzipFolder) <> folder Then Err.Raise 5, , "XML extraction destination not initialized"
    Me.txtXMLFolderPath.Value = folder & "\missing"
    On Error Resume Next
    PrepareXmlDestination
    errorNumber = Err.Number
    On Error GoTo Failed
    If errorNumber <> 76 Or CStr(oSaveUnzipFolder) <> folder Then Err.Raise 5, , "missing XML destination accepted"
    Me.chkXmlZip.Value = False
    PrepareXmlDestination
    Me.txtXMLFolderPath.Value = folder
    CoreTestFolder = "PASS"
    Exit Function
Failed:
    CoreTestFolder = "FAIL: " & Err.Description
End Function
'@)
    foreach ($direction in @(1,2)) {
        foreach ($mode in @('text','unknown','numbers','bad-number','metadata-conflict','legacy','clear','dedup','resume','changed','missing','scope','backend','replace','failed-replace','rollback','sort')) {
            $name = "$mode-$direction"
            $results[$name] = [string]$excel.Run("${prefix}CoreFixCase", [string]$mode, [int]$direction)
            Write-Output "$name=$($results[$name])"
            if ($results[$name] -ne 'PASS') { throw "Core regression failed: $name=$($results[$name])" }
        }
    }
    foreach ($direction in @(1,2)) {
        $results["xml-$direction"] = [string]$excel.Run("${prefix}CoreXmlFixture", [string]$testRoot, [int]$direction)
        Write-Output "xml-$direction=$($results["xml-$direction"])"
        if ($results["xml-$direction"] -ne 'PASS') { throw 'XML resume proof failed.' }
    }
    $results['CoreXmlDestination'] = [string]$excel.Run("${prefix}CoreXmlDestination", [string]$testRoot)
    Write-Output "CoreXmlDestination=$($results['CoreXmlDestination'])"
    if ($results['CoreXmlDestination'] -ne 'PASS') { throw 'XML destination initialization failed.' }
    foreach ($entry in @('CoreRetryDates','CoreLog','CoreErrorCounts','CoreSchedulerSkip','CoreStateConflict')) {
        $results[$entry] = [string]$excel.Run("${prefix}$entry")
        Write-Output "$entry=$($results[$entry])"
        if ($results[$entry] -ne 'PASS') { throw "Core regression failed: $entry=$($results[$entry])" }
    }
    $results['persist-prepare'] = [string]$excel.Run("${prefix}CorePersist", [bool]$true)
    if ($results['persist-prepare'] -ne 'PASS') { throw 'Persistence preparation failed.' }
    $book.SaveCopyAs($reopenPath)
    Release-ComObject $formCode; $formCode = $null
    Release-ComObject $code; $code = $null
    Release-ComObject $component; $component = $null
    $book.Close($false); Release-ComObject $book; $book = $null
    $book = $excel.Workbooks.Open($reopenPath, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $results['persist-reopen'] = [string]$excel.Run("${prefix}CorePersist", [bool]$false)
    Write-Output "persist-reopen=$($results['persist-reopen'])"
    if ($results['persist-reopen'] -ne 'PASS') { throw 'Resume failed after save/reopen.' }
    $results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $BuiltWorkbook) 'core-fixes-runtime-result.json') -Encoding UTF8
} catch {
    try {
        $pane = $excel.VBE.ActiveCodePane
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $pane.GetSelection([ref]$startLine,[ref]$startColumn,[ref]$endLine,[ref]$endColumn)
        Write-Output ("Compile location=" + $pane.CodeModule.Parent.Name + ':' + $startLine)
        Write-Output $pane.CodeModule.Lines($startLine,1)
        Release-ComObject $pane
    } catch {}
    throw
} finally {
    Release-ComObject $formCode; Release-ComObject $code; Release-ComObject $component
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -match '^hddt-core-fixes-[a-f0-9]{32}$') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Core fixture modified the built artifact.' }
Write-Output 'CORE FIXES RUNTIME TEST PASSED'
