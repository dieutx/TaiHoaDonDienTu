VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTaiHoaDon 
   Caption         =   "+++"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15930
   OleObjectBlob   =   "frmTaiHoaDon.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTaiHoaDon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim url As String, url_sco As String, url2 As String, url_ct As String, res As String, errMsg As String
Dim sort As String, size As Long, search As String, search_ex As String, tthai As String, ttxly As String
Dim arrDate(), arrHDChiTiet(), k As Long, i As Long, blnNgay As Boolean
Dim js As Object, jsErr As Object
Dim mUiHandlers As Collection
Dim mFinalRetryQueue As Collection
Dim mFinalRetryCooldownDone As Boolean
Dim mInvoiceWorkTotal As Long
Dim mInvoiceWorkDone As Long
Dim mDownloadRunning As Boolean
Dim mDownloadEntering As Boolean
Dim mControlStates As Object
Dim mRunToken As String
Dim mRunPurchase As Boolean

Private Function RunToken() As String
    If mDownloadRunning Then
        RunToken = mRunToken
    Else
        RunToken = getToken()
    End If
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mDownloadRunning Or mDownloadEntering Then
        Cancel = 1
        RequestStopDownload
    End If
End Sub

Private Sub cboDonVi_Change()
    Me.lblTenDV.caption = Me.cboDonVi.Column(2)
    cToken = Me.cboDonVi.Column(3)
End Sub

Private Sub chkXmlZip_Click()
    Me.txtXMLFolderPath.Enabled = Me.chkXmlZip
    Me.cmdChonFolder.Enabled = Me.chkXmlZip
    Me.cmdChiTietXML.Enabled = Me.chkXmlZip
    If Me.chkXmlZip = True Then
        MsgBoxUni "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i ch" & ChrW(7885) & "n " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng d" & ChrW(7851) & _
                "n th" & ChrW(432) & " m" & ChrW(7909) & "c " & ChrW(273) & ChrW(7875) & " l" & ChrW(432) & "u file zip ho" & ChrW(7863) & "c d" & ChrW(249) & _
                "ng " & ChrW(273) & ChrW(432) & ChrW(7903) & "ng d" & ChrW(7851) & "n m" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh.", vbInformation, UniConvert("Thoong baso")
    Me.txtXMLFolderPath.SetFocus
    End If
End Sub

Private Sub cmdChiTietXML_Click()
    frmTrichXuatXML.Show
End Sub

Private Sub cmdChonFolder_Click()
    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .Title = "Chon thu muc de luu file Zip XML"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
        Me.txtXMLFolderPath = .SelectedItems(1) & "\"
    End With
End Sub

Private Sub cmdPickTuNgay_Click()
    OpenStartDatePicker
End Sub

Private Sub cmdPickDenNgay_Click()
    OpenEndDatePicker
End Sub

Public Sub OpenStartDatePicker()
    ShowDatePicker Me.txtTuNgay
End Sub

Public Sub OpenEndDatePicker()
    ShowDatePicker Me.txtDenNgay
End Sub

Private Sub ShowDatePicker(ByVal targetBox As Object)
    frmDatePicker.SetTarget targetBox
    frmDatePicker.Show vbModal
End Sub

Public Sub RefreshPickedDate(ByVal targetName As String)
    If targetName = "txtTuNgay" Then
        txtTuNgay_AfterUpdate
    ElseIf targetName = "txtDenNgay" Then
        txtDenNgay_AfterUpdate
    End If
End Sub

Private Sub cmdDangXuat_Click()
    Unload Me
    frmDangNhap.Show
End Sub

Private Sub txtTuNgay_AfterUpdate()
    Dim blnErr As Boolean
    Dim inputText As String, parsedDate As Date
    inputText = CStr(Me.txtTuNgay.Value)
    parsedDate = correctDate(inputText, blnErr)
    Me.lblSaiNgay.Visible = blnErr
    If blnErr Then Exit Sub
    Me.txtTuNgay.Value = Format$(parsedDate, "dd/mm/yyyy")
End Sub

Private Sub txtTuNgay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtDenNgay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtDenNgay_AfterUpdate()
    Dim blnErr As Boolean
    Dim startText As String, endText As String
    Dim startDate As Date, endDate As Date
    endText = CStr(Me.txtDenNgay.Value)
    endDate = correctDate(endText, blnErr)
    Me.lblSaiNgay.Visible = blnErr
    If blnErr Then Exit Sub
    Me.txtDenNgay.Value = Format$(endDate, "dd/mm/yyyy")
    'Khong so sanh duoc khi mot trong hai o ngay con trong
    If Len(Trim(Me.txtTuNgay)) = 0 Or Len(Trim(Me.txtDenNgay)) = 0 Then Exit Sub
    startText = CStr(Me.txtTuNgay.Value)
    endText = CStr(Me.txtDenNgay.Value)
    endDate = correctDate(endText, blnErr)
    If blnErr Then Exit Sub
    startDate = correctDate(startText, blnErr)
    If blnErr Then Exit Sub
    If endDate < startDate Then
        MsgBoxUni "Sai ng" & ChrW(224) & "y. [Ng" & ChrW(224) & "y b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u] > [Ng" & ChrW(224) & "y k" & ChrW(7871) & "t th" & ChrW(250) & "c].", vbCritical
        Exit Sub
    End If
    Call lietKeThoiGian
End Sub

Private Sub AbortInvalidDateRange()
    MsgBoxUni "Ki" & ChrW(7875) & "m tra l" & ChrW(7841) & "i ng" & ChrW(224) & "y t" & ChrW(236) & "m ki" & ChrW(7871) & "m.", vbCritical, UniConvert("Thoong baso")
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function correctDate(ByRef textDate As String, ByRef err As Boolean) As Date
    Dim ngay As Long, thang As Long, nam As Long
    Dim dtNgayNhap As Date
    err = False
    On Error GoTo InvalidDate
    
    If Len(textDate) < 10 Or Len(textDate) > 10 Then
        correctDate = Date
        err = True
        Exit Function
    ElseIf Mid(textDate, 3, 1) <> "/" Or Mid(textDate, 6, 1) <> "/" Then
        correctDate = Date
        err = True
        Exit Function
    End If
    
    If Not (Left$(textDate, 2) Like "##") Or Not (Mid$(textDate, 4, 2) Like "##") Or _
        Not (Right$(textDate, 4) Like "####") Then GoTo InvalidDate
    ngay = CLng(Left$(textDate, 2))
    thang = CLng(Mid$(textDate, 4, 2))
    nam = CLng(Right$(textDate, 4))
    If thang < 1 Or thang > 12 Or ngay < 1 Or ngay > 31 Then GoTo InvalidDate

    dtNgayNhap = DateSerial(nam, thang, ngay)
    If Day(dtNgayNhap) <> ngay Or Month(dtNgayNhap) <> thang Or Year(dtNgayNhap) <> nam Then GoTo InvalidDate
    correctDate = dtNgayNhap
    Exit Function

InvalidDate:
    correctDate = Date
    err = True
    
End Function

Private Sub UserForm_Initialize()
    EnsureRuntimeControls
    ' Load listbox chon Donvi
    Dim arrDonVi()
    arrDonVi = ThisWorkbook.Sheets("MENU").Range("A7:D" & ThisWorkbook.Sheets("MENU").Cells(Rows.count, "A").End(xlUp).row).Value
    Me.cboDonVi.list = arrDonVi
    
    Call optMua_Change
    Me.cboKQKT.ListIndex = 0
    Me.cboTTHD.ListIndex = 0
    Me.lblStatus.Visible = False
    Me.lblSaiNgay.Visible = False
    Me.txtXMLFolderPath = Environ("USERPROFILE") & "\Documents\"
    Me.txtXMLFolderPath.Enabled = False
    Me.cmdChonFolder.Enabled = False
    Me.chkTH.Value = True
    Me.chkCT.Value = True
    ResetProgressUI
    Set mFinalRetryQueue = New Collection
    ResetGdtRequestSession

End Sub

Private Function CurrentDirectionName() As String
    If IIf(mDownloadRunning, mRunPurchase, Me.optMua.Value) Then
        CurrentDirectionName = UniConvert("Mua vafo")
    Else
        CurrentDirectionName = UniConvert("Basn ra")
    End If
End Function

Private Sub QueueFinalRetry( _
    ByVal apiSource As String, _
    ByVal sellerTaxCode As String, _
    ByVal templateCode As String, _
    ByVal invoiceSeries As String, _
    ByVal invoiceNumber As String, _
    ByVal invoiceDate As Variant, _
    ByVal stageName As String, _
    ByVal endpoint As String, _
    ByVal responseKind As String, _
    Optional ByVal targetRow As Long = 0, _
    Optional ByVal actionHeader As String = vbNullString)

    Dim item As clsGdtRetryItem
    Dim existing As clsGdtRetryItem
    Dim finalResult As String

    If GdtStopRequested Then Exit Sub

    If LastGdtAuthFailed Then
        finalResult = UniConvert("Token heest hajn")
    ElseIf Not LastGdtShouldQueue Then
        finalResult = UniConvert("Khoong tari dduwowjc")
    Else
        finalResult = UniConvert("Chowf thuwr laji cuoosi phieen")
    End If

    UpsertGdtErrorReport CurrentDirectionName(), apiSource, sellerTaxCode, _
        templateCode, invoiceSeries, invoiceNumber, invoiceDate, stageName, _
        endpoint, LastGdtStatus, LastGdtError, LastGdtAttempts, _
        LastGdtRetryAfter, finalResult

    If Not LastGdtShouldQueue Or LastGdtAuthFailed Then Exit Sub
    If mFinalRetryQueue Is Nothing Then Set mFinalRetryQueue = New Collection
    Set item = New clsGdtRetryItem
    item.Direction = CurrentDirectionName()
    item.ApiSource = apiSource
    item.SellerTaxCode = sellerTaxCode
    item.TemplateCode = templateCode
    item.InvoiceSeries = invoiceSeries
    item.InvoiceNumber = invoiceNumber
    item.InvoiceDate = invoiceDate
    item.Stage = stageName
    item.Endpoint = endpoint
    item.ResponseKind = responseKind
    item.Attempts = LastGdtAttempts
    item.TargetRow = targetRow
    item.ActionHeader = actionHeader

    For Each existing In mFinalRetryQueue
        If existing.UniqueKey = item.UniqueKey Then Exit Sub
    Next existing
    mFinalRetryQueue.Add item
    AppendSimpleLog UniConvert("DDax dduwa vafo hafng ddowji cuoosi phieen. Soos mujc: ") & mFinalRetryQueue.Count
End Sub

Private Sub EnsureFinalRetryCooldown()
    If GdtStopRequested Then Exit Sub
    If mFinalRetryCooldownDone Or mFinalRetryQueue Is Nothing Then Exit Sub
    If mFinalRetryQueue.Count = 0 Then Exit Sub
    ReportRetryProgress UniConvert("Chowf ") & GDT_FINAL_RETRY_COOLDOWN_SECONDS & UniConvert(" giaay ddeer thuwr laji cuoosi phieen...")
    WaitForFinalGdtRetry
    mFinalRetryCooldownDone = True
End Sub

Private Function TryParseGdtJson( _
    ByVal responseText As String, _
    ByRef parsed As Object, _
    ByVal apiSource As String, _
    ByVal endpoint As String, _
    Optional ByVal sellerTaxCode As String = vbNullString, _
    Optional ByVal templateCode As String = vbNullString, _
    Optional ByVal invoiceSeries As String = vbNullString, _
    Optional ByVal invoiceNumber As String = vbNullString, _
    Optional ByVal invoiceDate As Variant) As Boolean

    On Error GoTo ParseFailed
    Set parsed = JsonConverter.ParseJSON(responseText)
    If GdtJsonHasError(parsed) Then Err.Raise 5, , "API error response"
    Dim arrayData As Collection
    If InStr(1, endpoint, "/detail?", vbTextCompare) > 0 Then
        If TypeName(parsed) <> "Dictionary" Then Err.Raise 5, , "Expected invoice object"
        Set arrayData = GdtJsonArray(parsed, "hdhhdvu")
        If arrayData.count = 0 Then Err.Raise 5, , "Missing invoice items"
    ElseIf InStr(1, endpoint, "/purchase?", vbTextCompare) > 0 Or _
        InStr(1, endpoint, "/sold?", vbTextCompare) > 0 Then
        If TypeName(parsed) <> "Dictionary" Then Err.Raise 5, , "Expected list object"
        If Not parsed.Exists("datas") Then Err.Raise 5, , "Missing invoice list"
        If Not IsObject(parsed("datas")) Then Err.Raise 5, , "Invalid invoice list"
        Set arrayData = GdtJsonArray(parsed, "datas")
        Dim invoice As Variant, invoiceDateValue As Date, identityKey As Variant
        For Each invoice In arrayData
            If TypeName(invoice) <> "Dictionary" Then Err.Raise 5, , "Invalid invoice in list"
            For Each identityKey In Array("nbmst", "khmshdon", "khhdon", "shdon", "tdlap")
                If Not GdtJsonHasValue(invoice, CStr(identityKey)) Then Err.Raise 5, , "Missing invoice identity: " & identityKey
            Next identityKey
            invoiceDateValue = ISODateValue(invoice("tdlap"))
        Next invoice
    ElseIf InStr(1, endpoint, "/relative?", vbTextCompare) > 0 Then
        If TypeName(parsed) <> "Collection" Then Err.Raise 5, , "Expected related invoice array"
    End If
    TryParseGdtJson = True
    Exit Function

ParseFailed:
    UpsertGdtErrorReport CurrentDirectionName(), apiSource, sellerTaxCode, templateCode, _
        invoiceSeries, invoiceNumber, invoiceDate, "Parse " & UniConvert("duwx lieeju"), endpoint, _
        LastGdtStatus, Err.Description, LastGdtAttempts, LastGdtRetryAfter, UniConvert("Khoong tari dduwowjc")
    AppendSimpleLog UniConvert("Looxi ") & "parse " & UniConvert("duwx lieeju: ") & apiSource
    Err.Clear
End Function

Private Function NextStateEndpoint(ByVal endpoint As String, ByVal stateValue As String) As String
    Dim statePos As Long, searchPos As Long
    statePos = InStr(1, endpoint, "&state=", vbTextCompare)
    searchPos = InStr(1, endpoint, "&search=", vbTextCompare)
    If statePos > 0 And searchPos > statePos Then
        NextStateEndpoint = Left$(endpoint, statePos - 1) & "&state=" & stateValue & Mid$(endpoint, searchPos)
    ElseIf searchPos > 0 Then
        NextStateEndpoint = Left$(endpoint, searchPos - 1) & "&state=" & stateValue & Mid$(endpoint, searchPos)
    Else
        NextStateEndpoint = endpoint & "&state=" & stateValue
    End If
End Function

Private Sub ProcessQueuedListRetries( _
    ByRef invoiceBuffer As Variant, _
    ByRef invoiceCount As Long, _
    ByRef totalRow As Long, _
    ByRef sequenceNumber As Long, _
    ByVal invoiceType As Long)

    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    Dim parsed As Object, dataItem As Object
    Dim currentEndpoint As String, hasMore As Boolean
    Dim batchStartRow As Long, pageIndex As Long
    Dim seenStates As Object, nextState As String
    Dim retryPageNumber As Long, retryCount As Long, retryStarted As Double

    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "LIST" Then
            currentEndpoint = queued.Endpoint
            Set seenStates = CreateObject("Scripting.Dictionary")
            SeedListPageState currentEndpoint, seenStates
            retryPageNumber = 1
            retryCount = 0
            Do
                If Not WaitForGdtControl() Then Exit Sub
                hasMore = False
                SetProgress 9, UniConvert("Thuwr laji danh sasch ") & queued.ApiSource & _
                    UniConvert(" - trang ") & retryPageNumber, True
                retryStarted = Timer
                Set requestResult = ExecuteGdtRequest("GET", currentEndpoint, RunToken(), _
                    vbNullString, "application/json", "application/json, text/plain, */*", False, 1, "LIST")
                PublishLastGdtResult requestResult
                If requestResult.Success Then
                    If Not TryParseGdtJson(requestResult.ResponseText, parsed, queued.ApiSource, currentEndpoint, _
                        queued.SellerTaxCode, queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate) Then Exit Do
                    batchStartRow = totalRow
                    If Not WriteSummaryPage(requestResult.ResponseText, totalRow, invoiceType, sequenceNumber, _
                        queued.ApiSource, currentEndpoint, queued.InvoiceDate) Then Exit Do
                    MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, _
                        "Parse " & UniConvert("duwx lieeju"), queued.Attempts + requestResult.Attempts, vbNullString, currentEndpoint
                    pageIndex = 0
                    For Each dataItem In parsed("datas")
                        pageIndex = pageIndex + 1
                        EnsureInvoiceBufferCapacity invoiceBuffer, invoiceCount
                        invoiceBuffer(invoiceCount, 0) = dataItem("nbmst") & vbNullString
                        invoiceBuffer(invoiceCount, 1) = dataItem("khhdon")
                        invoiceBuffer(invoiceCount, 2) = dataItem("shdon")
                        invoiceBuffer(invoiceCount, 3) = dataItem("khmshdon")
                        invoiceBuffer(invoiceCount, 4) = IIf(queued.ApiSource = "query", 1, 2)
                        invoiceBuffer(invoiceCount, 5) = ISODateValue(dataItem("tdlap"))
                        invoiceBuffer(invoiceCount, 6) = batchStartRow + pageIndex - 1
                        invoiceBuffer(invoiceCount, 7) = CLng(Val(dataItem("tthai") & vbNullString))
                        invoiceCount = invoiceCount + 1
                    Next dataItem
                    retryCount = retryCount + pageIndex
                    MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.Stage, _
                        queued.Attempts + requestResult.Attempts, UniConvert("Danh sasch ddax tari thafnh coong"), currentEndpoint
                    nextState = vbNullString
                    hasMore = TryRegisterNextListState(parsed, seenStates, queued.ApiSource, nextState)
                    AppendSimpleLog UniConvert("Thuwr laji danh sasch ") & queued.ApiSource & _
                        UniConvert(" - trang ") & retryPageNumber & ": HTTP " & requestResult.StatusCode & _
                        UniConvert(", nhaajn ") & pageIndex & UniConvert(" HDD, toorng ") & retryCount & _
                        ", " & Format(ElapsedTimerSeconds(retryStarted), "0.0") & UniConvert(" giaay")
                    If hasMore Then
                        retryPageNumber = retryPageNumber + 1
                        currentEndpoint = NextStateEndpoint(currentEndpoint, nextState)
                    End If
                Else
                    UpsertGdtErrorReport queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate, _
                        queued.Stage, currentEndpoint, requestResult.StatusCode, requestResult.ErrorMessage, _
                        queued.Attempts + requestResult.Attempts, requestResult.RetryAfterSeconds, _
                        IIf(requestResult.AuthenticationFailure, UniConvert("Token heest hajn"), UniConvert("Khoong tari dduwowjc"))
                End If
                WaitGdtMilliseconds GetSleepDelayMs()
            Loop While hasMore And Not GdtAuthenticationFailed And Not GdtStopRequested
        End If
    Next queued
End Sub

Private Function WriteSummaryPage(ByVal responseText As String, ByRef targetRow As Long, _
    ByVal invoiceType As Long, ByRef sequenceNumber As Long, ByVal apiSource As String, _
    ByVal endpoint As String, ByVal invoiceDate As Variant) As Boolean
    On Error GoTo Failed
    ghiExcel_TongHop responseText, targetRow, invoiceType, sequenceNumber
    WriteSummaryPage = True
    Exit Function
Failed:
    UpsertGdtErrorReport CurrentDirectionName(), apiSource, "", "", "", "", invoiceDate, _
        "Parse " & UniConvert("duwx lieeju"), endpoint, LastGdtStatus, Err.Description, LastGdtAttempts, 0, "Write failed"
    AppendSimpleLog "Summary page write failed: " & apiSource
End Function

Private Sub SeedListPageState(ByVal endpoint As String, ByVal seenStates As Object)
    Dim statePos As Long, endPos As Long, currentState As String
    statePos = InStr(1, endpoint, "&state=", vbTextCompare)
    If statePos = 0 Then statePos = InStr(1, endpoint, "?state=", vbTextCompare)
    If statePos = 0 Then Exit Sub
    statePos = statePos + 7
    endPos = InStr(statePos, endpoint, "&")
    If endPos = 0 Then endPos = Len(endpoint) + 1
    currentState = Mid$(endpoint, statePos, endPos - statePos)
    If Len(currentState) > 0 Then seenStates(currentState) = True
End Sub

Private Function CountRelatedApiCalls(ByRef invoiceBuffer As Variant, ByVal invoiceCount As Long) As Long
    Dim invoiceIndex As Long, invoiceStatus As Long
    For invoiceIndex = 0 To invoiceCount - 1
        invoiceStatus = CLng(Val(invoiceBuffer(invoiceIndex, 7) & vbNullString))
        If invoiceStatus >= 2 And invoiceStatus <= 5 Then
            CountRelatedApiCalls = CountRelatedApiCalls + 2
        ElseIf invoiceStatus = 6 Then
            CountRelatedApiCalls = CountRelatedApiCalls + 1
        End If
    Next invoiceIndex
End Function

Private Function BuildRelationEndpoint( _
    ByVal apiSource As String, _
    ByVal endpointName As String, _
    ByVal sellerTaxCode As String, _
    ByVal templateCode As String, _
    ByVal invoiceSeries As String, _
    ByVal invoiceNumber As String) As String

    BuildRelationEndpoint = "https://hoadondientu.gdt.gov.vn/api/" & apiSource & _
        "/invoices/" & endpointName & "?nbmst=" & sellerTaxCode & _
        "&khmshdon=" & templateCode & "&khhdon=" & invoiceSeries & "&shdon=" & invoiceNumber
End Function

Private Function BuildRelationActionHeader( _
    ByVal endpointName As String, _
    ByVal apiSource As String, _
    ByVal invoiceType As Long) As String

    Dim actionPrefix As String, contextText As String
    If endpointName = "relative" Then
        actionPrefix = "Xem%20h%C3%B3a%20%C4%91%C6%A1n%20li%C3%AAn%20quan%20"
    Else
        actionPrefix = "Xem%20th%C3%B4ng%20tin%20li%C3%AAn%20quan%20"
    End If

    If apiSource = "sco-query" Then
        If invoiceType = 1 Then
            contextText = "(h%C3%B3a%20%C4%91%C6%A1n%20m%C3%A1y%20t%C3%ADnh%20ti%E1%BB%81n%20mua%20v%C3%A0o)"
        Else
            contextText = "(h%C3%B3a%20%C4%91%C6%A1n%20m%C3%A1y%20t%C3%ADnh%20ti%E1%BB%81n%20b%C3%A1n%20ra)"
        End If
    Else
        If invoiceType = 1 Then
            contextText = "(h%C3%B3a%20%C4%91%C6%A1n%20mua%20v%C3%A0o)"
        Else
            contextText = "(h%C3%B3a%20%C4%91%C6%A1n%20b%C3%A1n%20ra)"
        End If
    End If
    BuildRelationActionHeader = actionPrefix & contextText
End Function

Private Sub ProcessRelatedInvoiceApis( _
    ByRef invoiceBuffer As Variant, _
    ByVal invoiceCount As Long, _
    ByVal invoiceType As Long)

    Dim unusedDetailRow As Long
    ProcessJsonInvoiceRequests invoiceBuffer, invoiceCount, invoiceType, False, unusedDetailRow
End Sub

Private Function NextJsonInvoiceTask(ByRef invoiceBuffer As Variant, ByVal invoiceCount As Long, _
    ByVal invoiceType As Long, ByVal detailsOnly As Boolean, ByRef invoiceIndex As Long, _
    ByRef relationPart As Long) As clsGdtXmlTask
    Dim invoiceStatus As Long, apiSource As String, endpointName As String, endpoint As String
    Dim actionHeader As String, task As clsGdtXmlTask
    Do While invoiceIndex < invoiceCount
        apiSource = IIf(invoiceBuffer(invoiceIndex, 4) = 1, "query", "sco-query")
        If detailsOnly Then
            endpointName = "detail"
        Else
            invoiceStatus = CLng(Val(invoiceBuffer(invoiceIndex, 7) & vbNullString))
            If invoiceStatus < 2 Or invoiceStatus > 6 Then
                invoiceIndex = invoiceIndex + 1
                GoTo NextCandidate
            End If
            If relationPart = 0 Then
                If invoiceStatus = 6 Then
                    WriteNoRelativeInvoiceData invoiceType, CLng(invoiceBuffer(invoiceIndex, 6))
                    endpointName = "related"
                Else
                    endpointName = "relative"
                End If
            Else
                endpointName = "related"
            End If
            actionHeader = BuildRelationActionHeader(endpointName, apiSource, invoiceType)
        End If
        endpoint = BuildRelationEndpoint(apiSource, endpointName, CStr(invoiceBuffer(invoiceIndex, 0)), _
            CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), CStr(invoiceBuffer(invoiceIndex, 2)))
        Set task = New clsGdtXmlTask
        task.Configure invoiceIndex, endpoint, RunToken(), GDT_MAX_RETRIES, GDT_HTTP_TIMEOUT_SECONDS, False, actionHeader
        task.ResponseKind = UCase$(endpointName)
        Set NextJsonInvoiceTask = task
        If detailsOnly Or endpointName = "related" Then
            invoiceIndex = invoiceIndex + 1
            relationPart = 0
        Else
            relationPart = 1
        End If
        Exit Function
NextCandidate:
    Loop
End Function

Private Sub ProcessJsonInvoiceRequests(ByRef invoiceBuffer As Variant, ByVal invoiceCount As Long, _
    ByVal invoiceType As Long, ByVal detailsOnly As Boolean, ByRef detailRow As Long)
    Dim tasks As Collection, task As clsGdtXmlTask, throttle As clsGdtXmlThrottle
    Dim nextIndex As Long, relationPart As Long, active As Long, pos As Long, completed As Long
    Dim phaseName As String, failureNumber As Long, failureText As String, previousLimit As Long
    On Error GoTo Failed
    Set tasks = New Collection
    Set throttle = New clsGdtXmlThrottle
    throttle.Configure GDT_JSON_CONCURRENCY, GetSleepDelayMs()
    phaseName = IIf(detailsOnly, "DETAIL", "RELATED")
    AppendSimpleLog phaseName & " scheduler: " & throttle.CurrentConcurrency & " connections, " & throttle.CurrentIntervalMs & " ms"
    Do While nextIndex < invoiceCount Or tasks.Count > 0
        If GdtStopRequested Or GdtAuthenticationFailed Then Exit Do
        If Not GdtPauseRequested Then
            Do While nextIndex < invoiceCount And tasks.Count < GDT_JSON_BUFFER_SIZE
                Set task = NextJsonInvoiceTask(invoiceBuffer, invoiceCount, invoiceType, detailsOnly, nextIndex, relationPart)
                If task Is Nothing Then Exit Do
                tasks.Add task
            Loop
        End If
        active = 0
        For Each task In tasks
            If task.Running Then active = active + 1
        Next task
        previousLimit = throttle.CurrentConcurrency
        For Each task In tasks
            If task.Running Then active = active - 1
            task.Tick throttle, active
            If task.Running Then active = active + 1
            If task.Result.AuthenticationFailure Then
                HandleJsonInvoiceResult task, invoiceBuffer, invoiceType, detailRow
                Exit For
            End If
            If GdtStopRequested Then Exit For
        Next task
        If GdtStopRequested Or GdtAuthenticationFailed Then Exit Do
        If previousLimit <> throttle.CurrentConcurrency Then
            AppendSimpleLog phaseName & " throttle: " & throttle.CurrentConcurrency & " connections, " & _
                throttle.CurrentIntervalMs & " ms; cooldown " & Format$(throttle.CooldownUntil - GdtClockSeconds(), "0.0") & "s"
        End If
        If detailsOnly Then
            ' Keep original invoice order, with at most 12 responses buffered.
            ' A slow retry does not prevent the other HTTP slots from working.
            Do While tasks.Count > 0
                Set task = tasks(1)
                If Not task.Finished Then Exit Do
                HandleJsonInvoiceResult task, invoiceBuffer, invoiceType, detailRow
                tasks.Remove 1
                completed = completed + 1
                If GdtStopRequested Or GdtAuthenticationFailed Then Exit Do
            Loop
        Else
            ' Related responses have an explicit summary row and can be written
            ' immediately even if an earlier invoice is still waiting/retrying.
            For pos = tasks.Count To 1 Step -1
                Set task = tasks(pos)
                If task.Finished Then
                    HandleJsonInvoiceResult task, invoiceBuffer, invoiceType, detailRow
                    tasks.Remove pos
                    completed = completed + 1
                End If
                If GdtStopRequested Or GdtAuthenticationFailed Then Exit For
            Next pos
        End If
        GdtPumpWait
    Loop
    For Each task In tasks
        task.Cancel
    Next task
    AppendSimpleLog phaseName & " scheduler completed: " & completed
    Exit Sub
Failed:
    failureNumber = Err.Number: failureText = Err.Description
    If Not tasks Is Nothing Then
        For Each task In tasks
            task.Cancel
        Next task
    End If
    Err.Raise failureNumber, "ProcessJsonInvoiceRequests", failureText
End Sub

Private Sub HandleJsonInvoiceResult(ByVal task As clsGdtXmlTask, ByRef invoiceBuffer As Variant, _
    ByVal invoiceType As Long, ByRef detailRow As Long)
    Dim invoiceIndex As Long, apiSource As String, requestResult As clsGdtRequestResult
    invoiceIndex = task.InvoiceIndex
    apiSource = IIf(invoiceBuffer(invoiceIndex, 4) = 1, "query", "sco-query")
    Set requestResult = task.Result
    If task.ResponseKind <> "DETAIL" Then
        ProcessOneRelationRequest invoiceBuffer, invoiceIndex, invoiceType, apiSource, LCase$(task.ResponseKind), requestResult
        Exit Sub
    End If
    PublishLastGdtResult requestResult
    If requestResult.Success Then
        If WriteDetailRecord(requestResult.ResponseText, detailRow, apiSource, task.Endpoint, _
            CStr(invoiceBuffer(invoiceIndex, 0)), CStr(invoiceBuffer(invoiceIndex, 3)), _
            CStr(invoiceBuffer(invoiceIndex, 1)), CStr(invoiceBuffer(invoiceIndex, 2)), invoiceBuffer(invoiceIndex, 5)) Then
            MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
                CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
                CStr(invoiceBuffer(invoiceIndex, 2)), UniConvert("Laasy chi tieest"), requestResult.Attempts
            MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
                CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
                CStr(invoiceBuffer(invoiceIndex, 2)), "Parse " & UniConvert("duwx lieeju"), requestResult.Attempts
        End If
    ElseIf Not GdtStopRequested Then
        errMsg = errMsg & UniConvert("Looxi tari hoas ddown chi tieest: ") & invoiceBuffer(invoiceIndex, 1) & "_" & _
            invoiceBuffer(invoiceIndex, 2) & UniConvert(" ngafy ") & Format(invoiceBuffer(invoiceIndex, 5), "dd/mm/yyyy") & vbCrLf
        QueueFinalRetry apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), CStr(invoiceBuffer(invoiceIndex, 3)), _
            CStr(invoiceBuffer(invoiceIndex, 1)), CStr(invoiceBuffer(invoiceIndex, 2)), invoiceBuffer(invoiceIndex, 5), _
            UniConvert("Laasy chi tieest"), task.Endpoint, "DETAIL"
    End If
    AdvanceInvoiceWork UniConvert("DDax xuwr lys chi tieest ") & (invoiceIndex + 1)
End Sub

Private Sub ProcessOneRelationRequest( _
    ByRef invoiceBuffer As Variant, _
    ByVal invoiceIndex As Long, _
    ByVal invoiceType As Long, _
    ByVal apiSource As String, _
    ByVal endpointName As String, _
    Optional ByVal completedResult As clsGdtRequestResult = Nothing)

    Dim endpoint As String, actionHeader As String, stageName As String
    Dim requestResult As clsGdtRequestResult
    Dim parsed As Object
    Dim progressMessage As String
    Dim writeSucceeded As Boolean

    If completedResult Is Nothing Then
        If Not WaitForGdtControl() Then Exit Sub
    End If

    endpoint = BuildRelationEndpoint(apiSource, endpointName, CStr(invoiceBuffer(invoiceIndex, 0)), _
        CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
        CStr(invoiceBuffer(invoiceIndex, 2)))
    actionHeader = BuildRelationActionHeader(endpointName, apiSource, invoiceType)

    If endpointName = "relative" Then
        stageName = UniConvert("Laasy chuooxi hosa ddown lieen quan")
        progressMessage = UniConvert("DDang laasy chuooxi HDD lieen quan: ") & invoiceBuffer(invoiceIndex, 2)
    Else
        stageName = UniConvert("Laasy thoong tin lieen quan")
        progressMessage = UniConvert("DDang laasy thoong tin lieen quan: ") & invoiceBuffer(invoiceIndex, 2)
    End If
    If completedResult Is Nothing Then
        ShowInvoiceWork progressMessage, True
        Set requestResult = ExecuteGdtRequest("GET", endpoint, RunToken(), vbNullString, _
            "application/json", "application/json, text/plain, */*", False, GDT_MAX_RETRIES, _
            UCase$(endpointName), actionHeader)
    Else
        Set requestResult = completedResult
    End If
    PublishLastGdtResult requestResult

    If requestResult.Success Then
        If endpointName = "relative" Then
            If TryParseGdtJson(requestResult.ResponseText, parsed, apiSource, endpoint, _
                CStr(invoiceBuffer(invoiceIndex, 0)), CStr(invoiceBuffer(invoiceIndex, 3)), _
                CStr(invoiceBuffer(invoiceIndex, 1)), CStr(invoiceBuffer(invoiceIndex, 2)), _
                invoiceBuffer(invoiceIndex, 5)) Then
                WriteRelativeInvoiceData parsed, invoiceType, CLng(invoiceBuffer(invoiceIndex, 6)), _
                    CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
                    CStr(invoiceBuffer(invoiceIndex, 2))
                writeSucceeded = True
                MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
                    CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
                    CStr(invoiceBuffer(invoiceIndex, 2)), "Parse " & UniConvert("duwx lieeju"), requestResult.Attempts
            Else
                WriteRelationRequestError "RELATIVE", invoiceType, CLng(invoiceBuffer(invoiceIndex, 6)), _
                    requestResult.StatusCode, UniConvert("Pharn hoofi API khoong howjp leej"), requestResult.Attempts
            End If
        Else
            writeSucceeded = WriteRelatedRecord(requestResult.ResponseText, invoiceType, CLng(invoiceBuffer(invoiceIndex, 6)), _
                apiSource, endpoint, CStr(invoiceBuffer(invoiceIndex, 0)), CStr(invoiceBuffer(invoiceIndex, 3)), _
                CStr(invoiceBuffer(invoiceIndex, 1)), CStr(invoiceBuffer(invoiceIndex, 2)), invoiceBuffer(invoiceIndex, 5), requestResult.Attempts)
        End If
        If writeSucceeded Then
        MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
            CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
            CStr(invoiceBuffer(invoiceIndex, 2)), "Parse " & UniConvert("duwx lieeju"), requestResult.Attempts
        MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
            CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
            CStr(invoiceBuffer(invoiceIndex, 2)), stageName, requestResult.Attempts
        End If
    Else
        QueueFinalRetry apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
            CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
            CStr(invoiceBuffer(invoiceIndex, 2)), invoiceBuffer(invoiceIndex, 5), stageName, _
            endpoint, UCase$(endpointName), CLng(invoiceBuffer(invoiceIndex, 6)), actionHeader
        If Not LastGdtShouldQueue Or LastGdtAuthFailed Then
            WriteRelationRequestError UCase$(endpointName), invoiceType, _
                CLng(invoiceBuffer(invoiceIndex, 6)), requestResult.StatusCode, _
                requestResult.ErrorMessage, requestResult.Attempts
        End If
    End If

    AdvanceInvoiceWork UniConvert("DDax xuwr lys ") & endpointName & ": " & invoiceBuffer(invoiceIndex, 2)
    If completedResult Is Nothing Then WaitGdtMilliseconds GetSleepDelayMs()
End Sub

Private Function WriteRelatedRecord(ByVal responseText As String, ByVal invoiceType As Long, ByVal targetRow As Long, _
    ByVal apiSource As String, ByVal endpoint As String, ByVal seller As String, ByVal template As String, _
    ByVal series As String, ByVal number As String, ByVal invoiceDate As Variant, ByVal attempts As Long) As Boolean
    Dim parsed As Object, normalized As String
    On Error GoTo Failed
    normalized = Trim$(responseText)
    ' Empty related information is valid. Other responses must parse before
    ' the display formatter can keep their raw JSON as a compatibility fallback.
    If Len(normalized) > 0 And LCase$(normalized) <> "null" Then
        If Not TryParseGdtJson(responseText, parsed, apiSource, endpoint, seller, template, series, number, invoiceDate) Then GoTo InvalidResponse
    End If
    WriteRelatedInformationResponse responseText, invoiceType, targetRow
    WriteRelatedRecord = True
    Exit Function
Failed:
    UpsertGdtErrorReport CurrentDirectionName(), apiSource, seller, template, series, number, invoiceDate, _
        "Parse " & UniConvert("duwx lieeju"), endpoint, LastGdtStatus, Err.Description, attempts, 0, "Write failed"
InvalidResponse:
    WriteRelationRequestError "RELATED", invoiceType, targetRow, LastGdtStatus, _
        UniConvert("Pharn hoofi API khoong howjp leej"), attempts
End Function

Private Sub ProcessQueuedRelationRetries(ByVal invoiceType As Long)
    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    Dim parsed As Object, writeSucceeded As Boolean

    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "RELATIVE" Or queued.ResponseKind = "RELATED" Then
            ShowInvoiceWork UniConvert("Thuwr laji ") & LCase$(queued.ResponseKind) & ": " & queued.InvoiceNumber, True
            Set requestResult = ExecuteGdtRequest("GET", queued.Endpoint, RunToken(), vbNullString, _
                "application/json", "application/json, text/plain, */*", False, 1, _
                queued.ResponseKind, queued.ActionHeader)
            PublishLastGdtResult requestResult
            writeSucceeded = False

            If requestResult.Success Then
                If queued.ResponseKind = "RELATIVE" Then
                    If TryParseGdtJson(requestResult.ResponseText, parsed, queued.ApiSource, queued.Endpoint, _
                        queued.SellerTaxCode, queued.TemplateCode, queued.InvoiceSeries, _
                        queued.InvoiceNumber, queued.InvoiceDate) Then
                        WriteRelativeInvoiceData parsed, invoiceType, queued.TargetRow, queued.TemplateCode, _
                            queued.InvoiceSeries, queued.InvoiceNumber
                        writeSucceeded = True
                    Else
                        WriteRelationRequestError queued.ResponseKind, invoiceType, queued.TargetRow, _
                            requestResult.StatusCode, UniConvert("Pharn hoofi API khoong howjp leej"), _
                            queued.Attempts + requestResult.Attempts
                    End If
                Else
                    writeSucceeded = WriteRelatedRecord(requestResult.ResponseText, invoiceType, queued.TargetRow, _
                        queued.ApiSource, queued.Endpoint, queued.SellerTaxCode, queued.TemplateCode, _
                        queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate, queued.Attempts + requestResult.Attempts)
                End If

                If writeSucceeded Then
                    MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, "Parse " & UniConvert("duwx lieeju"), _
                        queued.Attempts + requestResult.Attempts
                    MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.Stage, _
                        queued.Attempts + requestResult.Attempts
                End If
            Else
                WriteRelationRequestError queued.ResponseKind, invoiceType, queued.TargetRow, _
                    requestResult.StatusCode, requestResult.ErrorMessage, _
                    queued.Attempts + requestResult.Attempts
                UpsertGdtErrorReport queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate, _
                    queued.Stage, queued.Endpoint, requestResult.StatusCode, requestResult.ErrorMessage, _
                    queued.Attempts + requestResult.Attempts, requestResult.RetryAfterSeconds, _
                    IIf(requestResult.AuthenticationFailure, UniConvert("Token heest hajn"), UniConvert("Khoong tari dduwowjc"))
            End If
            WaitGdtMilliseconds GetSleepDelayMs()
            If requestResult.AuthenticationFailure Then Exit For
        End If
    Next queued
End Sub


Private Function WriteDetailRecord(ByVal responseText As String, ByRef targetRow As Long, _
    ByVal apiSource As String, ByVal endpoint As String, ByVal seller As String, _
    ByVal template As String, ByVal series As String, ByVal number As String, ByVal invoiceDate As Variant) As Boolean
    Dim parsed As Object
    If Not TryParseGdtJson(responseText, parsed, apiSource, endpoint, seller, template, series, number, invoiceDate) Then Exit Function
    On Error GoTo Failed
    ghiExcel_ChiTiet responseText, targetRow, IIf(mRunPurchase, "mua", "ban")
    WriteDetailRecord = True
    Exit Function
Failed:
    UpsertGdtErrorReport CurrentDirectionName(), apiSource, seller, template, series, number, invoiceDate, _
        "Parse " & UniConvert("duwx lieeju"), endpoint, LastGdtStatus, Err.Description, LastGdtAttempts, 0, "Failed"
    AppendSimpleLog "Detail write failed: " & number
End Function

Private Sub ProcessQueuedDetailRetries(ByRef detailRow As Long)
    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "DETAIL" Then
            ShowInvoiceWork UniConvert("Thuwr laji chi tieest: ") & queued.InvoiceNumber, True
            Set requestResult = ExecuteGdtRequest("GET", queued.Endpoint, RunToken(), _
                vbNullString, "application/json", "application/json, text/plain, */*", False, 1, "DETAIL")
            PublishLastGdtResult requestResult
            If requestResult.Success Then
                If Not WriteDetailRecord(requestResult.ResponseText, detailRow, queued.ApiSource, queued.Endpoint, _
                    queued.SellerTaxCode, queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate) Then GoTo NextDetailRetry
                MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, "Parse " & UniConvert("duwx lieeju"), _
                    queued.Attempts + requestResult.Attempts
                MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.Stage, _
                    queued.Attempts + requestResult.Attempts
            Else
                UpsertGdtErrorReport queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate, _
                    queued.Stage, queued.Endpoint, requestResult.StatusCode, requestResult.ErrorMessage, _
                    queued.Attempts + requestResult.Attempts, requestResult.RetryAfterSeconds, _
                    IIf(requestResult.AuthenticationFailure, UniConvert("Token heest hajn"), UniConvert("Khoong tari dduwowjc"))
            End If
NextDetailRetry:
            WaitGdtMilliseconds GetSleepDelayMs()
        End If
        If GdtAuthenticationFailed Then Exit For
    Next queued
End Sub

Private Sub ResetProgressUI()
    Me.Controls("lblProgressFill").Width = 0
    Me.Controls("lblProgressPercent").caption = "0%"
    Me.Controls("lblProgressMessage").caption = UniConvert("Sawxn safng")
    Me.Controls("txtSimpleLog").Value = vbNullString
    SetDownloadControlState False
End Sub

Private Sub SetProgress(ByVal percentComplete As Long, ByVal message As String, Optional ByVal addToLog As Boolean = False)
    If percentComplete < 0 Then percentComplete = 0
    If percentComplete > 100 Then percentComplete = 100
    Me.Controls("lblProgressFill").Width = Me.Controls("lblProgressBack").Width * percentComplete / 100
    Me.Controls("lblProgressPercent").caption = CStr(percentComplete) & "%"
    Me.Controls("lblProgressMessage").caption = message
    Application.StatusBar = CStr(percentComplete) & "% - " & message
    If addToLog Then AppendSimpleLog message
    DoEvents
End Sub

Private Sub BeginInvoiceProgress(ByVal invoiceCount As Long, Optional ByVal relatedCallCount As Long = 0)
    Dim workKinds As Long
    workKinds = 0
    If Me.chkCT.Value = True Then workKinds = workKinds + 1
    If Me.chkXmlZip.Value = True Then workKinds = workKinds + 1
    mInvoiceWorkTotal = (invoiceCount * workKinds) + relatedCallCount
    mInvoiceWorkDone = 0
    If mInvoiceWorkTotal > 0 Then
        SetProgress 10, UniConvert("DDax laasy danh sasch ") & invoiceCount & UniConvert(" hosa ddown. Bawst ddaafu xuwr lys..."), True
    Else
        SetProgress 98, UniConvert("DDax laasy danh sasch ") & invoiceCount & UniConvert(" hosa ddown."), True
    End If
End Sub

Private Sub ShowInvoiceWork(ByVal message As String, Optional ByVal addToLog As Boolean = False)
    Dim currentPercent As Long
    If mInvoiceWorkTotal <= 0 Then
        currentPercent = 98
    Else
        currentPercent = 10 + CLng((mInvoiceWorkDone / mInvoiceWorkTotal) * 88)
    End If
    SetProgress currentPercent, message, addToLog
End Sub

Private Sub AdvanceInvoiceWork(ByVal message As String, Optional ByVal addToLog As Boolean = False)
    If mInvoiceWorkDone < mInvoiceWorkTotal Then mInvoiceWorkDone = mInvoiceWorkDone + 1
    ShowInvoiceWork message, addToLog
End Sub

Private Sub AppendSimpleLog(ByVal message As String)
    Dim logLine As String
    logLine = Format(Now, "hh:mm:ss") & "  " & message
    If Len(Me.Controls("txtSimpleLog").Value) = 0 Then
        Me.Controls("txtSimpleLog").Value = logLine
    Else
        Me.Controls("txtSimpleLog").Value = Me.Controls("txtSimpleLog").Value & vbCrLf & logLine
    End If
    If Len(Me.Controls("txtSimpleLog").Value) > 5000 Then Me.Controls("txtSimpleLog").Value = Right$(Me.Controls("txtSimpleLog").Value, 4000)
    Me.Controls("txtSimpleLog").SelStart = Len(Me.Controls("txtSimpleLog").Value)
End Sub

Private Function ElapsedTimerSeconds(ByVal started As Double) As Double
    ElapsedTimerSeconds = Timer - started
    If ElapsedTimerSeconds < 0 Then ElapsedTimerSeconds = ElapsedTimerSeconds + 86400#
End Function

Private Function ListPageLabel( _
    ByVal periodIndex As Long, _
    ByVal periodTotal As Long, _
    ByVal apiSource As String, _
    ByVal pageNumber As Long) As String

    ListPageLabel = UniConvert("Kyf ") & periodIndex & "/" & periodTotal & _
        " (" & Format(arrDate(periodIndex, 1), "dd/mm/yyyy") & "-" & _
        Format(arrDate(periodIndex, 2), "dd/mm/yyyy") & ") - " & apiSource & _
        UniConvert(" - trang ") & pageNumber
End Function

Private Sub ReportListPageStart( _
    ByVal periodIndex As Long, _
    ByVal periodTotal As Long, _
    ByVal apiSource As String, _
    ByVal pageNumber As Long)

    SetProgress 4 + CLng((periodIndex / periodTotal) * 5), _
        ListPageLabel(periodIndex, periodTotal, apiSource, pageNumber) & _
        UniConvert(": ddang guwri yeeu caafu; timeout ") & GDT_HTTP_TIMEOUT_SECONDS & UniConvert(" giaay"), True
End Sub

Private Sub ReportListPageComplete( _
    ByVal periodIndex As Long, _
    ByVal periodTotal As Long, _
    ByVal apiSource As String, _
    ByVal pageNumber As Long, _
    ByVal pageCount As Long, _
    ByVal cumulativeCount As Long, _
    ByVal started As Double, _
    ByVal hasMore As Boolean)

    Dim message As String
    message = ListPageLabel(periodIndex, periodTotal, apiSource, pageNumber) & _
        ": HTTP " & LastGdtStatus & UniConvert(", nhaajn ") & pageCount & _
        UniConvert(" HDD, toorng ") & cumulativeCount & ", " & _
        Format(ElapsedTimerSeconds(started), "0.0") & UniConvert(" giaay")
    If hasMore Then message = message & UniConvert(", cos trang tieesp") Else message = message & UniConvert(", heest trang")
    AppendSimpleLog message
End Sub

Private Sub ReportListPageFailure( _
    ByVal periodIndex As Long, _
    ByVal periodTotal As Long, _
    ByVal apiSource As String, _
    ByVal pageNumber As Long, _
    ByVal started As Double)

    AppendSimpleLog ListPageLabel(periodIndex, periodTotal, apiSource, pageNumber) & _
        ": HTTP " & LastGdtStatus & UniConvert(", thaast baji sau ") & _
        LastGdtAttempts & UniConvert(" laafn, ") & _
        Format(ElapsedTimerSeconds(started), "0.0") & UniConvert(" giaay")
End Sub

Private Function TryRegisterNextListState( _
    ByVal parsed As Object, _
    ByVal seenStates As Object, _
    ByVal apiSource As String, _
    ByRef nextState As String) As Boolean

    On Error GoTo NoNextState
    If IsNull(parsed("state")) Then Exit Function
    nextState = Trim$(CStr(parsed("state")))
    If Len(nextState) = 0 Then
        AppendSimpleLog apiSource & UniConvert(": state rooxng; duwfng phaan trang.")
        Exit Function
    End If
    If seenStates.Exists(nextState) Then
        AppendSimpleLog apiSource & UniConvert(": API trar laji state cux; duwfng ddeer trasnh lawjp voo hajn.")
        Exit Function
    End If
    seenStates.Add nextState, True
    TryRegisterNextListState = True
NoNextState:
    Err.Clear
End Function

Private Sub SetDownloadControlState(ByVal running As Boolean)
    Dim ctl As Object
    If running And Not mDownloadRunning Then
        mRunToken = cToken
        mRunPurchase = CBool(Me.optMua.Value)
        GdtReportAccount = CStr(Me.cboDonVi.Value)
        Set mControlStates = CreateObject("Scripting.Dictionary")
        For Each ctl In Me.Controls
            If TypeName(ctl) <> "Label" And ctl.Name <> "txtSimpleLog" And _
                ctl.Name <> "cmdPauseResume" And ctl.Name <> "cmdStopDownload" Then
                mControlStates(ctl.Name) = CBool(ctl.Enabled)
                ctl.Enabled = False
            End If
        Next ctl
    ElseIf Not running Then
        If Not mControlStates Is Nothing Then
            For Each ctl In Me.Controls
                If mControlStates.Exists(ctl.Name) Then ctl.Enabled = mControlStates(ctl.Name)
            Next ctl
        End If
        Set mControlStates = Nothing
        mRunToken = vbNullString
        mDownloadEntering = False
    End If
    mDownloadRunning = running
    On Error Resume Next
    Me.Controls("cmdPauseResume").Enabled = running
    Me.Controls("cmdPauseResume").caption = UniConvert("Tajm duwfng")
    Me.Controls("cmdStopDownload").Enabled = running
    On Error GoTo 0
End Sub

Public Sub TogglePauseDownload()
    If Not mDownloadRunning Or GdtStopRequested Then Exit Sub
    SetGdtPaused Not GdtPauseRequested
    If GdtPauseRequested Then
        Me.Controls("cmdPauseResume").caption = UniConvert("Tieesp tujc")
        AppendSimpleLog UniConvert("DDax tajm duwfng. Baasm Tieesp tujc ddeer chajy tieesp.")
    Else
        Me.Controls("cmdPauseResume").caption = UniConvert("Tajm duwfng")
        AppendSimpleLog UniConvert("Tieesp tujc xuwr lys.")
    End If
    DoEvents
End Sub

Public Sub RequestStopDownload()
    If Not mDownloadRunning Then Exit Sub
    RequestGdtStop
    Me.Controls("cmdPauseResume").Enabled = False
    Me.Controls("cmdStopDownload").Enabled = False
    AppendSimpleLog UniConvert("DDax yeeu caafu duwfng; chowf request hieejn taji keest thusc.")
    DoEvents
End Sub

Private Sub EnsureInvoiceBufferCapacity(ByRef buffer As Variant, ByVal targetIndex As Long)
    If targetIndex <= UBound(buffer, 1) Then Exit Sub
    Dim newCap As Long, tempBuf As Variant, r As Long, c As Long
    newCap = UBound(buffer, 1) * 2
    If newCap < targetIndex + 1000 Then newCap = targetIndex + 1000
    ReDim tempBuf(0 To newCap, 0 To 7)
    For r = 0 To targetIndex - 1
        For c = 0 To 7
            tempBuf(r, c) = buffer(r, c)
        Next c
    Next r
    buffer = tempBuf
End Sub

Private Function GetSleepDelayMs() As Long
    Dim delayVal As Long
    On Error Resume Next
    delayVal = CLng(Me.txtSleep.Value)
    On Error GoTo 0
    If delayVal <= 0 Then delayVal = GDT_REQUEST_DELAY_MS
    GetSleepDelayMs = delayVal
End Function


Private Sub EnsureRuntimeControls()
    Dim ctl As Object
    Dim handler As clsUiButtonHandler
    Set mUiHandlers = New Collection

    Set ctl = Me.fraChonNgay.Controls.Add("Forms.CommandButton.1", "cmdPickTuNgay", True)
    ctl.caption = ChrW(9660): ctl.Left = Me.fraChonNgay.InsideWidth - 28: ctl.Top = Me.txtTuNgay.Top: ctl.Width = 24: ctl.Height = Me.txtTuNgay.Height
    If Me.txtTuNgay.Left + Me.txtTuNgay.Width + 3 > ctl.Left Then Me.txtTuNgay.Width = ctl.Left - Me.txtTuNgay.Left - 3
    ctl.ControlTipText = UniConvert("Chojn ngafy bawst ddaafu")
    ctl.ZOrder 0
    Set handler = New clsUiButtonHandler: Set handler.Button = ctl: Set handler.Owner = Me: handler.ActionName = "OpenStartDatePicker": mUiHandlers.Add handler

    Set ctl = Me.fraChonNgay.Controls.Add("Forms.CommandButton.1", "cmdPickDenNgay", True)
    ctl.caption = ChrW(9660): ctl.Left = Me.fraChonNgay.InsideWidth - 28: ctl.Top = Me.txtDenNgay.Top: ctl.Width = 24: ctl.Height = Me.txtDenNgay.Height
    If Me.txtDenNgay.Left + Me.txtDenNgay.Width + 3 > ctl.Left Then Me.txtDenNgay.Width = ctl.Left - Me.txtDenNgay.Left - 3
    ctl.ControlTipText = UniConvert("Chojn ngafy keest thusc")
    ctl.ZOrder 0
    Set handler = New clsUiButtonHandler: Set handler.Button = ctl: Set handler.Owner = Me: handler.ActionName = "OpenEndDatePicker": mUiHandlers.Add handler

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblProgressMessage", True)
    ctl.caption = UniConvert("Sawxn safng"): ctl.Left = 15: ctl.Top = 378: ctl.Width = 515: ctl.Height = 24: ctl.Font.Bold = True
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblProgressBack", True)
    ctl.caption = "": ctl.Left = 15: ctl.Top = 399: ctl.Width = 752: ctl.Height = 18: ctl.BackColor = RGB(217, 217, 217): ctl.SpecialEffect = 1
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblProgressFill", True)
    ctl.caption = "": ctl.Left = 15: ctl.Top = 399: ctl.Width = 1: ctl.Height = 18: ctl.BackColor = RGB(128, 208, 128)
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblProgressPercent", True)
    ctl.caption = "0%": ctl.Left = 15: ctl.Top = 399: ctl.Width = 752: ctl.Height = 18: ctl.BackStyle = 0: ctl.TextAlign = 2: ctl.Font.Bold = True
    Set ctl = Me.Controls.Add("Forms.TextBox.1", "txtSimpleLog", True)
    ctl.Left = 15: ctl.Top = 423: ctl.Width = 752: ctl.Height = 84
    ctl.MultiLine = True: ctl.WordWrap = True: ctl.ScrollBars = 2: ctl.Locked = True: ctl.TabStop = False

    Set ctl = Me.Controls.Add("Forms.CommandButton.1", "cmdPauseResume", True)
    ctl.caption = UniConvert("Tajm duwfng"): ctl.Left = 545: ctl.Top = 375: ctl.Width = 100: ctl.Height = 24: ctl.Enabled = False
    ctl.ControlTipText = UniConvert("Tajm duwfng hoawjc tieesp tujc sau request hieejn taji")
    Set handler = New clsUiButtonHandler: Set handler.Button = ctl: Set handler.Owner = Me: handler.ActionName = "TogglePauseDownload": mUiHandlers.Add handler

    Set ctl = Me.Controls.Add("Forms.CommandButton.1", "cmdStopDownload", True)
    ctl.caption = UniConvert("Duwfng"): ctl.Left = 652: ctl.Top = 375: ctl.Width = 100: ctl.Height = 24: ctl.Enabled = False
    ctl.ControlTipText = UniConvert("Duwfng an toafn sau request hieejn taji")
    Set handler = New clsUiButtonHandler: Set handler.Button = ctl: Set handler.Owner = Me: handler.ActionName = "RequestStopDownload": mUiHandlers.Add handler
End Sub

Public Sub optMua_Change()
    Dim arrtthai(), arrttxly()
    If IIf(mDownloadRunning, mRunPurchase, Me.optMua.Value) Then
        arrtthai = ThisWorkbook.Sheets("LinkTraCuu").Range("H2:I8").Value
        Me.cboTTHD.list = arrtthai
        Me.cboTTHD.ListIndex = 0
        arrttxly = ThisWorkbook.Sheets("LinkTraCuu").Range("K2:L5").Value
        Me.cboKQKT.list = arrttxly
        Me.cboKQKT.ListIndex = 0
    Else
        arrtthai = ThisWorkbook.Sheets("LinkTraCuu").Range("H2:I8").Value
        Me.cboTTHD.list = arrtthai
        Me.cboTTHD.ListIndex = 0
        arrttxly = ThisWorkbook.Sheets("LinkTraCuu").Range("N2:O11").Value
        Me.cboKQKT.list = arrttxly
        Me.cboKQKT.ListIndex = 0
    End If

End Sub

Private Sub cmdTaiHoaDon_Click()
    If mDownloadRunning Or mDownloadEntering Then Exit Sub
    mDownloadEntering = True
    On Error GoTo DownloadFailed
    Application.ScreenUpdating = False
    ResetGdtOperationControl
    ResetGdtRequestSession
    Set mFinalRetryQueue = New Collection
    mFinalRetryCooldownDone = False
    EnsureGdtErrorReportSheet
    ResetProgressUI
    mDownloadEntering = True
    SetDownloadControlState True
    SetProgress 1, UniConvert("DDang kieerm tra duwx lieeju ddaafu vafo..."), True
    
    'Cap ngay phai hop le: ham nay cung tao lai bang ky thoi gian cho lan tai nay
    If Not lietKeThoiGian() Then
        AbortInvalidDateRange
        SetDownloadControlState False
        Exit Sub
    End If
    
    Me.lblStatus.Visible = True
    
    If Len(Trim(Me.cboDonVi)) = 0 Then
        MsgBoxUni UniConvert("Bajn chuwa chojn DDown vij caafn tari hosa ddown."), vbCritical
        Application.ScreenUpdating = True
        Application.StatusBar = False
        SetDownloadControlState False
        Exit Sub
    End If
    If Len(Trim(Me.txtTuNgay)) = 0 Or Len(Trim(Me.txtDenNgay)) = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        SetDownloadControlState False
        Exit Sub
    End If
    If Len(Trim(Me.cboTTHD)) = 0 Or Len(Trim(Me.cboKQKT)) = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        SetDownloadControlState False
        Exit Sub
    End If
    
    Dim sMsg As String, ret As Long
    Dim lrTongMua As Long, lrTongBan As Long, lrChiTietMua As Long, lrChiTietBan As Long
    
    'Lay dong cuoi tung Sheet de Xóa
    lrTongMua = ThisWorkbook.Sheets("TongHopHD_Mua").Cells(ThisWorkbook.Sheets("TongHopHD_mua").Rows.count, "C").End(xlUp).row
    lrTongBan = ThisWorkbook.Sheets("TongHopHD_Ban").Cells(ThisWorkbook.Sheets("TongHopHD_Ban").Rows.count, "C").End(xlUp).row
    lrChiTietMua = ThisWorkbook.Sheets("ChiTietHD_Mua").Cells(ThisWorkbook.Sheets("ChiTietHD_Mua").Rows.count, "C").End(xlUp).row
    lrChiTietBan = ThisWorkbook.Sheets("ChiTietHD_Ban").Cells(ThisWorkbook.Sheets("ChiTietHD_Ban").Rows.count, "C").End(xlUp).row
    
    sMsg = "B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n X" & ChrW(243) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(361) & " kh" & ChrW(244) & "ng?"
    sMsg = sMsg & vbCrLf & "Ch" & ChrW(7885) & "n [Yes] " & ChrW(273) & ChrW(7875) & " X" & ChrW(243) & "a ho" & ChrW(7863) & "c [No] " & ChrW(273) & ChrW(7875) & " ghi k" & ChrW(7871) & " ti" & ChrW(7871) & "p."
    ret = MsgBoxUni(sMsg, vbYesNo + vbInformation, UniConvert("Thoong baso"))
    
    'Timer----------
    Dim bd As Single, kt As Single
    bd = Timer
    SetProgress 3, UniConvert("DDang chuaarn bij tari hosa ddown..."), True
    SetDownloadControlState True
    '----------------
    
    If ret = vbYes Then
        If IIf(mDownloadRunning, mRunPurchase, Me.optMua.Value) Then
            If lrTongMua < 3 Then lrTongMua = 3
            If lrChiTietMua < 3 Then lrChiTietMua = 3
            ThisWorkbook.Sheets("TongHopHD_Mua").Range("A3:BZ" & lrTongMua).ClearContents
            ThisWorkbook.Sheets("ChiTietHD_Mua").Range("A3:BZ" & lrChiTietMua).ClearContents
        Else
            If lrTongBan < 3 Then lrTongBan = 3
            If lrChiTietBan < 3 Then lrChiTietBan = 3
            ThisWorkbook.Sheets("TongHopHD_Ban").Range("A3:BZ" & lrTongBan).ClearContents
            ThisWorkbook.Sheets("ChiTietHD_Ban").Range("A3:BZ" & lrChiTietBan).ClearContents
        End If
        Call taiHoaDon_Total
        
    Else
        If IIf(mDownloadRunning, mRunPurchase, Me.optMua.Value) Then
            Call taiHoaDon_Total(lrTongMua + 1, lrChiTietMua + 1)
        Else
            Call taiHoaDon_Total(lrTongBan + 1, lrChiTietBan + 1)
        End If
    End If
    
    '----------------------
    kt = Timer - bd
    Dim s As String
    s = " Thoi gian xu ly: " & Format(kt / 86400, "hh:mm:ss")
    '----------------------
    Me.lblStatus.caption = s
    Me.lblStatus.Visible = True
    If GdtAuthenticationFailed Then
        Me.Controls("lblProgressMessage").caption = UniConvert("Token heest hajn. Vui lofng ddawng nhaajp laji.")
        AppendSimpleLog Me.Controls("lblProgressMessage").caption
    ElseIf GdtStopRequested Then
        Me.Controls("lblProgressMessage").caption = UniConvert("DDax duwfng theo yeeu caafu. Duwx lieeju ddax tari vaaxn dduwowjc giuwx laji.")
        AppendSimpleLog Me.Controls("lblProgressMessage").caption
    Else
        SetProgress 100, UniConvert("Hoafn taast. ") & Trim$(s), True
    End If

    SetDownloadControlState False
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    'MsgBox "Xong."
    'Unload Me

    Exit Sub

DownloadFailed:
    AppendSimpleLog UniConvert("Looxi: ") & Err.Description
    MsgBoxUni UniConvert("Cos looxi phast sinh. Vui lofng kieerm tra BaoCao_LoiTaiHD."), vbExclamation, UniConvert("Thoong baso")
    SetDownloadControlState False
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub ReportRetryProgress(ByVal message As String)
    Me.Controls("lblProgressMessage").caption = message
    AppendSimpleLog message
    DoEvents
End Sub

Sub taiHoaDon_Total(Optional rowTotalstart As Long = 3, Optional rowDetailStart As Long = 3)
    Dim row As Long, row_ct As Long, stt As Long, n As Long
    Dim k As Long, j As Long, i As Long, loaiHD As Long, ret As Boolean
    Dim batchStartRow As Long, pageIndex As Long, relatedCallCount As Long
    Dim queryPageNumber As Long, scoPageNumber As Long
    Dim queryCount As Long, scoCount As Long
    Dim requestStarted As Double, hasNextPage As Boolean, nextState As String
    Dim seenQueryStates As Object, seenScoStates As Object
    Dim item As Object
    Dim arrHDChiTiet_tmp As Variant
    ReDim arrHDChiTiet_tmp(0 To 10000, 0 To 7)
    
    'Application.ScreenUpdating = False
    
    'Bang ky thoi gian phai co truoc khi duyet tung ky: arrDate bi Erase o cuoi moi
    'lan tai truoc va khong con phu thuoc su kien AfterUpdate cua o ngay
    If Not lietKeThoiGian() Then
        AbortInvalidDateRange
        Exit Sub
    End If

    'Tao bang tra cac dieu kien tim kiem
    Call LinkTraCuu
    Call tenCotTraCuu
    arrTrangThai = Sheets("LinkTraCuu").Range("I2:I8").Value
    arrKQKTHoaDon = Sheets("LinkTraCuu").Range("O2:O11").Value
    '------------------------------------
    
    stt = 1
    row = rowTotalstart: row_ct = rowDetailStart
    
    'Tham so cho request
    sort = "tdlap:desc"
    size = 50
    ttxly = Me.cboKQKT
    tthai = Me.cboTTHD
    bearer = RunToken()
    url2 = ""
    If IIf(mDownloadRunning, mRunPurchase, Me.optMua.Value) Then
        url = "https://hoadondientu.gdt.gov.vn/api/query/invoices/purchase?sort="
        url_sco = "https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/purchase?sort="
        loaiHD = 1 'mua
    Else
        url = "https://hoadondientu.gdt.gov.vn/api/query/invoices/sold?sort="
        url_sco = "https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/sold?sort="
        loaiHD = 2 'ban
    End If
    n = 0   'n: so luong HD
    'Duyet tung khoan thoi gian de trich xuat hoa don
    For k = 1 To UBound(arrDate)
        If Not WaitForGdtControl() Then GoTo cleanup
        SetProgress 4 + CLng((k / UBound(arrDate)) * 5), _
            UniConvert("Chuaarn bij kyf ") & k & "/" & UBound(arrDate) & " (" & _
            Format(arrDate(k, 1), "dd/mm/yyyy") & "-" & Format(arrDate(k, 2), "dd/mm/yyyy") & ")", True
        queryPageNumber = 1
        scoPageNumber = 1
        queryCount = 0
        scoCount = 0
        Set seenQueryStates = CreateObject("Scripting.Dictionary")
        Set seenScoStates = CreateObject("Scripting.Dictionary")
        If tthai = "All" Then 'Tat ca
            'Hd mua va ban giong nhau
            If ttxly = "All" Then  'Tat ca
                search = "tdlap=ge=" & ApiDateText(arrDate(k, 1)) & "T00:00:00;tdlap=le=" & ApiDateText(arrDate(k, 2)) & "T23:59:59"
            Else
                search = "tdlap=ge=" & ApiDateText(arrDate(k, 1)) & "T00:00:00;tdlap=le=" & ApiDateText(arrDate(k, 2)) & "T23:59:59;ttxly==" & ttxly
            End If
        Else
            If ttxly = "All" Then
                search = "tdlap=ge=" & ApiDateText(arrDate(k, 1)) & "T00:00:00;tdlap=le=" & ApiDateText(arrDate(k, 2)) & "T23:59:59;tthai==" & tthai
            Else
                search = "tdlap=ge=" & ApiDateText(arrDate(k, 1)) & "T00:00:00;tdlap=le=" & ApiDateText(arrDate(k, 2)) & "T23:59:59;tthai==" & tthai & ";ttxly==" & ttxly
            End If
        End If
        url2 = url & sort & "&size=" & size & "&search=" & search
        
nextPage:
        If Not WaitForGdtControl() Then GoTo cleanup
        ReportListPageStart k, UBound(arrDate), "query", queryPageNumber
        requestStarted = Timer
        res = ApiGet(url2, ListPageLabel(k, UBound(arrDate), "query", queryPageNumber), RunToken())
        If Len(res) = 0 Then
            ReportListPageFailure k, UBound(arrDate), "query", queryPageNumber, requestStarted
            If GdtStopRequested Then GoTo cleanup
            errMsg = errMsg & UniConvert("Looxi tari hoas ddown toorng howjp ngafy: ") & arrDate(k, 1) & " - " & arrDate(k, 2) & vbCrLf
            QueueFinalRetry "query", vbNullString, vbNullString, vbNullString, vbNullString, _
                arrDate(k, 1), UniConvert("Laasy danh sasch"), url2, "LIST"
            'If getStatus = 429 Or getStatus = 500 Then GoTo errLog
            'GoTo nextDatePeriod 'Khi co loi phat sinh --> tiep tuc lay du lieu khoang thoi gian ke tiep
            GoTo startSco
        End If
        
        
        If Not TryParseGdtJson(res, js, "query", url2, , , , , arrDate(k, 1)) Then
            ReportListPageFailure k, UBound(arrDate), "query", queryPageNumber, requestStarted
            GoTo startSco
        End If
        batchStartRow = row
        If Not WriteSummaryPage(res, row, loaiHD, stt, "query", url2, arrDate(k, 1)) Then GoTo startSco
        MarkGdtRetrySuccess CurrentDirectionName(), "query", vbNullString, vbNullString, vbNullString, _
            vbNullString, UniConvert("Laasy danh sasch"), LastGdtAttempts, vbNullString, url2
        MarkGdtRetrySuccess CurrentDirectionName(), "query", vbNullString, vbNullString, vbNullString, _
            vbNullString, "Parse " & UniConvert("duwx lieeju"), LastGdtAttempts, vbNullString, url2
        
        pageIndex = 0
        
        'Lay tham so HD chi tiet
        For Each item In js("datas")
            pageIndex = pageIndex + 1
            EnsureInvoiceBufferCapacity arrHDChiTiet_tmp, n
            arrHDChiTiet_tmp(n, 0) = item("nbmst") & ""
            arrHDChiTiet_tmp(n, 1) = item("khhdon")
            arrHDChiTiet_tmp(n, 2) = item("shdon")
            arrHDChiTiet_tmp(n, 3) = item("khmshdon")
            arrHDChiTiet_tmp(n, 4) = 1  'query
            arrHDChiTiet_tmp(n, 5) = ISODateValue(item("tdlap"))
            arrHDChiTiet_tmp(n, 6) = batchStartRow + pageIndex - 1
            arrHDChiTiet_tmp(n, 7) = CLng(Val(item("tthai") & vbNullString))
            n = n + 1
        Next
        queryCount = queryCount + pageIndex
        
        'Tai nhieu trang
        nextState = vbNullString
        hasNextPage = TryRegisterNextListState(js, seenQueryStates, "query", nextState)
        ReportListPageComplete k, UBound(arrDate), "query", queryPageNumber, pageIndex, queryCount, requestStarted, hasNextPage
        If hasNextPage Then
            queryPageNumber = queryPageNumber + 1
            url2 = url & sort & "&size=" & size & "&state=" & nextState & "&search=" & search
            GoTo nextPage 'Quay lai lay du lieu tiep trang 2 (>50)...
        End If
        
           '******************************************
        '// Chay sco query de lay hd tu may tinh tien
startSco:
        url2 = url_sco & sort & "&size=" & size & "&search=" & search
nextPage_sco:
        If Not WaitForGdtControl() Then GoTo cleanup
        ReportListPageStart k, UBound(arrDate), "sco-query", scoPageNumber
        requestStarted = Timer
        res = ApiGet(url2, ListPageLabel(k, UBound(arrDate), "sco-query", scoPageNumber), RunToken())
        If Len(res) = 0 Then
            ReportListPageFailure k, UBound(arrDate), "sco-query", scoPageNumber, requestStarted
            If GdtStopRequested Then GoTo cleanup
            errMsg = errMsg & UniConvert("Looxi tari hoas ddown toorng howjp - Tuwf masy tisnh tieefn ngafy: ") & arrDate(k, 1) & " - " & arrDate(k, 2) & vbCrLf
            QueueFinalRetry "sco-query", vbNullString, vbNullString, vbNullString, vbNullString, _
                arrDate(k, 1), UniConvert("Laasy danh sasch"), url2, "LIST"
            'If getStatus = 429 Or getStatus = 500 Then GoTo errLog
            GoTo nextDatePeriod
        End If
        
        
        If Not TryParseGdtJson(res, js, "sco-query", url2, , , , , arrDate(k, 1)) Then
            ReportListPageFailure k, UBound(arrDate), "sco-query", scoPageNumber, requestStarted
            GoTo nextDatePeriod
        End If
        batchStartRow = row
        If Not WriteSummaryPage(res, row, loaiHD, stt, "sco-query", url2, arrDate(k, 1)) Then GoTo nextDatePeriod
        MarkGdtRetrySuccess CurrentDirectionName(), "sco-query", vbNullString, vbNullString, vbNullString, _
            vbNullString, UniConvert("Laasy danh sasch"), LastGdtAttempts, vbNullString, url2
        MarkGdtRetrySuccess CurrentDirectionName(), "sco-query", vbNullString, vbNullString, vbNullString, _
            vbNullString, "Parse " & UniConvert("duwx lieeju"), LastGdtAttempts, vbNullString, url2
        
        pageIndex = 0
        
        'Lay tham so HD chi tiet
        For Each item In js("datas")
            pageIndex = pageIndex + 1
            EnsureInvoiceBufferCapacity arrHDChiTiet_tmp, n
            arrHDChiTiet_tmp(n, 0) = item("nbmst") & ""
            arrHDChiTiet_tmp(n, 1) = item("khhdon")
            arrHDChiTiet_tmp(n, 2) = item("shdon")
            arrHDChiTiet_tmp(n, 3) = item("khmshdon")
            arrHDChiTiet_tmp(n, 4) = 2  'sco-query
            arrHDChiTiet_tmp(n, 5) = ISODateValue(item("tdlap"))
            arrHDChiTiet_tmp(n, 6) = batchStartRow + pageIndex - 1
            arrHDChiTiet_tmp(n, 7) = CLng(Val(item("tthai") & vbNullString))
            n = n + 1
        Next
        scoCount = scoCount + pageIndex
        
        'Tai nhieu trang
        nextState = vbNullString
        hasNextPage = TryRegisterNextListState(js, seenScoStates, "sco-query", nextState)
        ReportListPageComplete k, UBound(arrDate), "sco-query", scoPageNumber, pageIndex, scoCount, requestStarted, hasNextPage
        If hasNextPage Then
            scoPageNumber = scoPageNumber + 1
            url2 = url_sco & sort & "&size=" & size & "&state=" & nextState & "&search=" & search
            GoTo nextPage_sco
        End If
        
        '++++++++++++++++++++++++++++++++
        
nextDatePeriod: 'arrDate ke tiep
        AppendSimpleLog UniConvert("Hoafn taast kyf ") & k & "/" & UBound(arrDate) & _
            ": query=" & queryCount & ", sco-query=" & scoCount & _
            UniConvert(", toorng ddax ghi=") & n
    Next k

    If GdtStopRequested Then GoTo cleanup
    ProcessQueuedListRetries arrHDChiTiet_tmp, n, row, stt, loaiHD
    If GdtStopRequested Then GoTo cleanup
    
    '/Resize mang arrHDChiTiet()
    ReDim arrHDChiTiet(n, 7)
    For k = 0 To n - 1
        For i = 0 To 7
            arrHDChiTiet(k, i) = arrHDChiTiet_tmp(k, i)
        Next i
    Next
    Erase arrHDChiTiet_tmp
    relatedCallCount = CountRelatedApiCalls(arrHDChiTiet, n)
    BeginInvoiceProgress n, relatedCallCount
    ProcessRelatedInvoiceApis arrHDChiTiet, n, loaiHD
    If GdtStopRequested Then GoTo cleanup
    ProcessQueuedRelationRetries loaiHD
    If GdtStopRequested Then GoTo cleanup
    
    
    If Me.chkCT = False Then GoTo taiXML
    
    '------------------------------------------
    ' HTTP runs concurrently; only this Excel thread writes the result rows.
    ProcessJsonInvoiceRequests arrHDChiTiet, n, loaiHD, True, row_ct
    If GdtStopRequested Then GoTo cleanup
    ProcessQueuedDetailRetries row_ct
    If GdtStopRequested Then GoTo cleanup
    
taiXML:
    If Me.chkXmlZip = True Then
        ShowInvoiceWork UniConvert("DDang tari vaf giari ne_sn XML/HTML..."), True
        Me.lblStatus.caption = "B" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u t" & ChrW(7843) & "i file XML/HTML v" & ChrW(224) & " gi" & ChrW(7843) & "i n" & ChrW(233) & "n..."
        
        DoEvents
        taiXML_zip n
    End If

errLog:
    If errMsg <> "" Then
        ghiLog errMsg
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If Not GdtStopRequested And Not GdtAuthenticationFailed Then MsgBox "Xong."

cleanup:
    FinalizeInvoiceSheets loaiHD
    Erase arrDate
    Erase arrHDChiTiet
    'Unload Me
End Sub

Function lietKeThoiGian() As Boolean
    Dim startYear As Long, startMonth As Long, startDay As Long
    Dim dtNgayCuoiThang As Date, dtNgayDauThang As Date, dtDenNgay As Date
    Dim tempArr(), scratchArr()   'De redim arrDate
    Dim soKy As Long, blnErr As Boolean, dtTuNgay As Date
    
    On Error GoTo InvalidDate

    'Cac o ngay duoc kiem tra bang chinh ham correctDate cua form
    Erase arrDate
    dtTuNgay = correctDate(Trim$(Me.txtTuNgay.Value), blnErr)
    If blnErr Then Exit Function
    dtDenNgay = correctDate(Trim$(Me.txtDenNgay.Value), blnErr)
    If blnErr Then Exit Function
    If dtDenNgay < dtTuNgay Then Exit Function

    startYear = Year(dtTuNgay)
    startMonth = Month(dtTuNgay)
    startDay = Day(dtTuNgay)
    
    dtNgayDauThang = DateSerial(startYear, startMonth, startDay)
    dtNgayCuoiThang = DateSerial(startYear, startMonth + 1, 0)
    'So ky toi da = so thang cham toi + ky cuoi, khong con gioi han 5 nam
    soKy = (Year(dtDenNgay) * 12 + Month(dtDenNgay)) - (Year(dtTuNgay) * 12 + Month(dtTuNgay)) + 2
    i = 1
    ReDim scratchArr(1 To soKy, 1 To 2)
    Do While dtNgayCuoiThang <= dtDenNgay
        
        scratchArr(i, 1) = dtNgayDauThang: scratchArr(i, 2) = dtNgayCuoiThang
        startMonth = startMonth + 1
        'Qua nam ke tiep
        If startMonth > 12 Then
            startMonth = 1
            startYear = startYear + 1
        End If
        dtNgayDauThang = DateSerial(startYear, startMonth, 1)
        dtNgayCuoiThang = DateSerial(startYear, startMonth + 1, 0)
        i = i + 1
    Loop
    If dtNgayDauThang <= dtDenNgay Then
        scratchArr(i, 1) = dtNgayDauThang: scratchArr(i, 2) = dtDenNgay
    Else
        i = i - 1
    End If
    
    If i < 1 Then Exit Function

    ReDim tempArr(1 To i, 1 To 2)
    Dim k As Long
    For k = 1 To i
        tempArr(k, 1) = scratchArr(k, 1)
        tempArr(k, 2) = scratchArr(k, 2)
    Next
    
    arrDate = tempArr
    lietKeThoiGian = True
    
    'For k = 1 To UBound(arrDate)
        'Debug.Print arrDate(k, 1), arrDate(k, 2)
    'Next
    Exit Function

InvalidDate:
    Erase arrDate
    lietKeThoiGian = False
End Function

Private Function ApiDateText(ByVal value As Date) As String
    ApiDateText = Format$(Day(value), "00") & "/" & Format$(Month(value), "00") & "/" & Format$(Year(value), "0000")
End Function

Sub ghiLog(errorMessage As String, Optional txtFile As Boolean = False)
'    Select Case txtFile
'        Case True
'            Dim filePath As String, fileNum As Integer
'            filePath = Me.txtFolderPath & "errLog_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
'            fileNum = FreeFile
'            Open filePath For Output As fileNum
'            Print #fileNum, errorMessage
'            Close fileNum
'            'Debug.Print "Da luu log file thanh cong!"
'        Case False
'            CurrentDb.Execute "Insert Into tblErrLog (NgayPS, NoiDung) Values (#" & Now & "#,'" & errorMessage & "')", dbFailOnError
'    End Select
End Sub

Sub taiXML_zip(soHD As Long)
    Dim tasks As Collection, task As clsGdtXmlTask, throttle As clsGdtXmlThrottle
    Dim nextIndex As Long, active As Long, pos As Long, finishedCount As Long, endpoint As String
    Dim cooldownLeft As Double
    On Error GoTo Failed
    Set tasks = New Collection
    Set throttle = New clsGdtXmlThrottle
    ' The XML phase keeps its own pacing policy. The "khoang nghi" cell on the
    ' form is the spacing of the serial list/detail/related requests; here the
    ' throttle starts at GDT_XML_INTERVAL_MS and raises itself on HTTP 429, so
    ' the faster start does not need to be undone by hand.
    throttle.Configure GDT_XML_CONCURRENCY, GDT_XML_INTERVAL_MS
    AppendSimpleLog "XML scheduler: " & throttle.CurrentConcurrency & " connections, " & throttle.CurrentIntervalMs & " ms"
    Do While nextIndex < soHD Or tasks.Count > 0
        If GdtStopRequested Or GdtAuthenticationFailed Then Exit Do
        If Not GdtPauseRequested Then
            Do While nextIndex < soHD And tasks.Count < throttle.CurrentConcurrency
                endpoint = "https://hoadondientu.gdt.gov.vn/api/" & IIf(arrHDChiTiet(nextIndex, 4) = 1, "query", "sco-query") & _
                    "/invoices/export-xml?nbmst=" & arrHDChiTiet(nextIndex, 0) & "&khhdon=" & arrHDChiTiet(nextIndex, 1) & _
                    "&shdon=" & arrHDChiTiet(nextIndex, 2) & "&khmshdon=" & arrHDChiTiet(nextIndex, 3)
                Set task = New clsGdtXmlTask
                task.Configure nextIndex, endpoint, RunToken()
                tasks.Add task
                nextIndex = nextIndex + 1
            Loop
        End If
        active = 0
        For Each task In tasks
            If task.Running Then active = active + 1
        Next task
        For pos = tasks.Count To 1 Step -1
            Set task = tasks(pos)
            If task.Running Then active = active - 1
            task.Tick throttle, active
            If task.Running Then active = active + 1
            If task.Finished Then
                If task.Result.StatusCode = 429 Then
                    cooldownLeft = throttle.CooldownUntil - GdtClockSeconds()
                    If cooldownLeft < 0 Then cooldownLeft = 0
                    AppendSimpleLog "XML 429 | cooldown " & Format$(cooldownLeft, "0.0") & "s | nodes " & _
                        throttle.CurrentConcurrency & " | " & throttle.CurrentIntervalMs & " ms"
                End If
                HandleXmlResult task
                tasks.Remove pos
                finishedCount = finishedCount + 1
                AdvanceInvoiceWork UniConvert("DDax xuwr lys XML ") & finishedCount & "/" & soHD & _
                    " (" & throttle.CurrentConcurrency & " connections; " & throttle.CurrentIntervalMs & " ms)", True
            End If
            If GdtStopRequested Or GdtAuthenticationFailed Then Exit For
        Next pos
        GdtPumpWait
    Loop
    For Each task In tasks
        task.Cancel
    Next task
    If Not GdtStopRequested And Not GdtAuthenticationFailed Then ProcessQueuedXmlRetries
    Exit Sub
Failed:
    AppendSimpleLog "XML scheduler: " & Err.Description
    If Not tasks Is Nothing Then
        For Each task In tasks
            task.Cancel
        Next task
    End If
End Sub

Private Sub HandleXmlResult(ByVal task As clsGdtXmlTask)
    Dim m As Long, apiSource As String, result As clsGdtRequestResult
    m = task.InvoiceIndex
    apiSource = IIf(arrHDChiTiet(m, 4) = 1, "query", "sco-query")
    Set result = task.Result
    PublishLastGdtResult result
    If result.Success Then
        If TrySaveXml(result.ResponseBody, apiSource, task.Endpoint, CStr(arrHDChiTiet(m, 0)), _
            CStr(arrHDChiTiet(m, 3)), CStr(arrHDChiTiet(m, 1)), CStr(arrHDChiTiet(m, 2)), arrHDChiTiet(m, 5)) Then
            MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(arrHDChiTiet(m, 0)), CStr(arrHDChiTiet(m, 3)), _
                CStr(arrHDChiTiet(m, 1)), CStr(arrHDChiTiet(m, 2)), UniConvert("Tari XML"), result.Attempts
        End If
    ElseIf Not GdtStopRequested Then
        QueueFinalRetry apiSource, CStr(arrHDChiTiet(m, 0)), CStr(arrHDChiTiet(m, 3)), _
            CStr(arrHDChiTiet(m, 1)), CStr(arrHDChiTiet(m, 2)), arrHDChiTiet(m, 5), UniConvert("Tari XML"), task.Endpoint, "XML"
    End If
End Sub

Private Function TrySaveXml(ByVal body As Variant, ByVal apiSource As String, ByVal endpoint As String, _
    ByVal seller As String, ByVal template As String, ByVal series As String, ByVal number As String, ByVal invoiceDate As Variant) As Boolean
    On Error GoTo Failed
    SaveAndUnzipGdtXml body, series, number, seller, template
    TrySaveXml = True
    Exit Function
Failed:
    UpsertGdtErrorReport CurrentDirectionName(), apiSource, seller, template, series, number, invoiceDate, _
        UniConvert("Tari XML"), endpoint, LastGdtStatus, Err.Description, LastGdtAttempts, 0, "Save/extract failed"
    AppendSimpleLog "XML save/extract failed: " & number
End Function

Private Sub SaveAndUnzipGdtXml(ByVal responseBody As Variant, ByVal invoiceSeries As String, ByVal invoiceNumber As String, ByVal sellerTaxCode As String, ByVal templateCode As String)
    Dim stream As Object, filePath As String, staging As String, stagedZip As String, parent As String
    Dim fso As Object, failureNumber As Long, failureText As String, stagingVerified As Boolean
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = fso.BuildPath(fso.GetAbsolutePathName(CStr(Me.txtXMLFolderPath.Value)), _
        GdtInvoiceFileBase(sellerTaxCode, templateCode, invoiceSeries, invoiceNumber) & ".zip")
    parent = fso.GetParentFolderName(filePath)
    staging = fso.BuildPath(parent, ".hddt-zip-" & Replace(CreateRequestID(), "-", ""))
    If StrComp(fso.GetParentFolderName(staging), parent, vbTextCompare) <> 0 Then Err.Raise 5, , "Invalid ZIP staging path"
    stagingVerified = True
    fso.CreateFolder staging
    stagedZip = fso.BuildPath(staging, fso.GetFileName(filePath))
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write responseBody
    stream.SaveToFile stagedZip, 2
    stream.Close
    Set stream = Nothing
    Unzip stagedZip, oSaveUnzipFolder
    ' A bad response must not overwrite the last valid ZIP of this invoice.
    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    fso.MoveFile stagedZip, filePath
    fso.DeleteFolder staging, True
    Exit Sub
Failed:
    failureNumber = Err.Number: failureText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If stagingVerified Then
        If fso.FolderExists(staging) Then fso.DeleteFolder staging, True
    End If
    On Error GoTo 0
    Err.Raise failureNumber, "SaveAndUnzipGdtXml", failureText
End Sub

Private Sub ProcessQueuedXmlRetries()
    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "XML" Then
            ShowInvoiceWork UniConvert("Thuwr laji XML: ") & queued.InvoiceNumber, True
            Set requestResult = ExecuteGdtBinaryRequest(queued.Endpoint, RunToken(), 1, "XML")
            PublishLastGdtResult requestResult
            If requestResult.Success Then
                If Not TrySaveXml(requestResult.ResponseBody, queued.ApiSource, queued.Endpoint, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate) Then GoTo NextXmlRetry
                MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.Stage, _
                    queued.Attempts + requestResult.Attempts
            Else
                UpsertGdtErrorReport queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                    queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate, _
                    queued.Stage, queued.Endpoint, requestResult.StatusCode, requestResult.ErrorMessage, _
                    queued.Attempts + requestResult.Attempts, requestResult.RetryAfterSeconds, _
                    IIf(requestResult.NoXml, UniConvert("Khoong cos XML"), _
                        IIf(requestResult.AuthenticationFailure, UniConvert("Token heest hajn"), UniConvert("Khoong tari dduwowjc")))
            End If
NextXmlRetry:
            WaitGdtMilliseconds GetSleepDelayMs()
            If requestResult.AuthenticationFailure Then Exit For
        End If
    Next queued
End Sub



