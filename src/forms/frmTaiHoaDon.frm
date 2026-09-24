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
    Me.txtTuNgay.Value = Format(correctDate(Me.txtTuNgay.Value, blnErr), "dd/mm/yyyy")
    If blnErr = True Then
        Me.lblSaiNgay.Visible = True
    Else
        Me.lblSaiNgay.Visible = False
    End If
End Sub

Private Sub txtTuNgay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtDenNgay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtDenNgay_AfterUpdate()
    Dim blnErr As Boolean
    Me.txtDenNgay.Value = Format(correctDate(Me.txtDenNgay.Value, blnErr), "dd/mm/yyyy")
    If blnErr = True Then
        Me.lblSaiNgay.Visible = True
    Else
        Me.lblSaiNgay.Visible = False
    End If
    'Khong so sanh duoc khi mot trong hai o ngay con trong
    If Len(Trim(Me.txtTuNgay)) = 0 Or Len(Trim(Me.txtDenNgay)) = 0 Then Exit Sub
    If CDate(Me.txtDenNgay) < CDate(Me.txtTuNgay) Then
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
    Dim dtNgayNhap As Date, dtNgayCuoiThang As Date
    
    If Len(textDate) < 10 Or Len(textDate) > 10 Then
        correctDate = Date
        err = True
        Exit Function
    ElseIf Mid(textDate, 3, 1) <> "/" Or Mid(textDate, 6, 1) <> "/" Then
        correctDate = Date
        err = True
        Exit Function
    End If
    
    ngay = CLng(Left(textDate, 2))
    thang = CLng(Mid(textDate, 4, 2))
    nam = CLng(Right(textDate, 4))
    
    If thang > 12 Then thang = 12
    
    dtNgayNhap = DateSerial(nam, thang, ngay)
    dtNgayCuoiThang = DateSerial(nam, thang + 1, 0)
    If dtNgayNhap > dtNgayCuoiThang Then dtNgayNhap = dtNgayCuoiThang
    
    correctDate = CStr(dtNgayNhap)
    
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
    If Me.optMua.Value Then
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
    If InStr(1, responseText, "error", vbTextCompare) > 0 Then _
        Err.Raise vbObjectError + 710, "TryParseGdtJson", UniConvert("Pharn hoofi API chuwsa looxi")
    Set parsed = JsonConverter.ParseJSON(responseText)
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
            retryPageNumber = 1
            retryCount = 0
            Do
                If Not WaitForGdtControl() Then Exit Sub
                hasMore = False
                SetProgress 9, UniConvert("Thuwr laji danh sasch ") & queued.ApiSource & _
                    UniConvert(" - trang ") & retryPageNumber, True
                retryStarted = Timer
                Set requestResult = ExecuteGdtRequest("GET", currentEndpoint, getToken(), _
                    vbNullString, "application/json", "application/json, text/plain, */*", False, 1, "LIST")
                PublishLastGdtResult requestResult
                If requestResult.Success Then
                    If Not TryParseGdtJson(requestResult.ResponseText, parsed, queued.ApiSource, currentEndpoint, _
                        queued.SellerTaxCode, queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.InvoiceDate) Then Exit Do
                    MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, _
                        "Parse " & UniConvert("duwx lieeju"), queued.Attempts + requestResult.Attempts
                    batchStartRow = totalRow
                    ghiExcel_TongHop requestResult.ResponseText, totalRow, invoiceType, sequenceNumber
                    pageIndex = 0
                    For Each dataItem In parsed("datas")
                        pageIndex = pageIndex + 1
                        EnsureInvoiceBufferCapacity invoiceBuffer, invoiceCount
                        invoiceBuffer(invoiceCount, 0) = dataItem("nbmst") & vbNullString
                        invoiceBuffer(invoiceCount, 1) = dataItem("khhdon")
                        invoiceBuffer(invoiceCount, 2) = dataItem("shdon")
                        invoiceBuffer(invoiceCount, 3) = dataItem("khmshdon")
                        invoiceBuffer(invoiceCount, 4) = IIf(queued.ApiSource = "query", 1, 2)
                        invoiceBuffer(invoiceCount, 5) = ISODATE(dataItem("tdlap"))
                        invoiceBuffer(invoiceCount, 6) = batchStartRow + pageIndex - 1
                        invoiceBuffer(invoiceCount, 7) = CLng(Val(dataItem("tthai") & vbNullString))
                        invoiceCount = invoiceCount + 1
                    Next dataItem
                    retryCount = retryCount + pageIndex
                    MarkGdtRetrySuccess queued.Direction, queued.ApiSource, queued.SellerTaxCode, _
                        queued.TemplateCode, queued.InvoiceSeries, queued.InvoiceNumber, queued.Stage, _
                        queued.Attempts + requestResult.Attempts, UniConvert("Danh sasch ddax tari thafnh coong")
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

    Dim invoiceIndex As Long, invoiceStatus As Long
    Dim apiSource As String

    'The summary list is already in memory for the report. Filter it locally so
    'we do not download and paginate the same list again for every status/date.
    For invoiceIndex = 0 To invoiceCount - 1
        If Not WaitForGdtControl() Then Exit For
        invoiceStatus = CLng(Val(invoiceBuffer(invoiceIndex, 7) & vbNullString))
        If invoiceStatus >= 2 And invoiceStatus <= 6 Then
            apiSource = IIf(invoiceBuffer(invoiceIndex, 4) = 1, "query", "sco-query")
            If invoiceStatus = 6 Then
                WriteNoRelativeInvoiceData invoiceType, CLng(invoiceBuffer(invoiceIndex, 6))
            Else
                ProcessOneRelationRequest invoiceBuffer, invoiceIndex, invoiceType, apiSource, "relative"
                If GdtAuthenticationFailed Or GdtStopRequested Then Exit For
            End If
            ProcessOneRelationRequest invoiceBuffer, invoiceIndex, invoiceType, apiSource, "related"
            If GdtAuthenticationFailed Or GdtStopRequested Then Exit For
        End If
    Next invoiceIndex
End Sub

Private Sub ProcessOneRelationRequest( _
    ByRef invoiceBuffer As Variant, _
    ByVal invoiceIndex As Long, _
    ByVal invoiceType As Long, _
    ByVal apiSource As String, _
    ByVal endpointName As String)

    Dim endpoint As String, actionHeader As String, stageName As String
    Dim requestResult As clsGdtRequestResult
    Dim parsed As Object
    Dim progressMessage As String

    If Not WaitForGdtControl() Then Exit Sub

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
    ShowInvoiceWork progressMessage, True

    Set requestResult = ExecuteGdtRequest("GET", endpoint, getToken(), vbNullString, _
        "application/json", "application/json, text/plain, */*", False, GDT_MAX_RETRIES, _
        UCase$(endpointName), actionHeader)
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
                MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
                    CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
                    CStr(invoiceBuffer(invoiceIndex, 2)), "Parse " & UniConvert("duwx lieeju"), requestResult.Attempts
            Else
                WriteRelationRequestError "RELATIVE", invoiceType, CLng(invoiceBuffer(invoiceIndex, 6)), _
                    requestResult.StatusCode, UniConvert("Pharn hoofi API khoong howjp leej"), requestResult.Attempts
            End If
        Else
            WriteRelatedInformationResponse requestResult.ResponseText, invoiceType, CLng(invoiceBuffer(invoiceIndex, 6))
        End If
        MarkGdtRetrySuccess CurrentDirectionName(), apiSource, CStr(invoiceBuffer(invoiceIndex, 0)), _
            CStr(invoiceBuffer(invoiceIndex, 3)), CStr(invoiceBuffer(invoiceIndex, 1)), _
            CStr(invoiceBuffer(invoiceIndex, 2)), stageName, requestResult.Attempts
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
    WaitGdtMilliseconds GetSleepDelayMs()
End Sub

Private Sub ProcessQueuedRelationRetries(ByVal invoiceType As Long)
    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    Dim parsed As Object, writeSucceeded As Boolean

    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "RELATIVE" Or queued.ResponseKind = "RELATED" Then
            ShowInvoiceWork UniConvert("Thuwr laji ") & LCase$(queued.ResponseKind) & ": " & queued.InvoiceNumber, True
            Set requestResult = ExecuteGdtRequest("GET", queued.Endpoint, getToken(), vbNullString, _
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
                    WriteRelatedInformationResponse requestResult.ResponseText, invoiceType, queued.TargetRow
                    writeSucceeded = True
                End If

                If writeSucceeded Then
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

Private Sub ProcessQueuedDetailRetries(ByRef detailRow As Long)
    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "DETAIL" Then
            ShowInvoiceWork UniConvert("Thuwr laji chi tieest: ") & queued.InvoiceNumber, True
            Set requestResult = ExecuteGdtRequest("GET", queued.Endpoint, getToken(), _
                vbNullString, "application/json", "application/json, text/plain, */*", False, 1, "DETAIL")
            PublishLastGdtResult requestResult
            If requestResult.Success Then
                If Me.optMua Then
                    ghiExcel_ChiTiet requestResult.ResponseText, detailRow, "mua"
                Else
                    ghiExcel_ChiTiet requestResult.ResponseText, detailRow, "ban"
                End If
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
    If Me.optMua = True Then
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
    On Error GoTo DownloadFailed
    Application.ScreenUpdating = False
    ResetGdtOperationControl
    ResetGdtRequestSession
    Set mFinalRetryQueue = New Collection
    mFinalRetryCooldownDone = False
    EnsureGdtErrorReportSheet
    ResetProgressUI
    SetProgress 1, UniConvert("DDang kieerm tra duwx lieeju ddaafu vafo..."), True
    
    'Cap ngay phai hop le: ham nay cung tao lai bang ky thoi gian cho lan tai nay
    If Not lietKeThoiGian() Then
        AbortInvalidDateRange
        Exit Sub
    End If
    
    Me.lblStatus.Visible = True
    
    If Len(Trim(Me.cboDonVi)) = 0 Then
        MsgBoxUni UniConvert("Bajn chuwa chojn DDown vij caafn tari hosa ddown."), vbCritical
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    If Len(Trim(Me.txtTuNgay)) = 0 Or Len(Trim(Me.txtDenNgay)) = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    If Len(Trim(Me.cboTTHD)) = 0 Or Len(Trim(Me.cboKQKT)) = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
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
        If Me.optMua = True Then
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
        If Me.optMua = True Then
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
    If GdtStopRequested Then
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
    bearer = cToken
    url2 = ""
    If Me.optMua = True Then
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
                search = "tdlap=ge=" & Format(arrDate(k, 1), "dd/mm/yyyy") & "T00:00:00;tdlap=le=" & Format(arrDate(k, 2), "dd/mm/yyyy") & "T23:59:59"
            Else
                search = "tdlap=ge=" & Format(arrDate(k, 1), "dd/mm/yyyy") & "T00:00:00;tdlap=le=" & Format(arrDate(k, 2), "dd/mm/yyyy") & "T23:59:59;ttxly==" & ttxly
            End If
        Else
            If ttxly = "All" Then
                search = "tdlap=ge=" & Format(arrDate(k, 1), "dd/mm/yyyy") & "T00:00:00;tdlap=le=" & Format(arrDate(k, 2), "dd/mm/yyyy") & "T23:59:59;tthai==" & tthai
            Else
                search = "tdlap=ge=" & Format(arrDate(k, 1), "dd/mm/yyyy") & "T00:00:00;tdlap=le=" & Format(arrDate(k, 2), "dd/mm/yyyy") & "T23:59:59;tthai==" & tthai & ";ttxly==" & ttxly
            End If
        End If
        url2 = url & sort & "&size=" & size & "&search=" & search
        
nextPage:
        If Not WaitForGdtControl() Then GoTo cleanup
        ReportListPageStart k, UBound(arrDate), "query", queryPageNumber
        requestStarted = Timer
        res = ApiGet(url2, ListPageLabel(k, UBound(arrDate), "query", queryPageNumber))
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
        
        If InStr(1, res, "error") > 0 Then
            ReportListPageFailure k, UBound(arrDate), "query", queryPageNumber, requestStarted
            TryParseGdtJson res, js, "query", url2, , , , , arrDate(k, 1)
            GoTo startSco
        End If
        
        If Not TryParseGdtJson(res, js, "query", url2, , , , , arrDate(k, 1)) Then
            ReportListPageFailure k, UBound(arrDate), "query", queryPageNumber, requestStarted
            GoTo startSco
        End If
        MarkGdtRetrySuccess CurrentDirectionName(), "query", vbNullString, vbNullString, vbNullString, _
            vbNullString, UniConvert("Laasy danh sasch"), LastGdtAttempts
        MarkGdtRetrySuccess CurrentDirectionName(), "query", vbNullString, vbNullString, vbNullString, _
            vbNullString, "Parse " & UniConvert("duwx lieeju"), LastGdtAttempts
        
        batchStartRow = row
        ghiExcel_TongHop res, row, loaiHD, stt
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
            arrHDChiTiet_tmp(n, 5) = ISODATE(item("tdlap"))
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
        res = ApiGet(url2, ListPageLabel(k, UBound(arrDate), "sco-query", scoPageNumber))
        If Len(res) = 0 Then
            ReportListPageFailure k, UBound(arrDate), "sco-query", scoPageNumber, requestStarted
            If GdtStopRequested Then GoTo cleanup
            errMsg = errMsg & UniConvert("Looxi tari hoas ddown toorng howjp - Tuwf masy tisnh tieefn ngafy: ") & arrDate(k, 1) & " - " & arrDate(k, 2) & vbCrLf
            QueueFinalRetry "sco-query", vbNullString, vbNullString, vbNullString, vbNullString, _
                arrDate(k, 1), UniConvert("Laasy danh sasch"), url2, "LIST"
            'If getStatus = 429 Or getStatus = 500 Then GoTo errLog
            GoTo nextDatePeriod
        End If
        
        If InStr(1, res, "error") > 0 Then
            ReportListPageFailure k, UBound(arrDate), "sco-query", scoPageNumber, requestStarted
            TryParseGdtJson res, js, "sco-query", url2, , , , , arrDate(k, 1)
            GoTo nextDatePeriod
        End If
        
        If Not TryParseGdtJson(res, js, "sco-query", url2, , , , , arrDate(k, 1)) Then
            ReportListPageFailure k, UBound(arrDate), "sco-query", scoPageNumber, requestStarted
            GoTo nextDatePeriod
        End If
        MarkGdtRetrySuccess CurrentDirectionName(), "sco-query", vbNullString, vbNullString, vbNullString, _
            vbNullString, UniConvert("Laasy danh sasch"), LastGdtAttempts
        MarkGdtRetrySuccess CurrentDirectionName(), "sco-query", vbNullString, vbNullString, vbNullString, _
            vbNullString, "Parse " & UniConvert("duwx lieeju"), LastGdtAttempts
        
        batchStartRow = row
        ghiExcel_TongHop res, row, loaiHD, stt
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
            arrHDChiTiet_tmp(n, 5) = ISODATE(item("tdlap"))
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
    'Ghi chi tiet hoa don
    For j = 0 To n - 1 'UBound(arrHDChiTiet) - 1
        If Not WaitForGdtControl() Then Exit For
        ShowInvoiceWork UniConvert("DDang tari chi tieest ") & (j + 1) & "/" & n, ((j Mod 10) = 0 Or j = n - 1)
        If arrHDChiTiet(j, 4) = 1 Then
            url_ct = "https://hoadondientu.gdt.gov.vn/api/query/invoices/detail?"
        Else
            url_ct = "https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/detail?"
        End If
        url_ct = url_ct & "nbmst=" & arrHDChiTiet(j, 0) & "&khhdon=" & arrHDChiTiet(j, 1) & "&shdon=" & arrHDChiTiet(j, 2) & "&khmshdon=" & arrHDChiTiet(j, 3)
        res = ApiGet(url_ct, "DETAIL " & (j + 1) & "/" & n)
        
        If Len(res) = 0 Then
            errMsg = errMsg & UniConvert("Looxi tari hoas ddown chi tieest: ") & arrHDChiTiet(j, 1) & "_" & arrHDChiTiet(j, 2) & UniConvert(" ngafy ") & Format(arrHDChiTiet(j, 5), "dd/mm/yyyy") & vbCrLf
            QueueFinalRetry IIf(arrHDChiTiet(j, 4) = 1, "query", "sco-query"), _
                CStr(arrHDChiTiet(j, 0)), CStr(arrHDChiTiet(j, 3)), CStr(arrHDChiTiet(j, 1)), _
                CStr(arrHDChiTiet(j, 2)), arrHDChiTiet(j, 5), UniConvert("Laasy chi tieest"), url_ct, "DETAIL"
            If GdtAuthenticationFailed Then Exit For
            GoTo nextInvoice
        End If
        If InStr(1, res, "error") > 0 Then
            TryParseGdtJson res, js, IIf(arrHDChiTiet(j, 4) = 1, "query", "sco-query"), _
                url_ct, CStr(arrHDChiTiet(j, 0)), CStr(arrHDChiTiet(j, 3)), _
                CStr(arrHDChiTiet(j, 1)), CStr(arrHDChiTiet(j, 2)), arrHDChiTiet(j, 5)
            GoTo nextInvoice
        End If
        
        If Not TryParseGdtJson(res, js, IIf(arrHDChiTiet(j, 4) = 1, "query", "sco-query"), _
            url_ct, CStr(arrHDChiTiet(j, 0)), CStr(arrHDChiTiet(j, 3)), _
            CStr(arrHDChiTiet(j, 1)), CStr(arrHDChiTiet(j, 2)), arrHDChiTiet(j, 5)) Then GoTo nextInvoice
        
        
        If Me.optMua = True Then
            ghiExcel_ChiTiet res, row_ct, "mua"
        Else
            ghiExcel_ChiTiet res, row_ct, "ban"
        End If
        MarkGdtRetrySuccess CurrentDirectionName(), IIf(arrHDChiTiet(j, 4) = 1, "query", "sco-query"), _
            CStr(arrHDChiTiet(j, 0)), CStr(arrHDChiTiet(j, 3)), CStr(arrHDChiTiet(j, 1)), _
            CStr(arrHDChiTiet(j, 2)), UniConvert("Laasy chi tieest"), LastGdtAttempts
        MarkGdtRetrySuccess CurrentDirectionName(), IIf(arrHDChiTiet(j, 4) = 1, "query", "sco-query"), _
            CStr(arrHDChiTiet(j, 0)), CStr(arrHDChiTiet(j, 3)), CStr(arrHDChiTiet(j, 1)), _
            CStr(arrHDChiTiet(j, 2)), "Parse " & UniConvert("duwx lieeju"), LastGdtAttempts
        
nextInvoice:                                    'arrHDChiTiet
        AdvanceInvoiceWork UniConvert("DDax xuwr lys chi tieest ") & (j + 1) & "/" & n
        WaitGdtMilliseconds GetSleepDelayMs()
    Next j
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
    
    If Not GdtStopRequested Then MsgBox "Xong."

cleanup:
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
    On Error GoTo EH
    Dim nbmst As String, khhdon As String, shdon As String, khmshdon As String
    Dim urlXml As String, payload As String, m As Long
    Dim apiSource As String
    Dim requestResult As clsGdtRequestResult
    
    '/ Tao folder luu file giai nen
    Dim strDate As String
    strDate = Format(Now, "_yyyymmdd_hhmmss")
    oSaveUnzipFolder = Me.txtXMLFolderPath & "FileGiaiNen" & strDate '& "\"
    MkDir oSaveUnzipFolder
    Application.Cursor = xlWait
    
    For m = 0 To soHD - 1
        If Not WaitForGdtControl() Then Exit For
        If arrHDChiTiet(m, 4) = 1 Then  'query
            urlXml = "https://hoadondientu.gdt.gov.vn/api/query/invoices/export-xml?"
            apiSource = "query"
        Else    'sco-query
            urlXml = "https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/export-xml?"
            apiSource = "sco-query"
        End If
        nbmst = arrHDChiTiet(m, 0)
        khhdon = arrHDChiTiet(m, 1)
        shdon = arrHDChiTiet(m, 2)
        khmshdon = arrHDChiTiet(m, 3)
        payload = "nbmst=" & nbmst & "&khhdon=" & khhdon & "&shdon=" & shdon & "&khmshdon=" & khmshdon
        urlXml = urlXml & payload
        ShowInvoiceWork UniConvert("DDang tari XML ") & (m + 1) & "/" & soHD, ((m Mod 10) = 0)
        Set requestResult = ExecuteGdtBinaryRequest(urlXml, bearer, GDT_MAX_RETRIES, "XML")
        PublishLastGdtResult requestResult

        If requestResult.Success Then
            SaveAndUnzipGdtXml requestResult.ResponseBody, khhdon, shdon
            MarkGdtRetrySuccess CurrentDirectionName(), apiSource, nbmst, khmshdon, khhdon, shdon, _
                UniConvert("Tari XML"), requestResult.Attempts
        ElseIf requestResult.NoXml Then
            UpsertGdtErrorReport CurrentDirectionName(), apiSource, nbmst, khmshdon, khhdon, shdon, _
                arrHDChiTiet(m, 5), UniConvert("Tari XML"), urlXml, requestResult.StatusCode, _
                requestResult.ErrorMessage, requestResult.Attempts, requestResult.RetryAfterSeconds, _
                UniConvert("Khoong cos XML")
        Else
            QueueFinalRetry apiSource, nbmst, khmshdon, khhdon, shdon, arrHDChiTiet(m, 5), _
                UniConvert("Tari XML"), urlXml, "XML"
        End If
        If requestResult.AuthenticationFailure Then Exit For
        
nextInvoice:
        AdvanceInvoiceWork UniConvert("DDax xuwr lys XML ") & (m + 1) & "/" & soHD
        WaitGdtMilliseconds GetSleepDelayMs()
    Next m
    ProcessQueuedXmlRetries

EH_Exit:
    If errMsg <> "" Then
        ghiLog errMsg
    End If
    Application.Cursor = xlDefault
    Exit Sub
EH:
    Application.Cursor = xlDefault
    MsgBoxUni "C" & ChrW(243) & " l" & ChrW(7895) & "i ph" & ChrW(225) & "t sinh trong qu" & ChrW(225) & " tr" & ChrW(236) & "nh t" & ChrW(7843) & "i h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n.", vbCritical, "L" & ChrW(7895) & "i t" & ChrW(7843) & "i h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
    MsgBox "Ma loi: " & err.Number & vbCrLf & "Noi dung: " & err.Description, vbCritical, "L" & ChrW(7895) & "i t" & ChrW(7843) & "i h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
    Resume EH_Exit
    
End Sub

Private Sub SaveAndUnzipGdtXml(ByVal responseBody As Variant, ByVal invoiceSeries As String, ByVal invoiceNumber As String)
    Dim stream As Object, filePath As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write responseBody
    filePath = Me.txtXMLFolderPath & invoiceSeries & "_" & invoiceNumber & ".zip"
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
    Unzip filePath, oSaveUnzipFolder
End Sub

Private Sub ProcessQueuedXmlRetries()
    Dim queued As clsGdtRetryItem, requestResult As clsGdtRequestResult
    If mFinalRetryQueue Is Nothing Then Exit Sub
    EnsureFinalRetryCooldown
    For Each queued In mFinalRetryQueue
        If Not WaitForGdtControl() Then Exit Sub
        If queued.ResponseKind = "XML" Then
            ShowInvoiceWork UniConvert("Thuwr laji XML: ") & queued.InvoiceNumber, True
            Set requestResult = ExecuteGdtBinaryRequest(queued.Endpoint, getToken(), 1, "XML")
            PublishLastGdtResult requestResult
            If requestResult.Success Then
                SaveAndUnzipGdtXml requestResult.ResponseBody, queued.InvoiceSeries, queued.InvoiceNumber
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
            WaitGdtMilliseconds GetSleepDelayMs()
            If requestResult.AuthenticationFailure Then Exit For
        End If
    Next queued
End Sub



