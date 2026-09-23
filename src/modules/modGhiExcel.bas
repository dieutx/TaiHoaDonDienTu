Attribute VB_Name = "modGhiExcel"
Option Explicit



Sub ghiExcel_TongHop(ByVal jsonText As String, row As Long, ByVal loaiHD As Long, sSTT As Long)
    
    Dim jsTH As Object, item As Object, itm As Object, subItms As Object, subItms2 As Object, l As Long
    Dim ws As Worksheet
    Dim arrCol, arrColName
    
    'On Error Resume Next
    
    arrCol = Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, _
        12, 13, 14, 15, 16, 17, 42, 43, 44, _
        47, 48, 49, 50, 51, 54)
    arrColName = Array("tlhdon", "khmshdon", "khhdon", "shdon", "tdlap", "dvtte", "tgia", "nbten", "nbmst", "nbdchi", _
        "nky", "mhdon", "ncma", "nmten", "nmmst", "nmdchi", "tgtcthue", "tgtkcthue", "tgtthue", _
        "ttcktmai", "tgtkhac", "tgtttbso", "tgtttbchu", "gchu", "msttcgp")
    
    Select Case loaiHD
        Case 1
            Set ws = ThisWorkbook.Sheets("TongHopHD_Mua")
        Case 2
            Set ws = ThisWorkbook.Sheets("TongHopHD_Ban")
    End Select
    EnsureRelatedInvoiceHeaders ws
    
    Set jsTH = JsonConverter.ParseJSON(jsonText)
    
    Dim dataArray As Object, i As Long, v As Variant
    Set dataArray = jsTH("datas")
    For l = 1 To jsTH("datas").count
        'On Error Resume Next    'Bo qua cac loi khi khong tim thay cac item trong json,item = null
        '/Cot STT
        ws.Cells(row, 1).Value = sSTT
        
        For i = 0 To UBound(arrCol)
            v = dataArray(l)(arrColName(i))
            If v = "" Or IsNull(v) Then GoTo next_col
            Select Case arrColName(i)
                Case "ncma", "nky", "ncnhat", "ntao", "ntnhan", "tdlap"
                    ws.Cells(row, arrCol(i)).Value = ISODATE(Format(v, "yyyy-mm-dd"))
                Case Else
                    ws.Cells(row, arrCol(i)).Value = v
            End Select
next_col:
        Next i
        
        '/Ghi cac truong dac biet
        If Not dataArray(l)("thttltsuat") Is Nothing Then
            Set subItms = dataArray(l)("thttltsuat")
            If subItms.count > 0 Then
                Dim c As Long
                For c = 0 To subItms.count - 1  'Co tat ca 6 cot tsuat: KKKCT, KCT, 0%, 5%, 8%, 10%
                    Set subItms2 = subItms(c + 1)
                    ws.Cells(row, 18 + (4 * c)).Value = subItms2("tsuat")
                    ws.Cells(row, 19 + (4 * c)).Value = subItms2("thtien")
                    ws.Cells(row, 20 + (4 * c)).Value = subItms2("tthue")
                    ws.Cells(row, 21 + (4 * c)).Value = subItms2("gttsuat")
                Next c
            End If
        End If
        
        If Not IsEmpty(dataArray(l)("thttlphi")) Then
            If Not dataArray(l)("thttlphi") Is Nothing Then
                Set subItms = dataArray(l)("thttlphi")
                If subItms.count > 0 Then
                    Set subItms2 = subItms(1)
                    ws.Cells(row, 45).Value = subItms2("tlphi")
                    ws.Cells(row, 46).Value = subItms2("tphi")
                End If
            End If
        End If
        
        ws.Cells(row, 52).Value = arrTrangThai(CInt(dataArray(l)("tthai")) + 1, 1)  'Bo qua dong "Tat ca"
        ws.Cells(row, 53).Value = arrKQKTHoaDon(CInt(dataArray(l)("ttxly")) + 2, 1)
        WriteRelatedInvoiceInfo dataArray(l), ws, row
        
        '//LAY LINK TRA CUU
        Dim msttcgp As String, mst As String, mccqt As String
        mst = ws.Cells(row, 10).Value: mccqt = ws.Cells(l + ROWSTART, 13).Value
        If Len(dataArray(l)("msttcgp")) > 0 Then
            msttcgp = dataArray(l)("msttcgp")
            Select Case msttcgp
                Case "0100684378"   'VNPT
                    If Len(mccqt) > 0 Then
                        ws.Cells(row, 55).Value = "https://" & mst & "-tt78.vnpt-invoice.com.vn/?strFkey=" & mccqt
                        ws.Cells(row, 56).Value = mccqt
                    Else
                        ws.Cells(row, 55).Value = "Khong co link tra cuu"
                    End If
                Case "0105987432"
                    ws.Cells(row, 55).Value = "https://" & mst & "hd.easyinvoice.com.vn"
                Case "0101360697"   'BKAV
                    Debug.Print dataArray(l)("id");
                    If Len(dataArray(l)("id") & "") > 0 Then
                        If Len(mccqt) > 0 Then
                            ws.Cells(row, 55).Value = "https://van.ehoadon.vn/Lookup?InvoiceGUID=" & dataArray(l)("id")
                            ws.Cells(row, 56).Value = dataArray(l)("id")
                        Else
                            ws.Cells(row, 55).Value = ""
                        End If
                    End If
                Case Else
                    ws.Cells(row, 55).Value = dicLink.item(msttcgp)
            End Select
            '------------------------------------------------------------------/
            
        Else    'truong hop khong co msttcgp
            Select Case mst
                Case "0110269067", "0110269067-002" 'hoa don Xanh SM
                    ws.Cells(row, 55).Value = dicLink.item("0110269067")
                Case Else
                    ws.Cells(row, 55).Value = "Khong co link tra cuu"
            End Select
        End If
        
        '/LAY MA TRA CUU
        'Con truong hop cttkhac khong exists!
        If Not dataArray(l)("cttkhac") Is Nothing Then
            Set subItms = dataArray(l)("cttkhac")
            For Each itm In subItms
                'If ws.Cells(row, 56).Value = "" Then   'Tru cac truong hop BKAV, VNPT... da co ms tra cuu
                 If Trim(ws.Cells(row, 56).Value) <> "" Then GoTo skipMTC 'Tru cac truong hop BKAV, VNPT... da co ms tra cuu
                    If dicTenCotTC.Exists(itm("ttruong")) Then
                        ws.Cells(row, 56).Value = itm("dlieu")
                        Exit For
                    End If
                'End If
            Next
        End If
        
        If Trim(ws.Cells(row, 56).Value) <> "" Then GoTo skipMTC
        If Not dataArray(l)("ttkhac") Is Nothing Then
            Set subItms = dataArray(l)("ttkhac")
            For Each itm In subItms
                If dicTenCotTC.Exists(itm("ttruong")) Then
                    ws.Cells(row, 56).Value = itm("dlieu")
                    Exit For
                End If
            Next
        End If
        '-------------------------------------/
skipMTC:
        'Tang so tt va so thu thu dong
        sSTT = sSTT + 1
        row = row + 1
        
next_invoice:
    Next l
    
End Sub

Private Sub EnsureRelatedInvoiceHeaders(ByVal ws As Worksheet)
    ws.Cells(2, 57).Value = UniConvert("Chuooxi hosa ddown lieen quan")
    ws.Cells(2, 58).Value = UniConvert("Loaji hosa ddown goosc")
    ws.Cells(2, 59).Value = UniConvert("Kys hieeju maaxu soos HDD goosc")
    ws.Cells(2, 60).Value = UniConvert("Kys hieeju HDD goosc")
    ws.Cells(2, 61).Value = UniConvert("Soos HDD goosc")
    ws.Cells(2, 62).Value = UniConvert("Ngafy laajp HDD goosc")
    ws.Cells(2, 63).Value = UniConvert("Ghi chus HDD goosc")
    ws.Cells(2, 64).Value = UniConvert("Cos HDD lieen quan")
    ws.Cells(2, 65).Value = UniConvert("Thoong tin lieen quan")
End Sub

Private Sub WriteRelatedInvoiceInfo(ByVal invoice As Object, ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim invoiceStatus As Long
    Dim relationSummary As String
    Dim relatedItems As Object, relatedItem As Variant

    invoiceStatus = CLng(Val(SafeJsonText(invoice, "tthai")))
    If invoiceStatus < 2 Or invoiceStatus > 6 Then Exit Sub

    ws.Cells(targetRow, 58).Value = SafeJsonText(invoice, "lhdgoc")
    ws.Cells(targetRow, 59).Value = SafeJsonText(invoice, "khmshdgoc")
    ws.Cells(targetRow, 60).Value = SafeJsonText(invoice, "khhdgoc")
    ws.Cells(targetRow, 61).Value = SafeJsonText(invoice, "shdgoc")
    ws.Cells(targetRow, 63).Value = SafeJsonText(invoice, "gchdgoc")
    ws.Cells(targetRow, 64).Value = SafeJsonText(invoice, "tthdclquan")

    If Len(SafeJsonText(invoice, "tdlhdgoc")) > 0 Then
        ws.Cells(targetRow, 62).Value = ISODATE(SafeJsonText(invoice, "tdlhdgoc"))
        ws.Cells(targetRow, 62).NumberFormat = "dd/mm/yyyy"
    End If

    relationSummary = BuildRelatedInvoiceKey(invoice, True)
    If invoice.Exists("hdonLquans") Then
        If IsObject(invoice("hdonLquans")) Then
            Set relatedItems = invoice("hdonLquans")
            If Not relatedItems Is Nothing Then
                For Each relatedItem In relatedItems
                    If IsObject(relatedItem) Then AppendRelatedInvoiceKey relationSummary, BuildRelatedInvoiceKey(relatedItem, False)
                Next relatedItem
            End If
        End If
    End If
    ws.Cells(targetRow, 57).Value = relationSummary
    ws.Cells(targetRow, 57).WrapText = True
End Sub

Private Function BuildRelatedInvoiceKey(ByVal invoice As Object, ByVal useOriginalFields As Boolean) As String
    Dim templateCode As String, invoiceSeries As String, invoiceNumber As String

    If useOriginalFields Then
        templateCode = SafeJsonText(invoice, "khmshdgoc")
        invoiceSeries = SafeJsonText(invoice, "khhdgoc")
        invoiceNumber = SafeJsonText(invoice, "shdgoc")
    Else
        templateCode = SafeJsonText(invoice, "khmshdon")
        invoiceSeries = SafeJsonText(invoice, "khhdon")
        invoiceNumber = SafeJsonText(invoice, "shdon")
        If Len(templateCode) = 0 Then templateCode = SafeJsonText(invoice, "khmshdgoc")
        If Len(invoiceSeries) = 0 Then invoiceSeries = SafeJsonText(invoice, "khhdgoc")
        If Len(invoiceNumber) = 0 Then invoiceNumber = SafeJsonText(invoice, "shdgoc")
    End If

    If Len(templateCode & invoiceSeries & invoiceNumber) = 0 Then Exit Function
    BuildRelatedInvoiceKey = templateCode & " | " & invoiceSeries & " | " & invoiceNumber
End Function

Private Sub AppendRelatedInvoiceKey(ByRef summary As String, ByVal relatedKey As String)
    If Len(relatedKey) = 0 Then Exit Sub
    If InStr(1, summary, relatedKey, vbTextCompare) > 0 Then Exit Sub
    If Len(summary) > 0 Then summary = summary & vbCrLf
    summary = summary & relatedKey
End Sub

Private Function SafeJsonText(ByVal source As Object, ByVal keyName As String) As String
    Dim value As Variant
    On Error GoTo MissingValue
    If source Is Nothing Then Exit Function
    If Not source.Exists(keyName) Then Exit Function
    If IsObject(source(keyName)) Then Exit Function
    value = source(keyName)
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    SafeJsonText = CStr(value)
MissingValue:
End Function

Public Sub WriteRelativeInvoiceData( _
    ByVal relatedItems As Object, _
    ByVal loaiHD As Long, _
    ByVal targetRow As Long, _
    ByVal currentTemplateCode As String, _
    ByVal currentInvoiceSeries As String, _
    ByVal currentInvoiceNumber As String)

    Dim ws As Worksheet
    Dim item As Variant
    Dim reportText As String, reportLine As String
    Dim itemIndex As Long, sourceIndex As Long

    Select Case loaiHD
        Case 1: Set ws = ThisWorkbook.Sheets("TongHopHD_Mua")
        Case 2: Set ws = ThisWorkbook.Sheets("TongHopHD_Ban")
        Case Else: Exit Sub
    End Select

    itemIndex = 0
    If Not relatedItems Is Nothing Then
        'API relative tra ve chuoi tu hoa don gan hoa don goc den hoa don moi nhat.
        'Bao cao can hien thi tu hoa don moi nhat nguoc ve hoa don dang tra cuu.
        For sourceIndex = relatedItems.Count To 1 Step -1
            Set item = relatedItems(sourceIndex)
            If IsObject(item) Then
                itemIndex = itemIndex + 1
                reportLine = CStr(itemIndex) & ". " & UniConvert("Hosa ddown cos lieen quan") & " | " & _
                    SafeJsonText(item, "khmshdon") & " | " & SafeJsonText(item, "khhdon") & " | " & _
                    SafeJsonText(item, "shdon")
                If Len(BuildOriginalInvoiceDescription(item)) > 0 Then _
                    reportLine = reportLine & " | " & BuildOriginalInvoiceDescription(item)
                AppendReportLine reportText, reportLine
            End If
        Next sourceIndex
    End If

    itemIndex = itemIndex + 1
    reportLine = CStr(itemIndex) & ". " & UniConvert("Hosa ddown ddang tra cuwsu") & " | " & _
        currentTemplateCode & " | " & currentInvoiceSeries & " | " & currentInvoiceNumber
    AppendReportLine reportText, reportLine

    ws.Cells(targetRow, 57).Value = reportText
    ws.Cells(targetRow, 57).WrapText = True
    ws.Rows(targetRow).AutoFit
    If ws.Rows(targetRow).RowHeight > 150 Then ws.Rows(targetRow).RowHeight = 150
End Sub

Public Sub WriteRelatedInformationResponse( _
    ByVal responseText As String, _
    ByVal loaiHD As Long, _
    ByVal targetRow As Long)

    Dim ws As Worksheet
    Dim normalized As String
    Dim parsed As Object
    Dim noticeSummary As String
    Dim hasNoticeContainer As Boolean

    Select Case loaiHD
        Case 1: Set ws = ThisWorkbook.Sheets("TongHopHD_Mua")
        Case 2: Set ws = ThisWorkbook.Sheets("TongHopHD_Ban")
        Case Else: Exit Sub
    End Select

    normalized = Trim$(responseText)
    If Len(normalized) = 0 Or normalized = "[]" Or normalized = "{}" Or LCase$(normalized) = "null" Then
        normalized = UniConvert("Khoong cos thoong tin lieen quan")
    Else
        On Error GoTo KeepRawJson
        Set parsed = JsonConverter.ParseJSON(normalized)
        noticeSummary = BuildRelatedNoticeSummary(parsed, hasNoticeContainer)
        If Len(noticeSummary) > 0 Then
            normalized = noticeSummary
        ElseIf hasNoticeContainer Then
            normalized = UniConvert("Khoong cos thoong tin lieen quan")
        Else
            normalized = JsonConverter.ConvertToJson(parsed, 2)
        End If
KeepRawJson:
        On Error GoTo 0
    End If

    If Len(normalized) > 32000 Then normalized = Left$(normalized, 32000)
    ws.Cells(targetRow, 65).Value = normalized
    ws.Cells(targetRow, 65).WrapText = True
    ws.Rows(targetRow).AutoFit
    If ws.Rows(targetRow).RowHeight > 150 Then ws.Rows(targetRow).RowHeight = 150
End Sub

Public Sub WriteNoRelativeInvoiceData(ByVal loaiHD As Long, ByVal targetRow As Long)
    Dim ws As Worksheet
    Select Case loaiHD
        Case 1: Set ws = ThisWorkbook.Sheets("TongHopHD_Mua")
        Case 2: Set ws = ThisWorkbook.Sheets("TongHopHD_Ban")
        Case Else: Exit Sub
    End Select
    ws.Cells(targetRow, 57).Value = UniConvert("Khoong cos thoong tin hieen thij")
End Sub

Private Function BuildRelatedNoticeSummary( _
    ByVal responseObject As Object, _
    ByRef hasNoticeContainer As Boolean) As String
    Dim notices As Object
    Dim notice As Variant
    Dim noticeLine As String
    Dim summaryText As String

    On Error GoTo NoNoticeData
    If responseObject Is Nothing Then Exit Function
    If responseObject.Exists("mtthdtbssrs") Then
        hasNoticeContainer = True
        If IsObject(responseObject("mtthdtbssrs")) Then Set notices = responseObject("mtthdtbssrs")
    ElseIf responseObject.Exists("hdtbssrses") Then
        hasNoticeContainer = True
        If IsObject(responseObject("hdtbssrses")) Then Set notices = responseObject("hdtbssrses")
    End If
    If notices Is Nothing Then Exit Function

    For Each notice In notices
        If IsObject(notice) Then
            noticeLine = BuildRelatedNoticeLine(notice)
            AppendReportLine summaryText, noticeLine
        End If
    Next notice
    BuildRelatedNoticeSummary = summaryText
    Exit Function

NoNoticeData:
    BuildRelatedNoticeSummary = vbNullString
End Function

Private Function BuildRelatedNoticeLine(ByVal notice As Object) As String
    Dim noticeName As String, noticeDate As String
    Dim noticeReason As String, noticeNature As String
    Dim receiveResult As String, receiveCode As String

    noticeName = SafeJsonText(notice, "ten")
    noticeDate = SafeJsonText(notice, "ngay")
    noticeReason = SafeJsonText(notice, "ldo")
    noticeNature = RelatedNoticeNature(SafeJsonText(notice, "loai"))
    receiveCode = SafeJsonText(notice, "kqtnhan")

    If Len(noticeDate) > 0 Then noticeDate = FormatRelatedNoticeDate(noticeDate)
    If Len(receiveCode) > 0 Then
        If CLng(Val(receiveCode)) = 1 Then
            receiveResult = UniConvert("Cow quan thuees tieesp nhaajn.")
        Else
            receiveResult = UniConvert("Cow quan thuees khoong tieesp nhaajn.")
        End If
    End If

    BuildRelatedNoticeLine = UniConvert("Hosa ddown cos ") & noticeName
    If Len(noticeDate) > 0 Then BuildRelatedNoticeLine = BuildRelatedNoticeLine & UniConvert(" ngafy ") & noticeDate
    If Len(noticeNature) > 0 Then BuildRelatedNoticeLine = BuildRelatedNoticeLine & UniConvert(". Tisnh chaast ") & noticeNature
    If Len(noticeReason) > 0 Then BuildRelatedNoticeLine = BuildRelatedNoticeLine & UniConvert(", lys do ") & noticeReason
    If Right$(BuildRelatedNoticeLine, 1) <> "." Then BuildRelatedNoticeLine = BuildRelatedNoticeLine & "."
    If Len(receiveResult) > 0 Then BuildRelatedNoticeLine = BuildRelatedNoticeLine & " " & receiveResult
End Function

Private Function FormatRelatedNoticeDate(ByVal isoText As String) As String
    Dim parsedDate As Date
    On Error GoTo InvalidDate

    parsedDate = DateSerial(CInt(Left$(isoText, 4)), CInt(Mid$(isoText, 6, 2)), CInt(Mid$(isoText, 9, 2)))
    If Len(isoText) >= 19 Then
        parsedDate = parsedDate + TimeSerial(CInt(Mid$(isoText, 12, 2)), _
            CInt(Mid$(isoText, 15, 2)), CInt(Mid$(isoText, 18, 2)))
    End If
    If Right$(Trim$(isoText), 1) = "Z" Then parsedDate = UTCToLocalTime(parsedDate)
    FormatRelatedNoticeDate = Format$(parsedDate, "dd/mm/yyyy")
    Exit Function

InvalidDate:
    FormatRelatedNoticeDate = isoText
End Function

Private Function RelatedNoticeNature(ByVal natureCode As String) As String
    Select Case CLng(Val(natureCode))
        Case 1: RelatedNoticeNature = UniConvert("Hury")
        Case 2: RelatedNoticeNature = UniConvert("DDieefu chirnh")
        Case 3: RelatedNoticeNature = UniConvert("Thay thees")
        Case 4: RelatedNoticeNature = UniConvert("Giari trifnh")
        Case Else: RelatedNoticeNature = natureCode
    End Select
End Function

Private Function BuildOriginalInvoiceDescription(ByVal item As Object) As String
    Dim originalTemplate As String, originalSeries As String, originalNumber As String
    Dim originalDescription As String
    Dim invoiceStatus As Long

    originalTemplate = SafeJsonText(item, "khmshdgoc")
    originalSeries = SafeJsonText(item, "khhdgoc")
    originalNumber = SafeJsonText(item, "shdgoc")
    If Len(originalTemplate & originalSeries & originalNumber) = 0 Then Exit Function
    originalDescription = originalTemplate & UniConvert(", kys hieeju hosa ddown ") & _
        originalSeries & UniConvert(", soos hosa ddown ") & originalNumber

    invoiceStatus = CLng(Val(SafeJsonText(item, "tthai")))
    Select Case invoiceStatus
        Case 2, 4
            BuildOriginalInvoiceDescription = UniConvert("Thay thees cho hosa ddown cos kys hieeju maaxu soos ") & originalDescription
        Case 3, 5
            BuildOriginalInvoiceDescription = UniConvert("DDieefu chirnh cho hosa ddown cos kys hieeju maaxu soos ") & originalDescription
        Case 6
            BuildOriginalInvoiceDescription = UniConvert("Hury/lieen quan ddeesn hosa ddown cos kys hieeju maaxu soos ") & originalDescription
        Case Else
            BuildOriginalInvoiceDescription = UniConvert("Lieen quan ddeesn hosa ddown cos kys hieeju maaxu soos ") & originalDescription
    End Select
End Function

Private Sub AppendReportLine(ByRef reportText As String, ByVal reportLine As String)
    If Len(reportLine) = 0 Then Exit Sub
    If Len(reportText) > 0 Then reportText = reportText & vbCrLf
    reportText = reportText & reportLine
End Sub

Sub ghiExcel_ChiTiet(jsonText As String, row_ct As Long, loaiHD As String)
    Dim jsCT As Object, subItms As Object, subItms2 As Object, item As Object, itm As Object, col As Long, c As Long, sumRow As Long, tongthueCT As Double
    Dim ws As Worksheet
    Dim arrCol, arrColName, arrCol_detail, arrColName_detail
    
    On Error Resume Next
    
    sumRow = row_ct
    tongthueCT = 0
    
    '/Dung chung cho tat ca cac dong hang hoa
    arrCol = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 31)
    arrColName = Array("khhdon", "shdon", "tdlap", "dvtte", "tgia", "nbten", "nbmst", "nbdchi", "nky", "mhdon", "ncma", "nmten", "nmmst", "nmdchi", "msttcgp")
    
    '/Nhieu dong hang hoa
    arrCol_detail = Array(15, 16, 17, 18, 19, 20, 21, 22, 23, 24, _
        25, 26, 27, 28)
    arrColName_detail = Array("stt", "tchat", "mhhdvu", "ten", "dvtinh", "sluong", "dgia", "tlckhau", "stckhau", "ltsuat", _
        "tsuat", "thtien", "tthue", "thtcthue")
    
    'Ten worksheet ghi du lieu chi tiet
    Select Case loaiHD
        Case "mua"
            Set ws = ThisWorkbook.Sheets("ChiTietHD_Mua")
        Case "ban"
            Set ws = ThisWorkbook.Sheets("ChiTietHD_Ban")
    End Select
    
    Set jsCT = JsonConverter.ParseJSON(jsonText)
    
    '/Ghi thong tin chung
    For col = 0 To UBound(arrCol)
        If jsCT.Exists(arrColName(col)) Then
            If Not IsNull(jsCT(arrColName(col))) Then
                'Dinh dang du lieu dang Date
                Select Case arrColName(col)
                    Case "ncma", "nky", "ncnhat", "ntao", "ntnhan", "tdlap"
                        ws.Cells(row_ct, col + 1).Value = CStr(ISODATE(Format(jsCT(arrColName(col)), "yyyy-mm-dd")))
                    Case Else
                        ws.Cells(row_ct, col + 1).Value = jsCT(arrColName(col))
                End Select
            End If
        End If
    Next
    
    '/Lay link tra cuu
    Dim msttcgp As String
    ws.Cells(row_ct, 31).Value = jsCT("msttcgp")
    msttcgp = ws.Cells(row_ct, 31).Value
    If msttcgp <> "" Then   'Co MSTTCGP
        Dim mst As String, mccqt As String
        mst = ws.Cells(row_ct, 7).Value: mccqt = ws.Cells(row_ct, 10).Value
        Select Case msttcgp
            Case "0100684378"   'VNPT
                If mccqt <> "" Then
                    ws.Cells(row_ct, 32).Value = "https://" & mst & "-tt78.vnpt-invoice.com.vn/?strFkey=" & mccqt
                Else
                    ws.Cells(row_ct, 32).Value = UniConvert("Khoong cos MCCQT")
                End If
            Case "0101360697"   'BKAV
                ws.Cells(row_ct, 32).Value = "https://van.ehoadon.vn/Lookup?InvoiceGUID= [DLHDon Id]" '& .Offset(, 0).Value
            Case "0105987432"
                ws.Cells(row_ct, 32).Value = "https://" & mst & "hd.easyinvoice.com.vn"
            Case Else
                ws.Cells(row_ct, 32).Value = dicLink.item(msttcgp)
        End Select
        
        '/Lay ma tra cuu
        'Con truong hop cttkhac khong exists!
        If Not jsCT("cttkhac") Is Nothing Then
            Set subItms = jsCT("cttkhac")
            For Each itm In subItms
                If dicTenCotTC.Exists(itm("ttruong")) Then
                    ws.Cells(row_ct, 33).Value = itm("dlieu")
                    Exit For
                End If
            Next
        End If
        
        If Not jsCT("ttkhac") Is Nothing Then
            Set subItms = jsCT("ttkhac")
            For Each itm In subItms
                If dicTenCotTC.Exists(itm("ttruong")) Then
                    ws.Cells(row_ct, 33).Value = itm("dlieu")
                    Exit For
                End If
            Next
        End If
        '-------------------------------------/
        
    Else    'truong hop khong co msttcgp
        ws.Cells(row_ct, 32).Value = "Khong co link tra cuu"
    End If
    
    '/Ghi thong tin tung ma HHDVu
    For Each item In jsCT("hdhhdvu")
        For col = 0 To UBound(arrCol_detail)
            On Error Resume Next    'bo qua loi khong co ten cot trong json
            'Cac cot con lai
            If col = 10 Then
                Select Case item(arrColName_detail(col - 1))
                    Case "KKKNT"
                        ws.Cells(row_ct, col + 15).Value = "KKKNT"
                    Case "KCT"
                        ws.Cells(row_ct, col + 15).Value = "KCT"
                    Case Else
                        ws.Cells(row_ct, col + 15).Value = item(arrColName_detail(col))
                End Select
            Else
                ws.Cells(row_ct, col + 15).Value = item(arrColName_detail(col))
            End If
        Next
        
        Dim arrTThue As Variant, arrThTiencoVAT As Variant
        arrTThue = Split(Sheets("LinkTraCuu").Range("O14"), ",")
        arrThTiencoVAT = Split(Sheets("LinkTraCuu").Range("O15"), ",")
        
        If item.Exists("ttkhac") Then
            If Not item("ttkhac") Is Nothing Then
                Dim ttruong As String
                Set subItms = item("ttkhac")
                If subItms.count > 0 Then
                    For c = 1 To subItms.count  'Co tat ca 3 cot
                        Set subItms2 = subItms(c)
                        ttruong = subItms2("ttruong")
                        If isInArray(ttruong, arrTThue) Then
                            ws.Cells(row_ct, 27).Value = subItms2("dlieu")
                            tongthueCT = tongthueCT + Val(subItms2("dlieu"))
                        ElseIf isInArray(ttruong, arrThTiencoVAT) Then
                            ws.Cells(row_ct, 28).Value = subItms2("dlieu")
                        End If
                    Next
                End If
            End If
        End If
        
        If ws.Cells(row_ct, 27).Value = 0 Then   'Khong co gia tri cho cot tthueVAT
            ws.Cells(row_ct, 27).Value = Val(ws.Cells(row_ct, 25).Value) * ws.Cells(row_ct, 26).Value
            ws.Cells(row_ct, 28).Value = ws.Cells(row_ct, 26).Value + ws.Cells(row_ct, 27).Value
        End If
        If ws.Cells(row_ct, 28).Value = 0 Then  'Khong co gia tri cho cot THTiencoVAT
            ws.Cells(row_ct, 28).Value = ws.Cells(row_ct, 26).Value + ws.Cells(row_ct, 27).Value
        End If
        
        '/Dung cho hd khong nhan ma loai 2: chi tiet hhdv ("tthue"= #,##)
        If item.Exists("tthue") Then
            If Not IsNull(item("tthue")) Then
                tongthueCT = item("tthue")
            End If
        End If
        '------------------------------/
        
        row_ct = row_ct + 1
    Next
    
    '// So sanh tien thue chi tiet va tong hop
    If jsCT.Exists("tgtthue") Then
        If Not IsNull(jsCT("tgtthue")) Then
            ws.Cells(sumRow, 29).Value = jsCT("tgtthue")
        End If
    End If
    
    If tongthueCT <> jsCT("tgtthue") Then
        ws.Cells(sumRow, 30).Value = "Kiem tra lai tien thue: [" & Format(tongthueCT, "standard") & "] <> [" & Format(jsCT("tgtthue"), "standard") & "]"
        ws.Cells(sumRow, 30).Font.Color = vbRed
    Else
        ws.Cells(sumRow, 30).Value = "Tien thue ok"
        ws.Cells(sumRow, 30).Font.Color = vbBlue
    End If
    
    '---------------------------------------/
    
    '/Copy thong tin chung
    If row_ct > sumRow Then
        CopyThongTinChung_CT ws.name, sumRow, row_ct - 1
    End If
    
End Sub

Sub LinkTraCuu()
    Dim n As Long, lr As Long, arrLinks As Variant
    
    Set dicLink = CreateObject("Scripting.Dictionary")
    With Sheets("LinkTraCuu")
        lr = .Range("B" & .Rows.count).End(xlUp).row
    End With
    arrLinks = Sheets("LinkTraCuu").Range("B2:E" & lr).Value
    
    For n = 1 To UBound(arrLinks)
        'Debug.Print arrLinks(n, 1), arrLinks(n, 3), , arrLinks(n, 4)
        dicLink(arrLinks(n, 1)) = arrLinks(n, 3)
    Next n
    Erase arrLinks
End Sub

Sub tenCotTraCuu()
    Dim arrVals(), m As Long
    
    Set dicTenCotTC = CreateObject("Scripting.Dictionary")
    arrVals = Sheets("LinkTraCuu").Range("E2:E" & Sheets("LinkTraCuu").Cells(Rows.count, "B").End(xlUp).row).Value
    For m = 1 To UBound(arrVals)
        If Len(arrVals(m, 1)) > 0 Then
            On Error Resume Next
            dicTenCotTC.Add arrVals(m, 1), "ttruong"
        End If
    Next m
    Erase arrVals
'    For m = 0 To dicTenCotTC.Count - 1
'        Debug.Print dicTenCotTC.Keys()(m), dicTenCotTC.Items()(m)
'    Next m
'    dicTenCotTC.RemoveAll
End Sub

Sub CopyThongTinChung_CT(shtName As String, frow As Long, lrow As Long)
    Dim ws As Worksheet, sourceRange As Range, targetRange As Range
    
    Set ws = ThisWorkbook.Sheets(shtName)
    Set sourceRange = ws.Range("A" & frow & ":N" & frow)
    Set targetRange = ws.Range("A" & frow + 1 & ":N" & lrow)
    sourceRange.Copy Destination:=targetRange
    ' Clear the clipboard
    Application.CutCopyMode = False
    'MsgBox "Copy thanh cong", vbInformation
End Sub

