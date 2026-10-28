Attribute VB_Name = "modGdtPdf"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existingName As LongPtr, ByVal newName As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32" (ByVal existingName As Long, ByVal newName As Long, ByVal flags As Long) As Long
#End If

Public Function GdtPdfEscape(ByVal value As String) As String
    value = Replace$(value, "&", "&amp;")
    value = Replace$(value, "<", "&lt;")
    value = Replace$(value, ">", "&gt;")
    value = Replace$(value, Chr$(34), "&quot;")
    value = Replace$(value, "'", "&#39;")
    ' XML text must never introduce another template token.
    value = Replace$(Replace$(value, "{", "&#123;"), "}", "&#125;")
    GdtPdfEscape = value
End Function

Private Sub PdfPut(ByRef html As String, ByVal key As String, ByVal value As String)
    html = Replace$(html, "{{" & key & "}}", GdtPdfEscape(value))
End Sub

Private Function PdfNumber(ByVal node As Object, ByVal path As String) As String
    Dim value As String, formatted As String
    value = GdtXmlText(node, path)
    If Len(Trim$(value)) = 0 Then Exit Function
    formatted = Format$(GdtXmlNumber(value, path), "#,##0.################")
    formatted = Replace$(formatted, Application.International(xlThousandsSeparator), "|")
    formatted = Replace$(formatted, Application.International(xlDecimalSeparator), ",")
    PdfNumber = Replace$(formatted, "|", ".")
    If Right$(PdfNumber, 1) = "," Then PdfNumber = Left$(PdfNumber, Len(PdfNumber) - 1)
End Function

Private Function PdfCell(ByVal value As String, Optional ByVal numeric As Boolean = False) As String
    PdfCell = "<td class='" & IIf(numeric, "tx-right", "tx-left") & "'>" & GdtPdfEscape(value) & "</td>"
End Function

Public Function GdtPdfIdentity(ByVal seller As String, ByVal template As String, ByVal series As String, ByVal number As String) As String
    ' API invoice numbers are numeric; XML can pad them with leading zeroes.
    ' Seller tax IDs and series remain exact strings.
    GdtPdfIdentity = GdtInvoiceIdentity(seller, PdfUnpadDigits(template), series, PdfUnpadDigits(number))
End Function

Private Function PdfUnpadDigits(ByVal value As String) As String
    Dim index As Long
    If Len(value) = 0 Then Exit Function
    For index = 1 To Len(value)
        If Not Mid$(value, index, 1) Like "#" Then PdfUnpadDigits = value: Exit Function
    Next index
    Do While Len(value) > 1 And Left$(value, 1) = "0"
        value = Mid$(value, 2)
    Loop
    PdfUnpadDigits = value
End Function

Public Function GdtBuildInvoiceHtml(ByVal xmlPath As String, Optional ByVal expectedIdentity As String = "") As String
    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.validateOnParse = False: doc.resolveExternals = False
    doc.setProperty "ProhibitDTD", True
    If Not doc.Load(xmlPath) Then Err.Raise 5, , "Invalid invoice XML"
    GdtBuildInvoiceHtml = PdfRenderDocument(doc, expectedIdentity, False)
End Function

Private Function PdfRenderDocument(ByVal doc As Object, ByVal expectedIdentity As String, ByVal fromApi As Boolean) As String
    Dim invoice As Object, node As Object, nodes As Object
    Dim html As String, rows As String, taxes As String, title As String, value As String, identity As String
    Set invoice = doc.SelectSingleNode(GdtXmlPath(doc, "/HDon/DLHDon"))
    If invoice Is Nothing Then Err.Raise 5, , "Missing HDon/DLHDon"
    If Len(GdtXmlText(invoice, "TTChung/SHDon")) = 0 Then Err.Raise 5, , "Missing invoice number"
    If Len(expectedIdentity) > 0 Then
        identity = GdtPdfIdentity(GdtXmlText(invoice, "NDHDon/NBan/MST"), GdtXmlText(invoice, "TTChung/KHMSHDon"), _
            GdtXmlText(invoice, "TTChung/KHHDon"), GdtXmlText(invoice, "TTChung/SHDon"))
        If identity <> expectedIdentity Then Err.Raise 5, , "PDF XML invoice identity mismatch"
    End If
    html = Replace$(GdtPdfTemplate(), "{{CSS}}", GdtPdfCss())
    title = GdtXmlText(invoice, "TTChung/THDon")
    If Len(title) = 0 Then title = UniConvert("HOSA DDOWN")
    PdfPut html, "TITLE", title
    value = GdtXmlText(invoice, "TTChung/NLap")
    If Len(value) > 0 Then value = Format$(ISODateValue(value), "dd/mm/yyyy")
    PdfPut html, "DATE", value
    PdfPut html, "TEMPLATE", GdtXmlText(invoice, "TTChung/KHMSHDon")
    PdfPut html, "SERIES", GdtXmlText(invoice, "TTChung/KHHDon")
    PdfPut html, "NUMBER", GdtXmlText(invoice, "TTChung/SHDon")
    PdfPut html, "SELLER", GdtXmlText(invoice, "NDHDon/NBan/Ten")
    PdfPut html, "SELLER_TAX", GdtXmlText(invoice, "NDHDon/NBan/MST")
    PdfPut html, "SELLER_ADDRESS", GdtXmlText(invoice, "NDHDon/NBan/DChi")
    PdfPut html, "SELLER_BANK", GdtXmlText(invoice, "NDHDon/NBan/STKNHang") & " " & GdtXmlText(invoice, "NDHDon/NBan/TNHang")
    PdfPut html, "BUYER_PERSON", GdtXmlText(invoice, "NDHDon/NMua/HVTNMHang")
    PdfPut html, "BUYER", GdtXmlText(invoice, "NDHDon/NMua/Ten")
    PdfPut html, "BUYER_TAX", GdtXmlText(invoice, "NDHDon/NMua/MST")
    PdfPut html, "BUYER_ADDRESS", GdtXmlText(invoice, "NDHDon/NMua/DChi")
    PdfPut html, "PAYMENT", GdtXmlText(invoice, "TTChung/HTTToan")
    PdfPut html, "CURRENCY", GdtXmlText(invoice, "TTChung/DVTTe")
    Set nodes = invoice.SelectNodes(GdtXmlPath(invoice, "NDHDon/DSHHDVu/HHDVu"))
    If nodes.Length = 0 Then Err.Raise 5, , "Missing invoice items"
    For Each node In nodes
        rows = rows & "<tr>" & PdfCell(GdtXmlText(node, "STT")) & PdfCell(PdfItemKind(GdtXmlText(node, "TChat"))) & _
            PdfCell(GdtXmlText(node, "THHDVu")) & PdfCell(GdtXmlText(node, "DVTinh")) & PdfCell(PdfNumber(node, "SLuong"), True) & _
            PdfCell(PdfNumber(node, "DGia"), True) & PdfCell(PdfNumber(node, "STCKhau"), True) & _
            PdfCell(GdtXmlText(node, "TSuat")) & PdfCell(PdfNumber(node, "ThTien"), True) & "</tr>"
    Next node
    Set nodes = invoice.SelectNodes(GdtXmlPath(invoice, "NDHDon/TToan/THTTLTSuat/LTSuat"))
    For Each node In nodes
        taxes = taxes & "<tr>" & PdfCell(GdtXmlText(node, "TSuat")) & PdfCell(PdfNumber(node, "ThTien"), True) & _
            PdfCell(PdfNumber(node, "TThue"), True) & "</tr>"
    Next node
    html = Replace$(html, "{{ITEMS}}", rows)
    html = Replace$(html, "{{TAXES}}", taxes)
    PdfPut html, "SUBTOTAL", PdfNumber(invoice, "NDHDon/TToan/TgTCThue")
    PdfPut html, "TAX_TOTAL", PdfNumber(invoice, "NDHDon/TToan/TgTThue")
    PdfPut html, "DISCOUNT", PdfNumber(invoice, "NDHDon/TToan/TTCKTMai")
    PdfPut html, "FEES", PdfNumber(invoice, "NDHDon/TToan/TgTPhi")
    PdfPut html, "TOTAL", PdfNumber(invoice, "NDHDon/TToan/TgTTTBSo")
    PdfPut html, "WORDS", GdtXmlText(invoice, "NDHDon/TToan/TgTTTBChu")
    PdfPut html, "AUTHORITY_CODE", GdtXmlText(doc, "/HDon/MCCQT")
    PdfPut html, "BUYER_SIGNATURE", PdfSigningTime(doc, "NMua")
    PdfPut html, "SELLER_SIGNATURE", PdfSigningTime(doc, "NBan")
    PdfPut html, "AUTHORITY_SIGNATURE", PdfSigningTime(doc, "CQT")
    If fromApi Then
        PdfPut html, "SOURCE_NOTE", UniConvert("Barn theer hieejn tuwf duwx lieeju API; khoong cos XML goosc.")
        PdfPut html, "SIGNATURE_NOTE", UniConvert("Khoong cos duwx lieeju chuwx kys soos goosc ddeer xasc minh.")
        PdfPut html, "FOOTER_NOTE", UniConvert("PDF dduwowjc duwng tuwf chi tieest API vaf duwx lieeju toorng howjp cufng hosa ddown.")
    Else
        PdfPut html, "SOURCE_NOTE", UniConvert("Barn theer hieejn tuwf duwx lieeju XML")
        PdfPut html, "SIGNATURE_NOTE", UniConvert("Thoong tin thowfi ddieerm kys dduwowjc ddocj tuwf XML; chuwowng trinhf chuwa xasc minh chuwx kys soos.")
        PdfPut html, "FOOTER_NOTE", UniConvert("PDF nafy laf barn theer hieejn. Luwu XML goosc ddeer ddoosi chieesu.")
    End If
    PdfRenderDocument = html
End Function

' Build an in-memory view for the existing renderer, never a replacement XML.
Public Function GdtBuildInvoiceHtmlFromDetail(ByVal detailJson As String, Optional ByVal summaryJson As String = "", _
    Optional ByVal expectedIdentity As String = "") As String
    Dim detail As Object, summary As Object, doc As Object, invoice As Object, parent As Object, node As Object
    Dim items As Collection, groups As Collection, item As Object, field As Variant, mapping As Variant
    Dim identity As String, value As String, kind As String
    Set detail = JsonConverter.ParseJSON(detailJson)
    If TypeName(detail) <> "Dictionary" Then Err.Raise 5, , "Expected API detail object"
    If GdtJsonHasError(detail) Then Err.Raise 5, , "API detail error response"
    For Each field In Array("nbmst", "khmshdon", "khhdon", "shdon")
        If Len(SafeJsonText(detail, CStr(field))) = 0 Then Err.Raise 5, , "Missing API PDF identity: " & field
    Next field
    identity = GdtPdfIdentity(SafeJsonText(detail, "nbmst"), SafeJsonText(detail, "khmshdon"), _
        SafeJsonText(detail, "khhdon"), SafeJsonText(detail, "shdon"))
    If Len(expectedIdentity) > 0 And identity <> expectedIdentity Then Err.Raise 5, , "PDF API invoice identity mismatch"
    Set summary = CreateObject("Scripting.Dictionary")
    If Len(summaryJson) > 0 Then
        Set summary = JsonConverter.ParseJSON(summaryJson)
        If TypeName(summary) <> "Dictionary" Then Err.Raise 5, , "Expected PDF summary object"
        If GdtPdfIdentity(SafeJsonText(summary, "nbmst"), SafeJsonText(summary, "khmshdon"), _
            SafeJsonText(summary, "khhdon"), SafeJsonText(summary, "shdon")) <> identity Then Err.Raise 5, , "PDF summary invoice identity mismatch"
    End If
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.resolveExternals = False: doc.setProperty "ProhibitDTD", True
    doc.LoadXML "<HDon><DLHDon><TTChung/><NDHDon><NBan/><NMua/><DSHHDVu/><TToan><THTTLTSuat/></TToan></NDHDon></DLHDon></HDon>"
    Set invoice = doc.SelectSingleNode("/HDon/DLHDon")
    For Each mapping In Array("thdon|TTChung/THDon", "khmshdon|TTChung/KHMSHDon", "khhdon|TTChung/KHHDon", _
        "shdon|TTChung/SHDon", "tdlap|TTChung/NLap", "dvtte|TTChung/DVTTe", "htttoan|TTChung/HTTToan", _
        "nbten|NDHDon/NBan/Ten", "nbmst|NDHDon/NBan/MST", "nbdchi|NDHDon/NBan/DChi", _
        "nbstkhoan|NDHDon/NBan/STKNHang", "nbtnhang|NDHDon/NBan/TNHang", "nmten|NDHDon/NMua/Ten", _
        "nmmst|NDHDon/NMua/MST", "nmdchi|NDHDon/NMua/DChi", "nmtnmua|NDHDon/NMua/HVTNMHang", _
        "tgtttbchu|NDHDon/TToan/TgTTTBChu")
        field = Split(CStr(mapping), "|")
        value = PdfApiValue(detail, summary, CStr(field(0)))
        PdfViewPut invoice, CStr(field(1)), value
    Next mapping
    If Len(GdtXmlText(invoice, "TTChung/NLap")) = 0 Then Err.Raise 5, , "Missing API PDF invoice date"
    If Len(GdtXmlText(invoice, "NDHDon/NBan/Ten")) = 0 Then Err.Raise 5, , "Missing API PDF seller name"
    For Each mapping In Array("tgtcthue|TgTCThue", "tgtthue|TgTThue", "tgtttbso|TgTTTBSo", "ttcktmai|TTCKTMai", "tgtphi|TgTPhi")
        field = Split(CStr(mapping), "|")
        value = PdfApiAmount(detail, summary, CStr(field(0)))
        If Len(value) = 0 And (field(0) = "tgtcthue" Or field(0) = "tgtthue" Or field(0) = "tgtttbso") Then _
            Err.Raise 5, , "Missing API PDF total: " & field(0)
        If Len(value) > 0 Then PdfViewPut invoice, "NDHDon/TToan/" & field(1), PdfApiNumber(value, CStr(field(0)))
    Next mapping
    PdfViewPut doc.documentElement, "MCCQT", PdfApiValue(detail, summary, "mhdon")
    Set items = GdtJsonArray(detail, "hdhhdvu")
    If items.Count = 0 Then Err.Raise 5, , "Missing API PDF invoice items"
    Set parent = doc.SelectSingleNode("/HDon/DLHDon/NDHDon/DSHHDVu")
    For Each item In items
        If TypeName(item) <> "Dictionary" Then Err.Raise 5, , "Invalid API PDF invoice item"
        If Len(SafeJsonText(item, "ten")) = 0 Then Err.Raise 5, , "Missing API PDF item name"
        kind = SafeJsonText(item, "tchat")
        If kind <> "4" And Not GdtJsonHasValue(item, "thtien") Then Err.Raise 5, , "Missing API PDF item amount"
        Set node = doc.createElement("HHDVu"): parent.appendChild node
        For Each mapping In Array("stt|STT", "tchat|TChat", "ten|THHDVu", "dvtinh|DVTinh")
            field = Split(CStr(mapping), "|")
            PdfViewPut node, CStr(field(1)), SafeJsonText(item, CStr(field(0)))
        Next mapping
        For Each mapping In Array("sluong|SLuong", "dgia|DGia", "stckhau|STCKhau", "thtien|ThTien")
            field = Split(CStr(mapping), "|")
            If GdtJsonHasValue(item, CStr(field(0))) Then _
                PdfViewPut node, CStr(field(1)), PdfApiNumber(item(field(0)), CStr(field(0)))
        Next mapping
        PdfViewPut node, "TSuat", PdfApiTaxRate(item)
    Next item
    Set groups = GdtJsonArray(detail, "thttltsuat")
    If groups.Count = 0 Then Set groups = GdtJsonArray(summary, "thttltsuat")
    Set parent = doc.SelectSingleNode("/HDon/DLHDon/NDHDon/TToan/THTTLTSuat")
    For Each item In groups
        If TypeName(item) <> "Dictionary" Then Err.Raise 5, , "Invalid API PDF tax group"
        Set node = doc.createElement("LTSuat"): parent.appendChild node
        PdfViewPut node, "TSuat", PdfApiTaxRate(item)
        If GdtJsonHasValue(item, "thtien") Then PdfViewPut node, "ThTien", PdfApiNumber(item("thtien"), "thtien")
        If GdtJsonHasValue(item, "tthue") Then PdfViewPut node, "TThue", PdfApiNumber(item("tthue"), "tthue")
    Next item
    GdtBuildInvoiceHtmlFromDetail = PdfRenderDocument(doc, identity, True)
End Function

Private Function PdfApiValue(ByVal detail As Object, ByVal summary As Object, ByVal key As String) As String
    PdfApiValue = SafeJsonText(detail, key)
    If Len(PdfApiValue) = 0 Then PdfApiValue = SafeJsonText(summary, key)
End Function

Private Function PdfApiAmount(ByVal detail As Object, ByVal summary As Object, ByVal key As String) As String
    If GdtJsonHasValue(detail, key) Then
        PdfApiAmount = PdfApiNumber(detail(key), key)
    ElseIf GdtJsonHasValue(summary, key) Then
        PdfApiAmount = PdfApiNumber(summary(key), key)
    End If
End Function

Private Function PdfApiNumber(ByVal value As Variant, ByVal field As String) As String
    PdfApiNumber = Replace$(Format$(GdtJsonNumber(value, field), "0.################"), Application.International(xlDecimalSeparator), ".")
    If Right$(PdfApiNumber, 1) = "." Then PdfApiNumber = Left$(PdfApiNumber, Len(PdfApiNumber) - 1)
End Function

Private Function PdfApiTaxRate(ByVal item As Object) As String
    Dim value As String
    value = SafeJsonText(item, "ltsuat")
    If value = "KCT" Or value = "KKKNT" Then PdfApiTaxRate = value: Exit Function
    value = SafeJsonText(item, "tsuat")
    If value = "KCT" Or value = "KKKNT" Or Right$(value, 1) = "%" Then PdfApiTaxRate = value: Exit Function
    If Len(value) > 0 Then PdfApiTaxRate = PdfApiNumber(GdtJsonNumber(item("tsuat"), "tsuat") * 100, "tsuat") & "%"
End Function

Private Sub PdfViewPut(ByVal parent As Object, ByVal path As String, ByVal value As String)
    Dim doc As Object, node As Object, part As Variant
    Set doc = parent.ownerDocument
    For Each part In Split(path, "/")
        Set node = parent.SelectSingleNode(CStr(part))
        If node Is Nothing Then Set node = doc.createElement(CStr(part)): parent.appendChild node
        Set parent = node
    Next part
    parent.Text = value
End Sub

Public Function GdtPdfSummaryFromRow(ByVal ws As Worksheet, ByVal row As Long, ByVal expectedIdentity As String) As String
    Dim record As Object, keys As Variant, columns As Variant, index As Long, value As Variant, groups As New Collection, group As Object
    If row < 3 Then Err.Raise 5, , "Missing PDF summary row"
    Set record = CreateObject("Scripting.Dictionary")
    keys = Array("khmshdon", "khhdon", "shdon", "dvtte", "nbten", "nbmst", "nbdchi", "mhdon", "nmten", "nmmst", "nmdchi", _
        "tgtcthue", "tgtthue", "ttcktmai", "tgtttbso", "tgtttbchu")
    columns = Array(3, 4, 5, 7, 9, 10, 11, 13, 15, 16, 17, 42, 44, 47, 49, 50)
    For index = 0 To UBound(keys)
        value = ws.Cells(row, columns(index)).Value2
        If Not IsEmpty(value) Then record.Add keys(index), value
    Next index
    value = ws.Cells(row, 6).Value2
    If Not IsEmpty(value) And IsNumeric(value) Then record.Add "tdlap", Format$(CDate(value), "yyyy-mm-dd")
    If GdtPdfIdentity(SafeJsonText(record, "nbmst"), SafeJsonText(record, "khmshdon"), SafeJsonText(record, "khhdon"), _
        SafeJsonText(record, "shdon")) <> expectedIdentity Then Err.Raise 5, , "PDF summary row identity mismatch"
    For index = 0 To 5
        value = ws.Cells(row, 18 + 4 * index).Value2
        If Not IsEmpty(value) And Len(CStr(value)) > 0 Then
            Set group = CreateObject("Scripting.Dictionary"): group.Add "tsuat", value
            If Not IsEmpty(ws.Cells(row, 19 + 4 * index).Value2) Then group.Add "thtien", ws.Cells(row, 19 + 4 * index).Value2
            If Not IsEmpty(ws.Cells(row, 20 + 4 * index).Value2) Then group.Add "tthue", ws.Cells(row, 20 + 4 * index).Value2
            groups.Add group
        End If
    Next index
    record.Add "thttltsuat", groups
    GdtPdfSummaryFromRow = JsonConverter.ConvertToJson(record)
End Function

Private Function PdfItemKind(ByVal value As String) As String
    Select Case value
        Case "1": PdfItemKind = UniConvert("Hafng hosa, dijch vuj")
        Case "2": PdfItemKind = UniConvert("Khuyeesn maji")
        Case "3": PdfItemKind = UniConvert("Chieest khaasu")
        Case "4": PdfItemKind = UniConvert("Ghi chus")
        Case Else: PdfItemKind = value
    End Select
End Function

Private Function PdfSigningTime(ByVal doc As Object, ByVal party As String) As String
    PdfSigningTime = GdtXmlText(doc, "/HDon/DSCKS/" & party & "/Signature/Object/SignatureProperties/SignatureProperty/SigningTime")
End Function

Public Function GdtPdfBrowserPath() As String
    Dim fso As Object, root As Variant, relative As Variant, path As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each root In Array(Environ$("ProgramFiles(x86)"), Environ$("ProgramFiles"), Environ$("LOCALAPPDATA"))
        If Len(CStr(root)) > 0 Then
            For Each relative In Array("Microsoft\Edge\Application\msedge.exe", "Google\Chrome\Application\chrome.exe")
                path = fso.BuildPath(CStr(root), CStr(relative))
                If fso.FileExists(path) Then GdtPdfBrowserPath = path: Exit Function
            Next relative
        End If
    Next root
    Err.Raise 53, , "PDF requires Microsoft Edge or Google Chrome"
End Function

Public Sub GdtExportXmlFolderPdf()
    Dim folder As String, fso As Object, file As Object, files As New Collection, path As Variant
    Dim done As Long, failed As Long
    On Error GoTo Failed
    With Application.FileDialog(4)
        .Title = UniConvert("Chojn thuw mujc XML ddeer taoj PDF")
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    GdtPdfBrowserPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each file In fso.GetFolder(folder).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "xml" Then files.Add CStr(file.Path)
    Next file
    ResetGdtOperationControl
    For Each path In files
        If Not WaitForGdtControl() Then Exit For
        On Error GoTo FileFailed
        GdtExportInvoicePdf CStr(path), fso.BuildPath(folder, fso.GetBaseName(CStr(path)) & ".pdf")
        MarkGdtRetrySuccess "XML", "local", "", "", fso.GetFileName(CStr(path)), "", "PDF", 1
        done = done + 1
NextFile:
        On Error GoTo Failed
    Next path
    MsgBoxUni "PDF: " & done & UniConvert(" thafnh coong; ") & failed & UniConvert(" looxi."), vbInformation
    Exit Sub
FileFailed:
    failed = failed + 1
    UpsertGdtErrorReport "XML", "local", "", "", fso.GetFileName(CStr(path)), "", Empty, _
        "PDF", "", 0, Err.Description, 1, 0, "PDF build failed; XML retained"
    Resume NextFile
Failed:
    MsgBoxUni UniConvert("Khoong taoj dduwowjc PDF: ") & Err.Description, vbExclamation
End Sub

Private Function PdfQuote(ByVal value As String) As String
    If InStr(value, Chr$(34)) > 0 Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then Err.Raise 5, , "Invalid PDF path"
    PdfQuote = Chr$(34) & value & Chr$(34)
End Function

Public Function GdtIsPdfFile(ByVal path As String) As Boolean
    Dim stream As Object
    On Error GoTo Invalid
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "windows-1252": stream.Open
    stream.LoadFromFile path
    If stream.Size > 100 Then GdtIsPdfFile = (stream.ReadText(5) = "%PDF-")
Invalid:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
End Function

Public Sub GdtExportInvoicePdf(ByVal xmlPath As String, ByVal pdfPath As String, Optional ByVal expectedIdentity As String = "")
    GdtPrintInvoiceHtml GdtBuildInvoiceHtml(xmlPath, expectedIdentity), pdfPath
End Sub

Public Sub GdtExportDetailPdf(ByVal detailJson As String, ByVal pdfPath As String, Optional ByVal summaryJson As String = "", _
    Optional ByVal expectedIdentity As String = "")
    GdtPrintInvoiceHtml GdtBuildInvoiceHtmlFromDetail(detailJson, summaryJson, expectedIdentity), pdfPath
End Sub

Private Sub GdtPrintInvoiceHtml(ByVal html As String, ByVal pdfPath As String)
    Dim fso As Object, staging As String, htmlPath As String, stagedPdf As String, browser As String
    Dim stream As Object, process As Object, command As String, deadline As Double, url As String
    Dim failureNumber As Long, failureText As String
    On Error GoTo Failed
    browser = GdtPdfBrowserPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    pdfPath = fso.GetAbsolutePathName(pdfPath)
    If LCase$(fso.GetExtensionName(pdfPath)) <> "pdf" Then Err.Raise 5, , "Expected PDF output path"
    staging = fso.BuildPath(fso.GetParentFolderName(pdfPath), ".hddt-pdf-" & Replace$(CreateRequestID(), "-", ""))
    fso.CreateFolder staging
    htmlPath = fso.BuildPath(staging, "invoice.html"): stagedPdf = fso.BuildPath(staging, "invoice.pdf")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText html
    stream.SaveToFile htmlPath, 2: stream.Close: Set stream = Nothing
    ' Percent-encode local paths, including #, %, spaces and Vietnamese names.
    url = "file:///" & Replace$(htmlPath, "\", "/")
    url = Replace$(Replace$(Replace$(url, "%", "%25"), "#", "%23"), " ", "%20")
    command = PdfQuote(browser) & " --headless --disable-gpu --no-first-run --no-default-browser-check" & _
        " --disable-extensions --disable-background-networking --no-pdf-header-footer" & _
        " --user-data-dir=" & PdfQuote(fso.BuildPath(staging, "profile")) & _
        " --print-to-pdf=" & PdfQuote(stagedPdf) & " " & PdfQuote(url)
    Set process = CreateObject("WScript.Shell").Exec(command)
    deadline = GdtClockSeconds() + 60
    Do While process.Status = 0
        If GdtStopRequested Or GdtClockSeconds() > deadline Then
            process.Terminate
            Err.Raise 5, , "PDF export stopped or timed out"
        End If
        GdtPumpWait
    Loop
    If Not GdtIsPdfFile(stagedPdf) Then Err.Raise 5, , "Browser did not produce a valid PDF"
    ' The existing PDF is replaced only after a new file has been validated.
    If MoveFileExW(StrPtr(stagedPdf), StrPtr(pdfPath), 1 Or 8) = 0 Then Err.Raise 5, , "Could not replace PDF; Windows error " & Err.LastDllError
    On Error Resume Next
    fso.DeleteFolder staging, True
    On Error GoTo 0
    Exit Sub
Failed:
    failureNumber = Err.Number: failureText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Not process Is Nothing Then
        If process.Status = 0 Then process.Terminate
    End If
    If Len(staging) > 0 And Not fso Is Nothing Then
        If fso.FolderExists(staging) Then fso.DeleteFolder staging, True
    End If
    On Error GoTo 0
    Err.Raise failureNumber, "GdtExportInvoicePdf", failureText
End Sub
