Attribute VB_Name = "modFunctions"
Option Explicit

Public Function GdtInvoiceFileBase(ByVal sellerTaxCode As String, ByVal templateCode As String, _
    ByVal invoiceSeries As String, ByVal invoiceNumber As String) As String
    GdtInvoiceFileBase = GdtSafeFilePart(sellerTaxCode) & "_" & GdtSafeFilePart(templateCode) & _
        "_" & GdtSafeFilePart(invoiceSeries) & "_" & GdtSafeFilePart(invoiceNumber)
End Function

Private Function GdtSafeFilePart(ByVal value As String) As String
    Dim i As Long, ch As String
    value = Trim$(value)
    If Len(value) = 0 Then Err.Raise 5, , "Missing invoice file identity"
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        ' Escape instead of replace: different identifiers must not collide.
        If ch Like "[A-Za-z0-9-]" Then
            GdtSafeFilePart = GdtSafeFilePart & ch
        Else
            GdtSafeFilePart = GdtSafeFilePart & "~" & Right$("0000" & Hex$(AscW(ch) And &HFFFF&), 4)
        End If
    Next i
End Function

Sub Unzip(ZipfilePath As Variant, savedFolderPath As Variant)
    Dim fso As Object, app As Object, archive As Object, destination As Object, entry As Object
    Dim staging As String, copied As String, target As String, baseName As String, extension As String
    Dim document As Object, seen As Object, files As Collection, index As Long, suffix As String
    Dim deadline As Double, stableAt As Double, errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CStr(savedFolderPath)) Then Err.Raise 76, , "Missing XML destination"
    staging = fso.BuildPath(CStr(savedFolderPath), ".hddt-xml-" & Replace(CreateRequestID(), "-", ""))
    fso.CreateFolder staging
    Set app = CreateObject("Shell.Application")
    Set archive = app.Namespace(CVar(fso.GetAbsolutePathName(CStr(ZipfilePath))))
    Set destination = app.Namespace(CVar(staging))
    If archive Is Nothing Or destination Is Nothing Then Err.Raise 5, , "Invalid ZIP archive"
    Set files = New Collection
    CollectGdtZipFiles archive, files, 0
    If files.count = 0 Then Err.Raise 5, , "Archive contains no XML/HTML"
    index = 0
    For Each entry In files
        If entry.Size > 52428800 Then Err.Raise 5, , "XML/HTML entry too large"
        copied = fso.BuildPath(staging, CStr(entry.Name))
        destination.CopyHere entry, 4 + 16 + 512 + 1024
        deadline = GdtClockSeconds() + 30
        stableAt = 0
        Do
            If fso.FileExists(copied) Then
                If CDbl(fso.GetFile(copied).Size) = CDbl(entry.Size) Then
                    If stableAt = 0 Then stableAt = GdtClockSeconds()
                    If GdtClockSeconds() - stableAt >= 0.25 Then Exit Do
                Else
                    stableAt = 0
                End If
            End If
            If GdtClockSeconds() > deadline Then Err.Raise 5, , "ZIP extraction timed out"
            GdtPumpWait
        Loop
        extension = LCase$(fso.GetExtensionName(copied))
        If extension = "xml" Then
            Set document = CreateObject("MSXML2.DOMDocument.6.0")
            document.async = False
            document.resolveExternals = False
            document.setProperty "ProhibitDTD", True
            If Not document.Load(copied) Then Err.Raise 5, , "Invalid XML in ZIP archive"
            Set document = Nothing
        End If
        index = index + 1
        fso.MoveFile copied, fso.BuildPath(staging, CStr(index) & ".ready." & extension)
    Next entry
    ' Validate all entries before replacing destination files.
    baseName = fso.GetBaseName(CStr(ZipfilePath))
    Set seen = CreateObject("Scripting.Dictionary")
    For index = 1 To files.count
        extension = LCase$(fso.GetExtensionName(CStr(files(index).Name)))
        copied = fso.BuildPath(staging, CStr(index) & ".ready." & extension)
        suffix = vbNullString
        If seen.Exists(extension) Then
            seen(extension) = seen(extension) + 1
            suffix = "_" & CStr(seen(extension))
        Else
            seen.Add extension, 1
        End If
        target = fso.BuildPath(CStr(savedFolderPath), baseName & suffix & "." & extension)
        If fso.FileExists(target) Then fso.DeleteFile target, True
        fso.MoveFile copied, target
    Next index
    Set destination = Nothing
    fso.DeleteFolder staging, True
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Set destination = Nothing
    On Error Resume Next
    If Len(staging) > 0 Then
        If Not fso Is Nothing Then
            If fso.FolderExists(staging) Then fso.DeleteFolder staging, True
        End If
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "Unzip", errorText
End Sub

Public Sub CollectGdtZipFiles(ByVal folder As Object, ByVal files As Collection, ByVal depth As Long)
    Dim entry As Object, extension As String
    If depth > 8 Then Err.Raise 5, , "ZIP nesting too deep"
    For Each entry In folder.Items
        If entry.Name = "." Or entry.Name = ".." Or InStr(entry.Name, "\") > 0 Or InStr(entry.Name, "/") > 0 Then _
            Err.Raise 5, , "Unsafe ZIP entry"
        If entry.IsFolder Then
            CollectGdtZipFiles entry.GetFolder, files, depth + 1
        Else
            extension = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(CStr(entry.Name)))
            If extension = "xml" Or extension = "html" Then files.Add entry
        End If
    Next entry
End Sub

Sub parseXML(xmlFolder As Variant, loaiHD As Boolean, Optional ByVal showMessages As Boolean = True)
    Dim xmlDoc As Object, oHHDVu As Object, rngData As Range, rngHHDV As Range, list As Object, ttruong As String
    Dim sht As Worksheet, i As Long, j As Long, k As Long, l As Long, count As Long
    
    On Error GoTo SetupFailed
    LinkTraCuu
    tenCotTraCuu
    
    If loaiHD = True Then
        Set sht = ThisWorkbook.Sheets("ChiTietHD_Mua_XML")
    Else
        Set sht = ThisWorkbook.Sheets("ChiTietHD_Ban_XML")
    End If
    
    '/Liet ke các ten field dai dien cho cot [tthueVAT] và [THTiencoVAT]
    Dim arrTThue As Variant, arrThTiencoVAT As Variant, tthue As Double
    arrTThue = Split(Sheets("LinkTraCuu").Range("O14"), ",")
    arrThTiencoVAT = Split(Sheets("LinkTraCuu").Range("O15"), ",")
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False: xmlDoc.validateOnParse = False
    xmlDoc.resolveExternals = False
    xmlDoc.setProperty "ProhibitDTD", True
    
    Set rngData = sht.Range("A3")
    '// Xoa du lieu cu
    Dim r As Long, ans As Long
    With rngData
        r = .Offset(.Parent.Rows.count - .row - 10).End(xlUp).row - .row + 1
        If r > 0 Then
            If showMessages Then
            ans = MsgBoxUni("B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n X" & ChrW(243) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(361) & " kh" & ChrW(244) & "ng?", vbYesNoCancel + vbDefaultButton2, "Thông báo")
            Else
                ans = vbNo
            End If
            If ans = vbYes Then
                .Resize(r, 35).ClearContents
            ElseIf ans = vbNo Then
                Set rngData = rngData.Offset(r)
            Else    'cancel
                Exit Sub
            End If
        End If
    End With
    
    '/Duyet tung file XML de trich xuat du lieu chi tiet
    Dim fileName As Variant, failureText As String, writingRows As Long
    fileName = Dir(xmlFolder & "\")
    count = 0
    While fileName <> ""
        If LCase$(Right$(fileName, 4)) = ".xml" Then
            On Error GoTo FileFailed
            writingRows = 0
            If Not xmlDoc.Load(xmlFolder & "\" & fileName) Then
                UpsertGdtErrorReport IIf(loaiHD, "mua", "ban"), "XML", "", "", CStr(fileName), "", Empty, _
                    "Parse XML", "", 0, "Invalid XML", 1, 0, "Failed"
                GoTo NextFile
            End If
            Set oHHDVu = xmlDoc.SelectNodes(GdtXmlPath(xmlDoc, "/HDon/DLHDon/NDHDon/DSHHDVu/HHDVu"))
            If oHHDVu.Length = 0 Then
                UpsertGdtErrorReport IIf(loaiHD, "mua", "ban"), "XML", "", "", CStr(fileName), "", Empty, _
                    "Parse XML", "", 0, "Missing invoice items", 1, 0, "Failed"
                GoTo NextFile
            End If
            ' No optional field may retain content from another file or old row.
            writingRows = oHHDVu.Length
            rngData.Resize(oHHDVu.Length, 35).ClearContents
            With rngData
                GdtWriteTextCell .Offset(, 0), xmlDoc.SelectSingleNode(GdtXmlPath(xmlDoc, "/HDon/DLHDon")).getAttribute("Id")
                GdtWriteTextCell .Offset(, 1), GdtXmlText(xmlDoc, "/HDon/DLHDon/TTChung/KHHDon")
                GdtWriteTextCell .Offset(, 2), GdtXmlText(xmlDoc, "/HDon/DLHDon/TTChung/SHDon")
                GdtWriteXmlDate .Offset(, 3), GdtXmlText(xmlDoc, "/HDon/DLHDon/TTChung/NLap")
                GdtWriteTextCell .Offset(, 4), GdtXmlText(xmlDoc, "/HDon/DLHDon/TTChung/DVTTe")
                GdtWriteXmlNumber .Offset(, 5), GdtXmlText(xmlDoc, "/HDon/DLHDon/TTChung/TGia"), "TGia"
                GdtWriteTextCell .Offset(, 6), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/NBan/Ten")
                GdtWriteTextCell .Offset(, 7), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/NBan/MST")
                GdtWriteTextCell .Offset(, 8), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/NBan/DChi")
                GdtWriteXmlDate .Offset(, 9), GdtXmlText(xmlDoc, "/HDon/DSCKS/NBan/Signature/Object/SignatureProperties/SignatureProperty/SigningTime")
                GdtWriteTextCell .Offset(, 10), GdtXmlText(xmlDoc, "/HDon/MCCQT")
                GdtWriteXmlDate .Offset(, 11), GdtXmlText(xmlDoc, "/HDon/DSCKS/CQT/Signature/Object/SignatureProperties/SignatureProperty/SigningTime")
                GdtWriteTextCell .Offset(, 12), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/NMua/Ten")
                GdtWriteTextCell .Offset(, 13), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/NMua/MST")
                GdtWriteTextCell .Offset(, 14), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/NMua/DChi")
                GdtWriteXmlNumber .Offset(, 27), GdtXmlText(xmlDoc, "/HDon/DLHDon/NDHDon/TToan/TgTThue"), "TgTThue"
                GdtWriteTextCell .Offset(, 28), GdtXmlText(xmlDoc, "/HDon/DLHDon/TTChung/MSTTCGP")
                
                '/Lay link tra cuu
                Dim msttcgp As String
                If .Offset(, 28).Value <> "" Then   'Co MSTTCGP
                    Dim mst As String, mccqt As String
                    msttcgp = .Offset(, 28).Value
                    mst = .Offset(, 7).Value: mccqt = .Offset(, 10).Value
                    Select Case msttcgp
                        Case "0100684378"   'VNPT
                            If mccqt <> "" Then
                                .Offset(, 29).Value = "https://" & mst & "-tt78.vnpt-invoice.com.vn/?strFkey=" & mccqt
                                .Offset(, 30).Value = mccqt
                            Else
                                .Offset(, 29).Value = "Không có MCCQT"
                            End If
                        Case "0101360697"   'BKAV
                            .Offset(, 29).Value = "https://van.ehoadon.vn/Lookup?InvoiceGUID=" & .Offset(, 0).Value
                            .Offset(, 30).Value = .Offset(, 0).Value
                        Case "0105987432"
                            .Offset(, 29).Value = "https://" & mst & "hd.easyinvoice.com.vn"
                            .Offset(, 30).Value = mst
                        Case Else
                            If dicLink.Exists(msttcgp) Then .Offset(, 29).Value = dicLink.item(msttcgp)
                    End Select
                Else
                    .Offset(, 29).Value = "Khong tim thay link tra cuu"
                End If
                
                '/Lay ma tra cuu ----------------------------------------
                If .Offset(, 30).Value <> "" Then GoTo Da_co_maTC
                Set list = xmlDoc.SelectSingleNode(GdtXmlPath(xmlDoc, "/HDon/DLHDon/TTKhac"))
                If Not list Is Nothing Then
                For k = 0 To list.ChildNodes.Length - 1
                    ttruong = GdtXmlText(list.ChildNodes(k), "TTruong")
                    If dicTenCotTC.Exists(ttruong) Then
                        GdtWriteTextCell .Offset(, 30), GdtXmlText(list.ChildNodes(k), "DLieu")
                        Exit For
                        'GoTo Da_co_maTC
                    Else
                        .Offset(, 30).Value = "Khong co ma tra cuu"
                    End If
                Next k
                End If
                Set list = xmlDoc.SelectSingleNode(GdtXmlPath(xmlDoc, "/HDon/DLHDon/TTChung/TTKhac"))
                If Not list Is Nothing Then
                For k = 0 To list.ChildNodes.Length - 1
                    ttruong = GdtXmlText(list.ChildNodes(k), "TTruong")
                    If dicTenCotTC.Exists(ttruong) Then
                        GdtWriteTextCell .Offset(, 30), GdtXmlText(list.ChildNodes(k), "DLieu")
                        Exit For
                        'GoTo Da_co_maTC
                    Else
                        .Offset(, 30).Value = "Khong co ma tra cuu"
                    End If
                Next k
                End If
Da_co_maTC:
                '-----------------------------------------------------/
                
                Set rngHHDV = .Offset(, 15)    ' O bat dau lay so thu tu
            End With
            
            Set oHHDVu = xmlDoc.SelectNodes(GdtXmlPath(xmlDoc, "/HDon/DLHDon/NDHDon/DSHHDVu/HHDVu"))
            
            Dim c As Long, ndong As Long, Thtien As Double, TSuat As Double
            With rngHHDV
                ndong = 0
                For i = 0 To oHHDVu.Length - 1
                    Thtien = 0: TSuat = 0: tthue = 0
                    For j = 0 To oHHDVu(i).ChildNodes.Length - 1
                        If CStr(oHHDVu(i).ChildNodes(j).namespaceURI) <> CStr(xmlDoc.documentElement.namespaceURI) Then GoTo NextItemField
                        Select Case oHHDVu(i).ChildNodes(j).baseName
                            Case "STT"
                                c = 0
                            Case "MHHDVu"
                                c = 1
                            Case "THHDVu", "Ten"    'Hoa don chua co ma co quan thue
                                c = 2
                            Case "DVTinh"
                                c = 3
                            Case "SLuong"
                                c = 4
                            Case "DGia"
                                c = 5
                            Case "TLCKhau"
                                c = 6
                            Case "STCKhau"
                                c = 7
                            Case "ThTien"
                                c = 8
                                Thtien = GdtXmlNumber(oHHDVu(i).ChildNodes(j).text, "ThTien")
                            Case "TSuat"
                                c = 9
                                TSuat = GdtXmlTaxRate(oHHDVu(i).ChildNodes(j).text)
                            Case Else
                                c = -1
                        End Select
                        If c >= 0 Then
                            If c >= 4 And c <= 8 Then
                                GdtWriteXmlNumber .Offset(i, c), oHHDVu(i).ChildNodes(j).text, oHHDVu(i).ChildNodes(j).baseName
                            Else
                                GdtWriteTextCell .Offset(i, c), oHHDVu(i).ChildNodes(j).text
                            End If
                        End If
NextItemField:
                    Next j
                    
                    Dim hasTax As Boolean, hasTotal As Boolean
                    hasTax = False: hasTotal = False
                    Set list = oHHDVu(i).SelectSingleNode(GdtXmlPath(oHHDVu(i), "TTKhac"))
                    If Not list Is Nothing Then
                        For k = 0 To list.ChildNodes.Length - 1
                            ttruong = GdtXmlText(list.ChildNodes(k), "TTruong")
                            If isInArray(ttruong, arrTThue) Then
                                If Len(Trim$(GdtXmlText(list.ChildNodes(k), "DLieu"))) > 0 Then
                                    tthue = GdtXmlNumber(GdtXmlText(list.ChildNodes(k), "DLieu"), "tax metadata")
                                    hasTax = True
                                End If
                            ElseIf isInArray(ttruong, arrThTiencoVAT) Then
                                If Len(Trim$(GdtXmlText(list.ChildNodes(k), "DLieu"))) > 0 Then
                                    .Offset(i, 11).Value2 = GdtXmlNumber(GdtXmlText(list.ChildNodes(k), "DLieu"), "total metadata")
                                    hasTotal = True
                                End If
                            End If
                        Next k
                    End If
                    If Not hasTax Then tthue = Thtien * TSuat
                    .Offset(i, 10).Value = tthue
                    If Not hasTotal Then .Offset(i, 11).Value = Thtien + tthue
                    ndong = ndong + 1
                Next i
            End With
            
            If ndong > 1 Then
                'Copy nhung dong du lieu chung xuong theo MHHDvu
                With rngData
                    Dim textColumn As Variant
                    For Each textColumn In Array(0, 1, 2, 4, 6, 7, 8, 10, 12, 13, 14)
                        .Resize(ndong, 1).Offset(, CLng(textColumn)).NumberFormat = "@"
                    Next textColumn
                    .Resize(ndong, 1).Offset(, 7).NumberFormat = "@"
                    .Resize(ndong, 1).Offset(, 13).NumberFormat = "@"
                    .Resize(ndong, 1).Offset(, 28).NumberFormat = "@"
                    .Resize(ndong, 1).Offset(, 3).NumberFormat = "dd/mm/yyyy"
                    .Resize(ndong, 1).Offset(, 9).NumberFormat = "dd/mm/yyyy"
                    .Resize(ndong, 1).Offset(, 11).NumberFormat = "dd/mm/yyyy"
                    .Resize(ndong, 15).FillDown
                End With
            End If
            MarkGdtRetrySuccess IIf(loaiHD, "mua", "ban"), "XML", "", "", CStr(fileName), "", "Parse XML", 1
            count = count + 1
            
            If ndong > 0 Then
                'Thay doi gia tri rngData sau khi ghi du lieu
                Set rngData = rngData.Offset(ndong + 0)
            End If
            
            '------------------------------------------------------------------------------------------------------------/
        End If
        
NextFile:
        On Error GoTo SetupFailed
        fileName = Dir
    Wend
    
    If Not showMessages Then Exit Sub
    If count > 0 Then
        MsgBox "Xong"
    Else
        MsgBoxUni "Th" & ChrW(432) & " m" & ChrW(7909) & "c kh" & ChrW(244) & "ng c" & ChrW(243) & " ch" & ChrW(7913) & "a file XML.", vbExclamation, "Thông báo"
    End If
    
    Set xmlDoc = Nothing
    Exit Sub

FileFailed:
    failureText = Err.Description
    If writingRows > 0 Then rngData.Resize(writingRows, 35).ClearContents
    UpsertGdtErrorReport IIf(loaiHD, "mua", "ban"), "XML", "", "", CStr(fileName), "", Empty, _
        "Parse XML", "", 0, failureText, 1, 0, "Failed"
    Resume NextFile

SetupFailed:
    Err.Raise Err.Number, "parseXML", Err.Description
End Sub

' Keep every path segment and restrict invoice nodes to the root namespace.
' Signature descendants may use the XMLDSIG namespace or the legacy namespace.
Public Function GdtXmlPath(ByVal context As Object, ByVal path As String) As String
    Dim doc As Object, parts As Variant, part As String, ns As String, predicate As String, index As Long
    If context.nodeType = 9 Then
        Set doc = context
    Else
        Set doc = context.ownerDocument
    End If
    ns = CStr(doc.documentElement.namespaceURI)
    If InStr(ns, "'") > 0 Then Err.Raise 5, , "Unsupported XML namespace"
    parts = Split(path, "/")
    For index = 0 To UBound(parts)
        part = CStr(parts(index))
        If Len(part) > 0 Then
            predicate = "namespace-uri()='" & ns & "'"
            Select Case CStr(part)
                Case "Signature", "Object", "SignatureProperties", "SignatureProperty", "SigningTime"
                    predicate = "(" & predicate & " or namespace-uri()='http://www.w3.org/2000/09/xmldsig#' or namespace-uri()='')"
            End Select
            parts(index) = "*[local-name()='" & part & "' and " & predicate & "]"
        End If
    Next index
    GdtXmlPath = Join(parts, "/")
End Function

Public Function GdtXmlText(ByVal context As Object, ByVal path As String) As String
    Dim node As Object
    Set node = context.SelectSingleNode(GdtXmlPath(context, path))
    If Not node Is Nothing Then GdtXmlText = node.text
End Function

Private Sub GdtWriteXmlDate(ByVal cell As Range, ByVal value As String)
    If Len(Trim$(value)) = 0 Then Exit Sub
    cell.Value2 = CDbl(ISODateValue(value))
    cell.NumberFormat = "dd/mm/yyyy"
End Sub

' XML numeric values use a dot, independent of the Windows decimal separator.
' Reject partial conversions (Val("invalid")=0, Val("10,5")=10).
Public Function GdtXmlNumber(ByVal value As String, ByVal field As String) As Double
    Static re As Object
    value = Trim$(value)
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)$"
    End If
    If Not re.test(value) Then Err.Raise 5, "GdtXmlNumber", "Invalid XML number: " & field
    GdtXmlNumber = Val(value)
End Function

Private Sub GdtWriteXmlNumber(ByVal cell As Range, ByVal value As String, ByVal field As String)
    If Len(Trim$(value)) > 0 Then cell.Value2 = GdtXmlNumber(value, field)
End Sub

Private Function GdtXmlTaxRate(ByVal value As String) As Double
    value = Trim$(value)
    Select Case UCase$(value)
        Case "", "KCT", "KKKNT", "KHAC"
            Exit Function
    End Select
    If Right$(value, 1) = "%" Then value = Trim$(Left$(value, Len(value) - 1))
    GdtXmlTaxRate = GdtXmlNumber(value, "TSuat") / 100
End Function

Public Function isInArray(ByRef FindValue As Variant, ByRef vArr As Variant) As Boolean
    Dim vArrEach As Variant
    For Each vArrEach In vArr
        isInArray = (FindValue = vArrEach)
        If isInArray Then Exit For
    Next
End Function

Public Function getToken() As String
    getToken = cToken
End Function
