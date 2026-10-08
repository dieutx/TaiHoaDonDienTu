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

Private Sub CollectGdtZipFiles(ByVal folder As Object, ByVal files As Collection, ByVal depth As Long)
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
    
    On Error Resume Next
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
    Dim fileName As Variant
    fileName = Dir(xmlFolder & "\")
    count = 0
    While fileName <> ""
        If LCase$(Right$(fileName, 4)) = ".xml" Then
            If Not xmlDoc.Load(xmlFolder & "\" & fileName) Then
                UpsertGdtErrorReport IIf(loaiHD, "mua", "ban"), "XML", "", "", CStr(fileName), "", Empty, _
                    "Parse XML", "", 0, "Invalid XML", 1, 0, "Failed"
                GoTo NextFile
            End If
            Set oHHDVu = xmlDoc.SelectNodes("/HDon/DLHDon/NDHDon/DSHHDVu/HHDVu")
            If oHHDVu.Length = 0 Then
                UpsertGdtErrorReport IIf(loaiHD, "mua", "ban"), "XML", "", "", CStr(fileName), "", Empty, _
                    "Parse XML", "", 0, "Missing invoice items", 1, 0, "Failed"
                GoTo NextFile
            End If
            count = count + 1
            ' No optional field may retain content from another file or old row.
            rngData.Resize(oHHDVu.Length, 35).ClearContents
            With rngData
                On Error Resume Next
                .Offset(, 0).Value = xmlDoc.SelectSingleNode("/HDon/DLHDon").getAttribute("Id")
                .Offset(, 1).Value = xmlDoc.SelectSingleNode("//TTChung/KHHDon").text     ' Lay ky hieu hoa don
                .Offset(, 2).Value = xmlDoc.SelectSingleNode("//TTChung/SHDon").text    ' lay so hoa don
                .Offset(, 3).Value = xmlDoc.SelectSingleNode("//TTChung/NLap").text    'ngay lap hoa don
                .Offset(, 4).Value = xmlDoc.SelectSingleNode("//TTChung/DVTTe").text    'don vi tien te
                .Offset(, 5).Value = xmlDoc.SelectSingleNode("//TTChung/TGia").text    'ty gia
                .Offset(, 6).Value = xmlDoc.SelectSingleNode("//NBan/Ten").text    'ten nguoi ban
                .Offset(, 7).Value = xmlDoc.SelectSingleNode("//NBan/MST").text    'ma so thue nguoi ban
                .Offset(, 8).Value = xmlDoc.SelectSingleNode("//NBan/DChi").text    'Dia chi nguoi ban
                .Offset(, 9).Value = xmlDoc.SelectSingleNode("/HDon/DSCKS/NBan/Signature/Object/SignatureProperties/SignatureProperty/SigningTime").text    'ngay ky so nguoi ban
                .Offset(, 10).Value = xmlDoc.SelectSingleNode("//HDon/MCCQT").text    'Lay ma co quan thue
                .Offset(, 11).Value = xmlDoc.SelectSingleNode("/HDon/DSCKS/CQT/Signature/Object/SignatureProperties/SignatureProperty/SigningTime").text    'Lay ngay cap ma co quan thue
                .Offset(, 12).Value = xmlDoc.SelectSingleNode("//NMua/Ten").text    'Lay Ten nguoi mua
                .Offset(, 13).Value = xmlDoc.SelectSingleNode("//NMua/MST").text    'Lay mst nguoi mua
                .Offset(, 14).Value = xmlDoc.SelectSingleNode("//NMua/DChi").text    'Lay Dia chi nguoi mua
                .Offset(, 27).Value = xmlDoc.SelectSingleNode("//NDHDon/TToan/TgTThue").text     'Tong thue tren HD
                .Offset(, 28).Value = xmlDoc.SelectSingleNode("//TTChung/MSTTCGP").text     'MSTTCGP
                
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
                            .Offset(, 29).Value = dicLink.item(msttcgp)
                    End Select
                Else
                    .Offset(, 29).Value = "Khong tim thay link tra cuu"
                End If
                
                '/Lay ma tra cuu ----------------------------------------
                If .Offset(, 30).Value <> "" Then GoTo Da_co_maTC
                Set list = xmlDoc.SelectSingleNode("//DLHDon/TTKhac")   'Truong hop 1
                If Not list Is Nothing Then
                For k = 0 To list.ChildNodes.Length - 1
                    ttruong = list.ChildNodes(k).getElementsByTagName("TTruong")(0).text
                    If dicTenCotTC.Exists(ttruong) Then
                        .Offset(, 30).Value = list.ChildNodes(k).getElementsByTagName("DLieu")(0).text
                        Exit For
                        'GoTo Da_co_maTC
                    Else
                        .Offset(, 30).Value = "Khong co ma tra cuu"
                    End If
                Next k
                End If
                Set list = xmlDoc.SelectSingleNode("//DLHDon/TTChung/TTKhac") 'Truong hop 2
                If Not list Is Nothing Then
                For k = 0 To list.ChildNodes.Length - 1
                    ttruong = list.ChildNodes(k).getElementsByTagName("TTruong")(0).text
                    If dicTenCotTC.Exists(ttruong) Then
                        .Offset(, 30).Value = list.ChildNodes(k).getElementsByTagName("DLieu")(0).text
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
            
            Set oHHDVu = xmlDoc.SelectNodes("/HDon/DLHDon/NDHDon/DSHHDVu/HHDVu")
            
            Dim c As Long, ndong As Long, Thtien As Double, TSuat As Double
            With rngHHDV
                ndong = 0
                For i = 0 To oHHDVu.Length - 1
                    Thtien = 0: TSuat = 0: tthue = 0
                    For j = 0 To oHHDVu(i).ChildNodes.Length - 1
                        Select Case oHHDVu(i).ChildNodes(j).tagName
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
                                Thtien = Val(oHHDVu(i).ChildNodes(j).text)
                            Case "TSuat"
                                c = 9
                                TSuat = Val(oHHDVu(i).ChildNodes(j).text) / 100
                            Case Else
                                c = -1
                        End Select
                        If c >= 0 Then
                            .Offset(i, c).Value = oHHDVu(i).ChildNodes(j).text
                        End If
                    Next j
                    
                    Dim hasTax As Boolean, hasTotal As Boolean
                    hasTax = False: hasTotal = False
                    Set list = oHHDVu(i).SelectSingleNode("TTKhac")
                    If Not list Is Nothing Then
                        For k = 0 To list.ChildNodes.Length - 1
                            ttruong = list.ChildNodes(k).getElementsByTagName("TTruong")(0).text
                            If isInArray(ttruong, arrTThue) Then
                                tthue = Val(list.ChildNodes(k).getElementsByTagName("DLieu")(0).text)
                                hasTax = True
                            ElseIf isInArray(ttruong, arrThTiencoVAT) Then
                                .Offset(i, 11).Value = list.ChildNodes(k).getElementsByTagName("DLieu")(0).text
                                hasTotal = True
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
                    .Offset(1).Resize(ndong - 1, 15).Value = .Resize(, 15).Value    '14: la so cot cuoi cung cua du lieu chung
                End With
            End If
            
            If ndong > 0 Then
                'Thay doi gia tri rngData sau khi ghi du lieu
                Set rngData = rngData.Offset(ndong + 0)
            End If
            
            '------------------------------------------------------------------------------------------------------------/
        End If
        
NextFile:
        fileName = Dir
    Wend
    
    If Not showMessages Then Exit Sub
    If count > 0 Then
        MsgBox "Xong"
    Else
        MsgBoxUni "Th" & ChrW(432) & " m" & ChrW(7909) & "c kh" & ChrW(244) & "ng c" & ChrW(243) & " ch" & ChrW(7913) & "a file XML.", vbExclamation, "Thông báo"
    End If
    
    Set xmlDoc = Nothing
    
    
    
End Sub

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
