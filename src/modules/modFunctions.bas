Attribute VB_Name = "modFunctions"
Option Explicit




Sub Unzip(ZipfilePath As Variant, savedFolderPath As Variant)
    '* ZipFilePath:     toan bo duong dan + ten file zip
    '* savedFolderPath: duong dan toi folder luu file trich xuat (xml). Khong co dau "\" cuoi.
    
    Dim FSO As Object, oApp As Object
    Dim fileNameInZip As Variant, zipFileName As String, newFileName As String
    
    zipFileName = Replace(Mid(ZipfilePath, InStrRev(ZipfilePath, "\") + 1), ".zip", "") 'Lay ten file zip --> de rename ten file xml tuong ung
    Set oApp = CreateObject("Shell.Application")
    For Each fileNameInZip In oApp.Namespace(ZipfilePath).Items 'Duyet tung file trong file zip
        If LCase(fileNameInZip) Like LCase("*.xml") Then
            newFileName = zipFileName & ".xml"
            oApp.Namespace(savedFolderPath).CopyHere _
                oApp.Namespace(ZipfilePath).Items.item(CStr(fileNameInZip))
            Name savedFolderPath & "\" & fileNameInZip As savedFolderPath & "\" & newFileName
            'Exit For
        ElseIf LCase(fileNameInZip) Like LCase("*.html") Then
            newFileName = zipFileName & ".html"
            oApp.Namespace(savedFolderPath).CopyHere _
                oApp.Namespace(ZipfilePath).Items.item(CStr(fileNameInZip))
            Name savedFolderPath & "\" & fileNameInZip As savedFolderPath & "\" & newFileName
        End If
        
    Next
    
    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")
    FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    
    Set FSO = Nothing
    Set oApp = Nothing
End Sub

Sub parseXML(xmlFolder As Variant, loaiHD As Boolean)
    Dim xmlDoc As Object, oHHDVu As Object, rngData As Range, rngHHDV As Range, list As Object, ttruong As String
    Dim sht As Worksheet, i As Long, j As Long, k As Long, l As Long, count As Long
    
    On Error Resume Next
    
    If loaiHD = True Then
        Set sht = ThisWorkbook.Sheets("ChiTietHD_Mua_XML")
    Else
        Set sht = ThisWorkbook.Sheets("ChiTietHD_Ban_XML")
    End If
    
    '/Liet ke các ten field dai dien cho cot [tthueVAT] và [THTiencoVAT]
    Dim arrTThue As Variant, arrThTiencoVAT As Variant, tthue As Double
    arrTThue = Split(Sheets("LinkTraCuu").Range("O14"), ",")
    arrThTiencoVAT = Split(Sheets("LinkTraCuu").Range("O15"), ",")
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False: xmlDoc.validateOnParse = False
    
    Set rngData = sht.Range("A3")
    '// Xoa du lieu cu
    Dim r As Long, ans As Long
    With rngData
        r = .Offset(.Parent.Rows.count - .row - 10).End(xlUp).row - .row + 1
        If r > 0 Then
            ans = MsgBoxUni("B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n X" & ChrW(243) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(361) & " kh" & ChrW(244) & "ng?", vbYesNoCancel + vbDefaultButton2, "Thông báo")
            If ans = vbYes Then
                .Resize(r, 35).ClearContents
            ElseIf ans = vbNo Then
                Set rngData = rngData.Offset(r)
            Else    'cancel
                End
            End If
        End If
    End With
    
    '/Duyet tung file XML de trich xuat du lieu chi tiet
    Dim fileName As Variant
    fileName = Dir(xmlFolder & "\")
    count = 0
    While fileName <> ""
        If Right(fileName, 3) = "xml" Then
            count = count + 1
            xmlDoc.Load (xmlFolder & "\" & fileName)
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
                Set list = xmlDoc.SelectSingleNode("//DLHDon/TTChung/TTKhac") 'Truong hop 2
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
Da_co_maTC:
                '-----------------------------------------------------/
                
                Set rngHHDV = .Offset(, 15)    ' O bat dau lay so thu tu
            End With
            
            Set oHHDVu = xmlDoc.SelectNodes("/HDon/DLHDon/NDHDon/DSHHDVu/HHDVu")
            
            Dim c As Long, ndong As Long, Thtien As Double, TSuat As Double
            With rngHHDV
                ndong = 0
                For i = 0 To oHHDVu.Length - 1
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
                    
                    Set list = oHHDVu(i).SelectSingleNode("//HHDVu/TTKhac")
                    tthue = 0
                    For k = 0 To list.ChildNodes.Length - 1
                        ttruong = list.ChildNodes(k).getElementsByTagName("TTruong")(0).text
                        If isInArray(ttruong, arrTThue) Then
                            .Offset(i, 10).Value = list.ChildNodes(k).getElementsByTagName("DLieu")(0).text
                            tthue = Val(list.ChildNodes(k).getElementsByTagName("DLieu")(0).text)
                        ElseIf isInArray(ttruong, arrThTiencoVAT) Then
                            .Offset(i, 11).Value = list.ChildNodes(k).getElementsByTagName("DLieu")(0).text
                        End If
                        
                        'Truong hop khong co cot [tthue]
                        If tthue = 0 Then
                            .Offset(i, 10).Value = Thtien * TSuat / 100
                        End If
                        If .Offset(i, 11).Value = 0 Then  'Khong co gia tri cho cot THTiencoVAT
                            .Offset(i, 11).Value = Thtien + tthue
                        End If
                    Next k
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
        
        fileName = Dir
    Wend
    
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
