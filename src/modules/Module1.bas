Attribute VB_Name = "Module1"


'=== Module: Download PDF from onclick="DownloadFile('path', 1)" ===
Option Explicit

'--- T?i HTML c?a trang ---
Private Function GetHtml(ByVal pageUrl As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", pageUrl, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    If http.Status = 200 Then
        GetHtml = http.responseText
    Else
        err.Raise vbObjectError + 1001, , "HTTP status: " & http.Status & " " & http.statusText
    End If
End Function

'--- Tìm path trong onclick c?a th? <a id="LinkDownPDF" ...> ---
Public Function ExtractPdfPath(ByVal html As String, Optional ByVal anchorId As String = "LinkDownPDF") As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True
    ' B?t nhóm d?i s? d?u tiên trong DownloadFile('...'
    re.Pattern = "<a[^>]*\bid\s*=\s*[""']" & anchorId & "['""][^>]*\bonclick\s*=\s*[""']\s*DownloadFile\s*\(\s*['""]([^'""]+)['""]"
    
    If re.test(html) Then
        Set m = re.Execute(html)(0)
        ExtractPdfPath = m.SubMatches(0) ' chính là 'C2/5T/...pdf'
    Else
        ExtractPdfPath = "..."
    End If
End Function

'--- Ghép path tuong d?i thành URL tuy?t d?i ---
Private Function MakeAbsoluteUrl(ByVal pageUrl As String, ByVal path As String) As String
    If LCase$(Left$(path, 4)) = "http" Then
        MakeAbsoluteUrl = path
        Exit Function
    End If
    
    Dim origin As String, baseDir As String
    origin = GetOrigin(pageUrl)
    baseDir = Left$(pageUrl, InStrRev(pageUrl, "/"))
    
    If Left$(path, 1) = "/" Then
        MakeAbsoluteUrl = origin & path
    Else
        ' path tuong d?i thu m?c hi?n t?i c?a trang
        MakeAbsoluteUrl = baseDir & path
    End If
End Function

'--- L?y origin (protocol + host + :port) t? URL ---
Private Function GetOrigin(ByVal url As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^(https?://[^/]+)"
    If re.test(url) Then
        Set m = re.Execute(url)(0)
        GetOrigin = m.SubMatches(0)
    Else
        GetOrigin = ""
    End If
End Function

'--- T?i file nh? phân và luu ---
Private Sub DownloadBinary(ByVal fileUrl As String, ByVal savePath As String)
    Dim http As Object, stm As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", fileUrl, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    If http.Status <> 200 Then
        err.Raise vbObjectError + 1002, , "Download failed: " & http.Status & " " & http.statusText
    End If
    
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile savePath, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

'=== Ví d? s? d?ng: t? l?y path trong onclick r?i t?i PDF ===
Public Sub SaveInvoicePdf()
    'On Error GoTo EH
    
    Dim pageUrl As String, saveFolder As String
    pageUrl = "https://van.ehoadon.vn/Lookup?InvoiceGUID=00000000-0000-0000-0000-000000000000" ' <-- ÐI?N URL trang ch?a th? <a id="LinkDownPDF">
    saveFolder = "C:\Temp"                                      ' <-- ÐI?N thu m?c luu
    
    Dim html As String
    html = Sheet6.Range("AI1")
    'html = GetHtml(pageUrl)
    'Debug.Print html
    Dim relPath As String
    relPath = ExtractPdfPath(html, "LinkDownPDF")
    Debug.Print relPath
    If relPath = "" Then err.Raise vbObjectError + 1003, , "Không tìm th?y path PDF trong onclick."
    
    Dim fileUrl As String
    'fileUrl = MakeAbsoluteUrl(pageUrl, relPath)
    
    fileUrl = "https://van.ehoadon.vn/DownloadFile?FilePath=" & relPath & "&BFType=1"
    
    Dim fileName As String
    fileName = Mid$(relPath, InStrRev(relPath, "/") + 1)
    If LenB(fileName) = 0 Then fileName = "download.pdf"
    
    Dim savePath As String
    savePath = saveFolder & IIf(Right$(saveFolder, 1) = "\", "", "\") & fileName
    
    DownloadBinary fileUrl, savePath
    MsgBox "Ðã t?i: " & savePath, vbInformation
    Exit Sub
EH:
    MsgBox "L?i: " & err.Description, vbExclamation
End Sub

'=== Tru?ng h?p b?n dã bi?t tru?c path 'C2/5T/...pdf' t? snippet ===
Public Sub SaveInvoicePdf_KnowPath()
    On Error GoTo EH
    
    Dim pageUrl As String, relPath As String, saveFolder As String
    pageUrl = "https://example.com/your/page/that/has_the_link"     ' d? ghép origin/baseDir
    relPath = "C2/5T/C25TBD-00000527-QJ34EHMPEB9-DPH.pdf"           ' path l?y t? onclick
    saveFolder = "C:\Temp"
    
    Dim fileUrl As String
    fileUrl = MakeAbsoluteUrl(pageUrl, relPath)
    
    Dim fileName As String
    fileName = Mid$(relPath, InStrRev(relPath, "/") + 1)
    If LenB(fileName) = 0 Then fileName = "download.pdf"
    
    Dim savePath As String
    savePath = saveFolder & IIf(Right$(saveFolder, 1) = "\", "", "\") & fileName
    
    DownloadBinary fileUrl, savePath
    MsgBox "Ðã t?i: " & savePath, vbInformation
    Exit Sub
EH:
    MsgBox "L?i: " & err.Description, vbExclamation
End Sub





Sub GetOnclickValue()
    Dim html As New MSHTML.HTMLDocument
    Dim http As Object, el As Object
    Dim pageUrl As String
    
    'pageUrl = "https://example.com/page/with/link" ' URL trang ch?a th?
    
    '--- t?i HTML ---
'    Set http = CreateObject("MSXML2.XMLHTTP")
'    http.Open "GET", pageUrl, False
'    http.send
    'html.body.innerHTML = http.responseText
    html.body.innerHTML = Sheet6.Range("AI1")
    
    
    
    '--- tìm ph?n t? theo ID ---
    Set el = html.getElementById("LinkDownPDF")
    If Not el Is Nothing Then
        Debug.Print el.getAttribute("onclick")
        ' ví d? in ra: DownloadFile('C2/5T/C25TBD-00000527-QJ34EHMPEB9-DPH.pdf',1);
    Else
        MsgBox "Không tìm th?y ph?n t?", vbExclamation
    End If
End Sub


Sub layLinkTaiPDF(urlTraCuu As String)
    Dim html As New MSHTML.HTMLDocument
    Dim http As Object, el As Object
    Dim urlTai As String
    
    'pageUrl = "https://example.com/page/with/link" ' URL trang ch?a th?
    
    '--- t?i HTML ---
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", urlTraCuu, False
    http.send
    html.body.innerHTML = http.responseText
    'html.body.innerHTML = Sheet6.Range("AI1")
    
    '--- tìm ph?n t? theo ID ---
    Set el = html.getElementById("LinkDownPDF")
    If Not el Is Nothing Then
        Debug.Print el.getAttribute("onclick")
        urlTai = el.getAttribute("onclick")
        ' ví d? in ra: DownloadFile('C2/5T/C25TBD-00000527-QJ34EHMPEB9-DPH.pdf',1);
    Else
        urlTai = ""
        MsgBox "Không tìm th?y ph?n t?", vbExclamation
    End If
    urlTai = Mid(urlTai, InStr(1, urlTai, "'") + 1, InStrRev(urlTai, "'") - InStr(1, urlTai, "'"))
    Debug.Print urlTai
End Sub


