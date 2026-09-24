Attribute VB_Name = "modUpdate"
Option Explicit

' Called once from ThisWorkbook.Workbook_Open. Failure never blocks the workbook.
Public Sub CheckForUpdate()
    On Error GoTo QuietFailure

    Dim request As Object, metadata As Object, notes As Object
    Dim newVersion As String, downloadUrl As String, releaseDate As String
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    request.setTimeouts 1500, 1500, 2500, 2500
    request.Open "GET", UPDATE_URL, False
    request.send
    If request.Status <> 200 Then Exit Sub

    Set metadata = JsonConverter.ParseJSON(CStr(request.responseText))
    If Not metadata.Exists("version") Then Exit Sub
    If Not metadata.Exists("downloadUrl") Then Exit Sub
    newVersion = CStr(metadata("version"))
    downloadUrl = CStr(metadata("downloadUrl"))
    If CompareVersions(newVersion, CURRENT_VERSION) <= 0 Then Exit Sub
    If Left$(downloadUrl, Len("https://github.com/dieutx/TaiHoaDonDienTu/releases/download/")) <> _
       "https://github.com/dieutx/TaiHoaDonDienTu/releases/download/" Then Exit Sub

    releaseDate = ""
    If metadata.Exists("releaseDate") Then releaseDate = CStr(metadata("releaseDate"))
    If metadata.Exists("releaseNote") Then Set notes = metadata("releaseNote")

    Dim notice As frmUpdate
    Set notice = New frmUpdate
    notice.SetUpdateInfo newVersion, releaseDate, downloadUrl, notes
    notice.Show vbModal
    Unload notice
    Exit Sub

QuietFailure:
    Debug.Print "Update check skipped: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Compare numeric version components, so 6.7.10 is newer than 6.7.9.
Public Function CompareVersions(ByVal leftVersion As String, ByVal rightVersion As String) As Long
    Dim leftParts() As String, rightParts() As String
    Dim i As Long, leftNumber As Long, rightNumber As Long, maxIndex As Long
    leftParts = Split(leftVersion, ".")
    rightParts = Split(rightVersion, ".")
    If UBound(leftParts) > 3 Or UBound(rightParts) > 3 Then Err.Raise 5
    maxIndex = UBound(leftParts)
    If UBound(rightParts) > maxIndex Then maxIndex = UBound(rightParts)
    For i = 0 To maxIndex
        leftNumber = 0: rightNumber = 0
        If i <= UBound(leftParts) Then leftNumber = VersionPart(leftParts(i))
        If i <= UBound(rightParts) Then rightNumber = VersionPart(rightParts(i))
        If leftNumber > rightNumber Then CompareVersions = 1: Exit Function
        If leftNumber < rightNumber Then CompareVersions = -1: Exit Function
    Next i
End Function

Private Function VersionPart(ByVal value As String) As Long
    Dim i As Long, digit As Long, result As Long
    If Len(value) = 0 Or Len(value) > 9 Then Err.Raise 5
    For i = 1 To Len(value)
        digit = Asc(Mid$(value, i, 1)) - 48
        If digit < 0 Or digit > 9 Then Err.Raise 5
        result = result * 10 + digit
    Next i
    VersionPart = result
End Function
