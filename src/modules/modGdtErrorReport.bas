Attribute VB_Name = "modGdtErrorReport"
Option Explicit

Private Const GDT_ERROR_SHEET As String = "BaoCao_LoiTaiHD"
Private Const GDT_ERROR_COLUMN_COUNT As Long = 17

Public Function EnsureGdtErrorReportSheet() As Worksheet
    Dim ws As Worksheet

    On Error GoTo CreateSheet
    Set ws = ThisWorkbook.Worksheets(GDT_ERROR_SHEET)
    On Error GoTo 0
    If CStr(ws.Cells(1, 1).Value2) <> "STT" Then FormatGdtErrorReport ws
    Set EnsureGdtErrorReportSheet = ws
    Exit Function

CreateSheet:
    Err.Clear
    On Error GoTo CreateFailed
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = GDT_ERROR_SHEET
    On Error GoTo 0
    FormatGdtErrorReport ws
    Set EnsureGdtErrorReportSheet = ws
    Exit Function

CreateFailed:
    Err.Raise Err.Number, "EnsureGdtErrorReportSheet", Err.Description
End Function

Public Sub UpsertGdtErrorReport( _
    ByVal direction As String, _
    ByVal apiSource As String, _
    ByVal sellerTaxCode As String, _
    ByVal templateCode As String, _
    ByVal invoiceSeries As String, _
    ByVal invoiceNumber As String, _
    ByVal invoiceDate As Variant, _
    ByVal failureStage As String, _
    ByVal endpoint As String, _
    ByVal statusCode As Long, _
    ByVal errorText As String, _
    ByVal attempts As Long, _
    ByVal retryAfterSeconds As Long, _
    ByVal finalResult As String, _
    Optional ByVal note As String = vbNullString)

    Dim ws As Worksheet
    Dim targetRow As Long
    Dim rowData(1 To 1, 1 To GDT_ERROR_COLUMN_COUNT) As Variant

    Set ws = EnsureGdtErrorReportSheet()
    targetRow = FindGdtErrorRow(ws, direction, apiSource, sellerTaxCode, _
                                templateCode, invoiceSeries, invoiceNumber, failureStage)
    If targetRow = 0 Then targetRow = LastGdtErrorRow(ws) + 1

    rowData(1, 1) = targetRow - 1
    rowData(1, 2) = Now
    rowData(1, 3) = direction
    rowData(1, 4) = apiSource
    rowData(1, 5) = sellerTaxCode
    rowData(1, 6) = templateCode
    rowData(1, 7) = invoiceSeries
    rowData(1, 8) = invoiceNumber
    rowData(1, 9) = invoiceDate
    rowData(1, 10) = failureStage
    rowData(1, 11) = endpoint
    If statusCode > 0 Then rowData(1, 12) = statusCode
    rowData(1, 13) = RedactGdtSensitiveText(errorText)
    rowData(1, 14) = attempts
    If retryAfterSeconds > 0 Then rowData(1, 15) = retryAfterSeconds
    rowData(1, 16) = finalResult
    rowData(1, 17) = RedactGdtSensitiveText(note)
    ws.Cells(targetRow, 1).Resize(1, GDT_ERROR_COLUMN_COUNT).Value = rowData
End Sub

Public Sub MarkGdtRetrySuccess( _
    ByVal direction As String, _
    ByVal apiSource As String, _
    ByVal sellerTaxCode As String, _
    ByVal templateCode As String, _
    ByVal invoiceSeries As String, _
    ByVal invoiceNumber As String, _
    ByVal failureStage As String, _
    ByVal attempts As Long, _
    Optional ByVal note As String = vbNullString)

    Dim ws As Worksheet
    Dim targetRow As Long

    Set ws = EnsureGdtErrorReportSheet()
    Do
        targetRow = FindGdtErrorRow(ws, direction, apiSource, sellerTaxCode, _
                                    templateCode, invoiceSeries, invoiceNumber, failureStage)
        If targetRow = 0 Then Exit Do
        ws.Rows(targetRow).Delete Shift:=xlUp
    Loop
    RenumberGdtErrorRows ws
End Sub

Public Function BuildGdtErrorKey( _
    ByVal direction As String, _
    ByVal apiSource As String, _
    ByVal sellerTaxCode As String, _
    ByVal templateCode As String, _
    ByVal invoiceSeries As String, _
    ByVal invoiceNumber As String, _
    ByVal failureStage As String) As String

    BuildGdtErrorKey = NormalizeKeyPart(direction) & "|" & _
                       NormalizeKeyPart(apiSource) & "|" & _
                       NormalizeKeyPart(sellerTaxCode) & "|" & _
                       NormalizeKeyPart(templateCode) & "|" & _
                       NormalizeKeyPart(invoiceSeries) & "|" & _
                       NormalizeKeyPart(invoiceNumber) & "|" & _
                       NormalizeKeyPart(failureStage)
End Function

Private Sub FormatGdtErrorReport(ByVal ws As Worksheet)
    Dim headers(1 To 1, 1 To GDT_ERROR_COLUMN_COUNT) As Variant

    headers(1, 1) = "STT"
    headers(1, 2) = UniConvert("Thowfi gian ghi nhaajn")
    headers(1, 3) = UniConvert("Loaji hosa ddown")
    headers(1, 4) = UniConvert("Nguoofn API")
    headers(1, 5) = UniConvert("Max soos thuees nguwowfi basn")
    headers(1, 6) = UniConvert("Kys hieeju maaxu soos")
    headers(1, 7) = UniConvert("Kys hieeju hosa ddown")
    headers(1, 8) = UniConvert("Soos hosa ddown")
    headers(1, 9) = UniConvert("Ngafy laajp hosa ddown")
    headers(1, 10) = UniConvert("Coong ddoajn looxi")
    headers(1, 11) = "Endpoint"
    headers(1, 12) = "HTTP status"
    headers(1, 13) = UniConvert("Nooji dung looxi")
    headers(1, 14) = UniConvert("Soos laafn ddax thuwr")
    headers(1, 15) = "Retry-After"
    headers(1, 16) = UniConvert("Keest quar cuoosi")
    headers(1, 17) = UniConvert("Ghi chus")

    With ws
        .Cells(1, 1).Resize(1, GDT_ERROR_COLUMN_COUNT).Value = headers
        With .Cells(1, 1).Resize(1, GDT_ERROR_COLUMN_COUNT)
            .Font.Bold = True
            .Interior.Color = RGB(217, 225, 242)
            .WrapText = True
        End With
        If .AutoFilterMode Then .AutoFilterMode = False
        .Cells(1, 1).Resize(1, GDT_ERROR_COLUMN_COUNT).AutoFilter
        .Columns(1).ColumnWidth = 7
        .Columns(2).ColumnWidth = 19
        .Columns(3).ColumnWidth = 14
        .Columns(4).ColumnWidth = 12
        .Columns(5).ColumnWidth = 20
        .Columns(6).ColumnWidth = 18
        .Columns(7).ColumnWidth = 18
        .Columns(8).ColumnWidth = 16
        .Columns(9).ColumnWidth = 16
        .Columns(10).ColumnWidth = 18
        .Columns(11).ColumnWidth = 45
        .Columns(12).ColumnWidth = 12
        .Columns(13).ColumnWidth = 45
        .Columns(14).ColumnWidth = 14
        .Columns(15).ColumnWidth = 13
        .Columns(16).ColumnWidth = 24
        .Columns(17).ColumnWidth = 35
        .Columns(2).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Columns(9).NumberFormat = "dd/mm/yyyy"
    End With

    ' Freeze panes keeps the headings visible while scrolling the report.  It is
    ' best effort because it needs a window, which a headless build may not have.
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 0
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Function FindGdtErrorRow( _
    ByVal ws As Worksheet, _
    ByVal direction As String, _
    ByVal apiSource As String, _
    ByVal sellerTaxCode As String, _
    ByVal templateCode As String, _
    ByVal invoiceSeries As String, _
    ByVal invoiceNumber As String, _
    ByVal failureStage As String) As Long

    Dim data As Variant
    Dim rowIndex As Long
    Dim expectedKey As String
    Dim lastRow As Long

    lastRow = LastGdtErrorRow(ws)
    If lastRow < 2 Then Exit Function
    data = ws.Cells(2, 1).Resize(lastRow - 1, GDT_ERROR_COLUMN_COUNT).Value2
    expectedKey = BuildGdtErrorKey(direction, apiSource, sellerTaxCode, templateCode, _
                                   invoiceSeries, invoiceNumber, failureStage)

    For rowIndex = 1 To UBound(data, 1)
        If BuildGdtErrorKey(CStr(data(rowIndex, 3)), CStr(data(rowIndex, 4)), _
                            CStr(data(rowIndex, 5)), CStr(data(rowIndex, 6)), _
                            CStr(data(rowIndex, 7)), CStr(data(rowIndex, 8)), _
                            CStr(data(rowIndex, 10))) = expectedKey Then
            FindGdtErrorRow = rowIndex + 1
            Exit Function
        End If
    Next rowIndex
End Function

Private Function LastGdtErrorRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    LastGdtErrorRow = lastRow
End Function

Private Sub RenumberGdtErrorRows(ByVal ws As Worksheet)
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastGdtErrorRow(ws)
    For rowIndex = 2 To lastRow
        ws.Cells(rowIndex, 1).Value = rowIndex - 1
    Next rowIndex
End Sub

Private Function NormalizeKeyPart(ByVal value As String) As String
    NormalizeKeyPart = UCase$(Trim$(value))
End Function

Private Function RedactGdtSensitiveText(ByVal value As String) As String
    Dim lowerValue As String
    Dim markerPosition As Long

    lowerValue = LCase$(value)
    markerPosition = InStr(1, lowerValue, "authorization", vbTextCompare)
    If markerPosition > 0 Then
        RedactGdtSensitiveText = Left$(value, markerPosition - 1) & "[REDACTED]"
        Exit Function
    End If
    markerPosition = InStr(1, lowerValue, "bearer ", vbTextCompare)
    If markerPosition > 0 Then
        RedactGdtSensitiveText = Left$(value, markerPosition - 1) & "[REDACTED]"
        Exit Function
    End If
    RedactGdtSensitiveText = value
End Function
