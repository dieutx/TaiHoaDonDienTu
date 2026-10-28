Attribute VB_Name = "modGdtDatePeriods"
Option Explicit

' HAR 2026-10-28: query accepts monthly windows; sco-query reports 10 days.
' Windows count calendar days inclusively (00:00:00 through 23:59:59).
Public Function GdtListPeriods(ByRef months As Variant, ByVal source As String) As Variant
    Dim rows As New Collection, result() As Variant, period As Variant
    Dim i As Long, j As Long, first As Date, last As Date, finish As Date
    For i = 1 To UBound(months, 1)
        If source <> "sco-query" Then rows.Add Array(months(i, 1), months(i, 2), "query")
        If source <> "query" Then
            first = months(i, 1): finish = months(i, 2)
            Do
                last = first + 9
                If last > finish Then last = finish
                rows.Add Array(first, last, "sco-query")
                If last = finish Then Exit Do
                first = last + 1
            Loop
        End If
    Next i
    ReDim result(1 To rows.Count, 1 To 3)
    For i = 1 To rows.Count
        period = rows(i)
        For j = 1 To 3: result(i, j) = period(j - 1): Next j
    Next i
    GdtListPeriods = result
End Function

Public Function GdtReportedDayLimit(ByVal status As Long, ByVal message As String) As Long
    Dim re As Object, matches As Object
    If status <> 400 Then Exit Function
    If InStr(1, message, UniConvert("Khoarng thowfi gian tifm kieesm"), vbTextCompare) = 0 Then Exit Function
    If InStr(1, message, UniConvert("khoong dduwowjc lowsn hown"), vbTextCompare) = 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([0-9]{1,6})\s+" & UniConvert("ngafy")
    Set matches = re.Execute(message)
    If matches.Count > 0 Then GdtReportedDayLimit = CLng(matches(0).SubMatches(0))
End Function

' Replace only the rejected source window, keeping its pagination isolated.
Public Function GdtSplitListPeriod(ByRef periods As Variant, ByVal index As Long, ByVal days As Long) As Boolean
    Dim result() As Variant, count As Long, extra As Long, i As Long, j As Long, target As Long
    Dim first As Date, finish As Date, last As Date
    If days < 1 Then Exit Function
    first = periods(index, 1): finish = periods(index, 2)
    count = DateDiff("d", first, finish) + 1
    If days >= count Then Exit Function
    extra = (count - 1) \ days
    ReDim result(1 To UBound(periods, 1) + extra, 1 To 3)
    target = 1
    For i = 1 To UBound(periods, 1)
        If i = index Then
            Do
                last = first + days - 1
                If last > finish Then last = finish
                result(target, 1) = first: result(target, 2) = last
                result(target, 3) = periods(i, 3): target = target + 1
                If last = finish Then Exit Do
                first = last + 1
            Loop
        Else
            For j = 1 To 3: result(target, j) = periods(i, j): Next j
            target = target + 1
        End If
    Next i
    periods = result
    GdtSplitListPeriod = True
End Function
