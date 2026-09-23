Attribute VB_Name = "modParseIso"
Option Explicit

Public Declare PtrSafe Function SystemTimeToFileTime Lib _
  "kernel32" (lpSystemTime As SYSTEMTIME, _
  lpFileTime As FILETIME) As Long

Public Declare PtrSafe Function FileTimeToLocalFileTime Lib _
  "kernel32" (lpLocalFileTime As FILETIME, _
  lpFileTime As FILETIME) As Long

Public Declare PtrSafe Function FileTimeToSystemTime Lib _
  "kernel32" (lpFileTime As FILETIME, lpSystemTime _
  As SYSTEMTIME) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'---------------------------------------------------------------------
' Convert ISO8601 dateTimes to Excel Dates
'---------------------------------------------------------------------
Public Function ISODATE(iso As String)
    ' Find location of delimiters in input string
    Dim tPos As Integer: tPos = InStr(iso, "T")
    If tPos = 0 Then tPos = Len(iso) + 1
    Dim zPos As Integer: zPos = InStr(iso, "Z")
    If zPos = 0 Then zPos = InStr(iso, "+")
    If zPos = 0 Then zPos = InStr(tPos, iso, "-")
    If zPos = 0 Then zPos = Len(iso) + 1
    If zPos = tPos Then zPos = tPos + 1

    ' Get the relevant parts out
    Dim datePart As String: datePart = Mid(iso, 1, tPos - 1)
    Dim timePart As String: timePart = Mid(iso, tPos + 1, zPos - tPos - 1)
    Dim dotPos As Integer: dotPos = InStr(timePart, ".")
    If dotPos = 0 Then dotPos = Len(timePart) + 1
    timePart = Left(timePart, dotPos - 1)

    ' Have them parsed separately by Excel
    Dim d As Date: d = DateValue(datePart)
    Dim t As Date: If timePart <> "" Then t = TimeValue(timePart)
    Dim dt As Date: dt = d + t

    ' Add the timezone
    Dim tz As String: tz = Mid(iso, zPos)
    If tz <> "" And Left(tz, 1) <> "Z" Then
        Dim colonPos As Integer: colonPos = InStr(tz, ":")
        If colonPos = 0 Then colonPos = Len(tz) + 1

        Dim minutes As Integer: minutes = CInt(Mid(tz, 2, colonPos - 2)) * 60 + CInt(Mid(tz, colonPos + 1))
        If Left(tz, 1) = "+" Then minutes = -minutes
        dt = DateAdd("n", minutes, dt)
    End If

    ' Return value is the ISO8601 date in the local time zone
    dt = UTCToLocalTime(dt)
    ISODATE = Format(dt, "dd/mm/yyyy")
End Function

'---------------------------------------------------------------------
' Got this function to convert local date to UTC date from
' http://excel.tips.net/Pages/T002185_Automatically_Converting_to_GMT.html
'---------------------------------------------------------------------
Public Function UTCToLocalTime(dteTime As Date) As Date
    Dim infile As FILETIME
    Dim outfile As FILETIME
    Dim insys As SYSTEMTIME
    Dim outsys As SYSTEMTIME

    insys.wYear = CInt(Year(dteTime))
    insys.wMonth = CInt(Month(dteTime))
    insys.wDay = CInt(Day(dteTime))
    insys.wHour = CInt(Hour(dteTime))
    insys.wMinute = CInt(Minute(dteTime))
    insys.wSecond = CInt(Second(dteTime))

    Call SystemTimeToFileTime(insys, infile)
    Call FileTimeToLocalFileTime(infile, outfile)
    Call FileTimeToSystemTime(outfile, outsys)

    'Dung DateSerial/TimeSerial de khong phu thuoc dinh dang ngay cua Windows.
    'CDate("7/5/2024") co the bi hieu thanh 07/05 hoac 05/07 tuy locale.
    UTCToLocalTime = DateSerial(outsys.wYear, outsys.wMonth, outsys.wDay) + _
        TimeSerial(outsys.wHour, outsys.wMinute, outsys.wSecond)
End Function

'---------------------------------------------------------------------
' Tests for the ISO Date functions
'---------------------------------------------------------------------
Public Sub ISODateTest()
    ' [[ Verify that all dateTime formats parse sucesfully ]]
    Dim d1 As Date: d1 = ISODATE("2011-01-01")
    Dim d2 As Date: d2 = ISODATE("2011-01-01T00:00:00")
    Dim d3 As Date: d3 = ISODATE("2011-01-01T00:00:00Z")
    Dim d4 As Date: d4 = ISODATE("2011-01-01T12:00:00Z")
    Dim d5 As Date: d5 = ISODATE("2011-01-01T12:00:00+05:00")
    Dim d6 As Date: d6 = ISODATE("2011-01-01T12:00:00-05:00")
    Dim d7 As Date: d7 = ISODATE("2011-01-01T12:00:00.05381+05:00")
    AssertEqual "Date and midnight", d1, d2
    AssertEqual "With and without Z", d2, d3
    AssertEqual "With timezone", -5, DateDiff("h", d4, d5)
    AssertEqual "Timezone Difference", 10, DateDiff("h", d5, d6)
    AssertEqual "Ignore subsecond", d5, d7

    ' [[ Independence of local DST ]]
    ' Verify that a date in winter and a date in summer parse to the same Hour value
    Dim w As Date: w = ISODATE("2010-02-23T21:04:48+01:00")
    Dim s As Date: s = ISODATE("2010-07-23T21:04:48+01:00")
    AssertEqual "Winter/Summer hours", Hour(w), Hour(s)

    MsgBox "All tests passed succesfully!"
End Sub

Sub AssertEqual(name, X, Y)
    If X <> Y Then err.Raise 1234, Description:="Failed: " & name & ": '" & X & "' <> '" & Y & "'"
End Sub

