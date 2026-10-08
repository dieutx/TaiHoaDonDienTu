Attribute VB_Name = "modGdtXmlScheduler"
Option Explicit

' Port of the bounded XML scheduling policy in hddt-downloader-windows.
' MSXML performs network I/O asynchronously; Excel and file writes stay serial.
'
' Throughput of this phase is bounded by the interval between two request
' starts, not by the number of connections: sliding the interval from 800 ms
' to 300 ms raises the ceiling from 1.25 to about 3.3 requests per second, and
' the extra connections keep that ceiling reachable when a single export takes
' longer than the interval.  This is deliberately faster than the 4 / 800 ms
' defaults of the reference tool; the 429 governor below (shared cooldown,
' one connection dropped per 30 s, interval x1.5 up to 5 s, recovery every
' 10 s) and the Retry-After handling are unchanged, so the load always comes
' back down automatically when GDT pushes back.
Public Const GDT_XML_MAX_CONCURRENCY As Long = 6
Public Const GDT_XML_CONCURRENCY As Long = 6
Public Const GDT_XML_INTERVAL_MS As Long = 300

' JSON phases share the bounded HTTP task and throttle; Excel writes stay serial.
Public Const GDT_JSON_CONCURRENCY As Long = 4
Public Const GDT_JSON_BUFFER_SIZE As Long = 12

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If

Public Sub GdtPumpWait()
    DoEvents
    Sleep 10
End Sub

Public Function GdtClockSeconds() As Double
    Dim today As Date, ticks As Double
    today = Date
    ticks = Timer
    If today <> Date Then today = Date: ticks = Timer
    GdtClockSeconds = CDbl(today) * 86400# + ticks
End Function
