Attribute VB_Name = "modUiTheme"
Option Explicit

' Shared visual language for frmTaiHoaDon: a light enterprise dashboard on a
' pale grey page, white surfaces, one blue accent and semantic state colours.
'
' The colour values are the numbers VBA stores for RGB(r, g, b), written as
' literals so they can be used in Const declarations and stay ASCII in source
' control.  The comment next to each value is the RGB(...) it represents.

Public Const UI_FONT As String = "Segoe UI"

Public Const UI_PAGE_BG As Long = 16381684       ' RGB(244, 246, 249)
Public Const UI_CARD As Long = 16777215          ' RGB(255, 255, 255)
Public Const UI_BORDER As Long = 15788258        ' RGB(226, 232, 240)
Public Const UI_TRACK As Long = 15460325         ' RGB(229, 231, 235)
Public Const UI_TEXT As Long = 3615007           ' RGB(31, 41, 55)
Public Const UI_MUTED As Long = 8417899          ' RGB(107, 114, 128)
Public Const UI_BLUE As Long = 15426341          ' RGB(37, 99, 235)
Public Const UI_BLUE_DARK As Long = 14175773     ' RGB(29, 78, 216)
Public Const UI_BLUE_SOFT As Long = 16706267     ' RGB(219, 234, 254)
Public Const UI_INFO_BG As Long = 16774895       ' RGB(239, 246, 255)
Public Const UI_SUCCESS As Long = 4030485        ' RGB(21, 128, 61)
Public Const UI_SUCCESS_SOFT As Long = 15203548  ' RGB(220, 252, 231)
Public Const UI_WARN As Long = 934034            ' RGB(146, 64, 14)
Public Const UI_WARN_SOFT As Long = 13104126     ' RGB(254, 243, 199)
Public Const UI_DANGER As Long = 2500316         ' RGB(220, 38, 38)
Public Const UI_DANGER_SOFT As Long = 14869246   ' RGB(254, 226, 226)

' A remaining-time estimate is only shown once it carries information.
Public Const UI_ETA_MIN_ITEMS As Long = 8
Public Const UI_ETA_MIN_SECONDS As Double = 10#

' One decimal place, built with integer arithmetic so the text does not depend
' on the Windows decimal separator: 75 of 862 reads 8,7%, not 8.7% or 18%.
Public Function UiPercentText(ByVal part As Long, ByVal total As Long) As String
    Dim tenths As Long
    If total <= 0 Then Exit Function
    If part < 0 Then part = 0
    If part > total Then part = total
    tenths = CLng(Int(CDbl(part) * 1000# / CDbl(total) + 0.5))
    UiPercentText = CStr(tenths \ 10) & "," & CStr(tenths Mod 10) & "%"
End Function

' An empty result means "not reliable yet"; the caller keeps its placeholder.
Public Function UiEtaText(ByVal done As Long, ByVal total As Long, ByVal elapsedSeconds As Double) As String
    Dim remaining As Long
    If total <= 0 Or done <= 0 Then Exit Function
    If done < UI_ETA_MIN_ITEMS Or elapsedSeconds < UI_ETA_MIN_SECONDS Then Exit Function
    If done >= total Then
        UiEtaText = "00:00:00"
        Exit Function
    End If
    remaining = total - done
    UiEtaText = GdtElapsedText(Int((elapsedSeconds / done) * remaining))
End Function

' Font, weight and colour in one call; every MSForms control used on the form
' exposes Font and ForeColor.
Public Sub UiFont(ByVal ctl As Object, ByVal size As Single, ByVal bold As Boolean, ByVal color As Long)
    ctl.Font.Name = UI_FONT
    ctl.Font.Size = size
    ctl.Font.Bold = bold
    ctl.ForeColor = color
End Sub
