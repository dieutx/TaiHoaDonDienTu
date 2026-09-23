VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "Date picker"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "frmDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTarget As Object
Private mHandlers As Collection

Private Sub UserForm_Initialize()
    Me.Caption = UniConvert("Chojn ngafy")
    BuildPickerControls
    SetPickerDate Date
End Sub

Private Sub BuildPickerControls()
    Dim ctl As Object, i As Long, h As clsUiButtonHandler
    Set mHandlers = New Collection
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblDay", True): ctl.caption = UniConvert("Ngafy"): ctl.Left = 18: ctl.Top = 12: ctl.Width = 48: ctl.Height = 18
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblMonth", True): ctl.caption = UniConvert("Thasng"): ctl.Left = 78: ctl.Top = 12: ctl.Width = 72: ctl.Height = 18
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblYear", True): ctl.caption = UniConvert("Nawm"): ctl.Left = 168: ctl.Top = 12: ctl.Width = 72: ctl.Height = 18
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", "cboDay", True): ctl.Left = 18: ctl.Top = 33: ctl.Width = 48: ctl.Height = 22
    For i = 1 To 31: ctl.AddItem Format$(i, "00"): Next i
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", "cboMonth", True): ctl.Left = 78: ctl.Top = 33: ctl.Width = 72: ctl.Height = 22
    For i = 1 To 12: ctl.AddItem Format$(i, "00"): Next i
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", "cboYear", True): ctl.Left = 168: ctl.Top = 33: ctl.Width = 72: ctl.Height = 22
    For i = Year(Date) - 10 To Year(Date) + 10: ctl.AddItem CStr(i): Next i
    AddPickerButton "cmdToday", UniConvert("Hoom nay"), 18, "SelectToday"
    AddPickerButton "cmdOK", UniConvert("Chojn"), 102, "AcceptDate"
    AddPickerButton "cmdCancel", UniConvert("Huyr"), 180, "CancelPicker"
End Sub

Private Sub AddPickerButton(ByVal controlName As String, ByVal caption As String, ByVal leftPos As Single, ByVal action As String)
    Dim ctl As Object, h As clsUiButtonHandler
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    ctl.caption = caption: ctl.Left = leftPos: ctl.Top = 75: ctl.Width = 66: ctl.Height = 27
    Set h = New clsUiButtonHandler: Set h.Button = ctl: Set h.Owner = Me: h.ActionName = action: mHandlers.Add h
End Sub

Public Sub SetTarget(ByVal targetControl As Object)
    Set mTarget = targetControl
    If IsDate(targetControl.Value) Then SetPickerDate CDate(targetControl.Value)
End Sub

Private Sub SetPickerDate(ByVal selectedDate As Date)
    Me.Controls("cboDay").Value = Format$(Day(selectedDate), "00")
    Me.Controls("cboMonth").Value = Format$(Month(selectedDate), "00")
    Me.Controls("cboYear").Value = CStr(Year(selectedDate))
End Sub

Public Sub SelectToday()
    SetPickerDate Date
End Sub

Public Sub AcceptDate()
    Dim selectedDate As Date
    On Error GoTo InvalidDate
    selectedDate = DateSerial(CLng(Me.Controls("cboYear").Value), CLng(Me.Controls("cboMonth").Value), CLng(Me.Controls("cboDay").Value))
    If Day(selectedDate) <> CLng(Me.Controls("cboDay").Value) Or Month(selectedDate) <> CLng(Me.Controls("cboMonth").Value) Then GoTo InvalidDate
    mTarget.Value = Format$(selectedDate, "dd/mm/yyyy")
    frmTaiHoaDon.RefreshPickedDate mTarget.name
    Unload Me
    Exit Sub
InvalidDate:
    MsgBoxUni UniConvert("Ngafy ddax chojn khoong howjp leej."), vbExclamation, UniConvert("Chojn ngafy")
End Sub

Public Sub CancelPicker()
    Unload Me
End Sub
