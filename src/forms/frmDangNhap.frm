VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDangNhap 
   Caption         =   "+++"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "frmDangNhap.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDangNhap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CAPTCHA_URL  As String = "/captcha"
Private Const LOGIN_URL    As String = "/security-taxpayer/authenticate"
Private Const PROFILE_URL    As String = "/profile"

Dim res As String, blnCaptcha As Boolean
Dim arrUsers

Private Sub cmdDauMK_Click()
    Me.txtPass.PasswordChar = "*"
    Me.cmdHienMK.Visible = True
    Me.cmdDauMK.Visible = False
End Sub

Private Sub cmdHienMK_Click()
    Me.txtPass.PasswordChar = ""
    Me.cmdHienMK.Visible = False
    Me.cmdDauMK.Visible = True
End Sub

Private Sub lblCapNhatMK_Click()
    If Me.txtUser <> "" Then
        sMST = Me.txtUser
        frmCapNhatMK.Show
    Else
        MsgBoxUni "T" & ChrW(234) & "n " & ChrW(273) & ChrW(259) & "ng nh" & ChrW(7853) & "p kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c " & ChrW(273) & ChrW(7875) & " tr" & ChrW(7889) & "ng.", vbCritical
        Me.txtUser.SetFocus
        Exit Sub
    End If
End Sub

Private Sub lsbUsers_Change()
    Dim user As String, pass As String, tenDV As String
    user = Me.lsbUsers.list(Me.lsbUsers.ListIndex)
    pass = Me.lsbUsers.list(Me.lsbUsers.ListIndex, 1)
    tenDV = Me.lsbUsers.list(Me.lsbUsers.ListIndex, 2)
    Me.txtUser = user: Me.txtPass = pass: Me.txtTenDV = tenDV: Me.lblTenDV.caption = tenDV
End Sub

Private Sub lsbUsers_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.lsbUsers.Visible = False
End Sub

Private Sub txtPass_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.lsbUsers.Visible = False
    If layCaptcha Then
        blnCaptcha = True
    Else
        blnCaptcha = False
    End If
    Me.cmdDangNhap.SetFocus
End Sub

Private Sub txtUser_AfterUpdate()
    On Error Resume Next    'Vuot loi could not set property visible
    Me.lsbUsers.Visible = False
End Sub

Private Sub txtUser_Change()
    Me.lsbUsers.Visible = True
End Sub

Private Sub UserForm_Initialize()
    Me.lblThanhCong.Visible = False
    Me.lblThatBai.Visible = False
    Me.cmdHienMK.Visible = True
    Me.cmdDauMK.Visible = False
    layUserPass
End Sub

Private Sub txtPass_AfterUpdate()
    If layCaptcha Then
        blnCaptcha = True
    Else
        blnCaptcha = False
    End If
End Sub

Private Sub cmdDangNhap_Click()
    Application.ScreenUpdating = False
    
    If blnCaptcha = False Then Exit Sub
    
    Dim http As Object, formData As String
    formData = "{""username"":""" & Me.txtUser & """,""password"":""" & Me.txtPass & """,""cvalue"":""" & Me.txtNhapCaptcha & """,""ckey"":""" & ckey & """}"
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", BASE_URL & LOGIN_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Request-Id", CreateRequestID
    http.setRequestHeader "User-Agent", "Mozilla/5.0 ExcelVBA-HDDT"
    On Error GoTo HttpErr
    http.send formData
    On Error GoTo 0
    
    If http.Status <> 200 Then
        'LogMsg "Dang nhap that bai. HTTP " & http.Status & " - " & http.responseText
        MsgBoxUni UniConvert("DDawng nhaajp khoong thafnh coong (HTTP " & http.Status & ").") & vbCrLf & vbCrLf & _
            UniConvert("Kieerm tra MST/maajt khaaru hoawjc CAPTCHA roofi thuwr laji."), vbCritical
        Exit Sub
    End If
    
    res = http.responseText
    Set js = CreateObject("Scripting.Dictionary")
    Set js = JsonConverter.ParseJSON(res)
    
    If InStr(1, res, "message") > 0 Then
        If InStr(1, res, "captcha") Then
            MsgBoxUni js("message") & vbCrLf & "Vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p l" & ChrW(7841) & "i.", vbCritical, "Thông báo"
        Else
            MsgBoxUni js("message")
        End If

        Me.lblThanhCong.Visible = False
        Me.lblThatBai.Visible = True
        Exit Sub
    Else
        cToken = js("token")
        bearer = cToken
        
        Me.lblThanhCong.Visible = True
        Me.lblThatBai.Visible = False
        
        ' Luu User/pass neu chua co trong danh sach, va Token
        Call saveUserPassToken
        
        Application.OnTime Time + TimeValue("00:00:01"), "closeForm"
    End If
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
HttpErr:
    'LogMsg "Loi mang khi dang nhap: " & err.Description
    MsgBox "Loi mang: " & err.Description, vbCritical

End Sub

Function layCaptcha() As Boolean
    layCaptcha = True
    'Kiem tra ket noi internet
    If GetInternetConnectedState = False Then
        MsgBoxUni "Kh" & ChrW(244) & "ng c" & ChrW(243) & " k" & ChrW(7871) & "t n" & ChrW(7889) & "i internet.", vbCritical, "No internet"
        layCaptcha = False
        Exit Function
    End If
    
    Dim http As Object, cContent As String, sMsg As String
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", BASE_URL & CAPTCHA_URL, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Request-Id", CreateRequestID
    On Error GoTo HttpErr
    http.send
    On Error GoTo 0
    
    If http.Status <> 200 Then
        sMsg = "C" & ChrW(243) & " l" & ChrW(7895) & "i k" & ChrW(7871) & "t n" & ChrW(7889) & "i v" & ChrW(224) & " l" & ChrW(7845) & "y d" & ChrW(7919) & " li" & ChrW(7879) & "u."
        sMsg = sMsg & vbCrLf & "Vui l" & ChrW(242) & "ng kh" & ChrW(7903) & "i " & ChrW(273) & ChrW(7897) & "ng l" & ChrW(7841) & "i " & ChrW(7913) & "ng d" & ChrW(7909) & "ng."
        MsgBoxUni sMsg, vbCritical
        layCaptcha = False
        Exit Function
    End If
    
    Set js = JsonConverter.ParseJSON(http.responseText)
    ckey = js("key")
    cContent = js("content")
    
    DoEvents
    Me.txtNhapCaptcha = detectSVGCaptcha(http.responseText)
    
    Exit Function
    
HttpErr:
    MsgBox "Loi mang khi lay CAPTCHA: " & err.Description
    layCaptcha = False
    
End Function

Sub layUserPass()
    arrUsers = ThisWorkbook.Sheets("MENU").Range("A7:C" & ThisWorkbook.Sheets("MENU").Cells(Rows.count, "A").End(xlUp).row).Value
    Me.lsbUsers.list = arrUsers
    Me.lsbUsers.Width = Me.txtUser.Width
    Me.lsbUsers.Height = 100
    Me.lsbUsers.Left = Me.txtUser.Left
    Me.lsbUsers.Top = Me.txtUser.Top + Me.txtUser.Height
End Sub

Sub saveUserPassToken()
    Dim X As Long, lr As Long
    For X = 1 To UBound(arrUsers, 1)
        If Me.txtUser = arrUsers(X, 1) Then
            'Chi luu Token
            With ThisWorkbook.Sheets("MENU")
                .Range("D" & X + 6).Value = cToken
                .Range("E" & X + 6).Value = Now
                ThisWorkbook.Save
            End With
            Exit Sub
        End If
    Next
    'Luu user/pass moi/ Token
    With ThisWorkbook.Sheets("MENU")
        lr = .Cells(Rows.count, "A").End(xlUp).row
        .Range("A" & lr + 1).Value = Me.txtUser
        .Range("B" & lr + 1).Value = Me.txtPass
        .Range("C" & lr + 1).Value = Me.txtTenDV
        .Range("D" & lr + 1).Value = cToken
        .Range("E" & lr + 1).Value = Now
        ThisWorkbook.Save
    End With
End Sub
