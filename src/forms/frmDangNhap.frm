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
Private mLoginBusy As Boolean

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
End Sub

Private Sub txtUser_AfterUpdate()
    On Error Resume Next    'Vuot loi could not set property visible
    Me.lsbUsers.Visible = False
End Sub

Private Sub txtUser_Change()
    InvalidateCaptcha
    Me.lsbUsers.Visible = True
End Sub

Private Sub UserForm_Initialize()
    Me.lblThanhCong.Visible = False
    Me.lblThatBai.Visible = False
    Me.cmdHienMK.Visible = True
    Me.cmdDauMK.Visible = False
    InvalidateCaptcha
    layUserPass
End Sub

Private Sub txtPass_Change()
    InvalidateCaptcha
End Sub

Private Sub InvalidateCaptcha()
    blnCaptcha = False
    ckey = vbNullString
    Me.txtNhapCaptcha.Value = vbNullString
End Sub

Private Sub cmdDangNhap_Click()
    Dim http As Object, formData As String, sessionToken As String
    Dim originalCaption As String
    If mLoginBusy Then Exit Sub
    On Error GoTo HttpErr
    originalCaption = Me.cmdDangNhap.Caption
    mLoginBusy = True
    Me.cmdDangNhap.Enabled = False
    Me.lsbUsers.Visible = False
    Me.lblThanhCong.Visible = False
    Me.lblThatBai.Visible = False

    If Len(Trim$(CStr(Me.txtUser.Value))) = 0 Or Len(CStr(Me.txtPass.Value)) = 0 Then
        MsgBoxUni UniConvert("Vui lofng nhaajp teen ddawng nhaajp vaf maajt khaaru."), vbExclamation
        GoTo CleanExit
    End If

    If Not blnCaptcha Or Len(ckey) = 0 Or Len(Trim$(CStr(Me.txtNhapCaptcha.Value))) = 0 Then
        Me.cmdDangNhap.Caption = UniConvert("DDang laays max...")
        Me.Repaint
        blnCaptcha = layCaptcha()
        If Not blnCaptcha Then GoTo CleanExit
    End If

    Me.cmdDangNhap.Caption = UniConvert("DDang ddawng nhaajp...")
    Me.Repaint
    formData = BuildGdtLoginJson(CStr(Me.txtUser.Value), CStr(Me.txtPass.Value), CStr(Me.txtNhapCaptcha.Value), CStr(ckey))
    ' A submitted challenge must not be reused after a rejection or timeout.
    blnCaptcha = False
    ckey = vbNullString
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 15000, 15000
    http.Open "POST", BASE_URL & LOGIN_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Request-Id", CreateRequestID
    http.setRequestHeader "User-Agent", "Mozilla/5.0 ExcelVBA-HDDT"
    On Error GoTo HttpErr
    http.send formData
    
    If http.Status <> 200 Then
        'LogMsg "Dang nhap that bai. HTTP " & http.Status & " - " & http.responseText
        MsgBoxUni UniConvert("DDawng nhaajp khoong thafnh coong (HTTP " & http.Status & ").") & vbCrLf & vbCrLf & _
            UniConvert("Kieerm tra MST/maajt khaaru hoawjc CAPTCHA roofi thuwr laji."), vbCritical
        GoTo CleanExit
    End If
    
    res = http.responseText
    On Error GoTo ResponseErr
    Set js = ParseGdtAuthObject(res)
    
    If GdtJsonHasValue(js, "message") Then
        If InStr(1, res, "captcha") Then
            MsgBoxUni js("message") & vbCrLf & "Vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p l" & ChrW(7841) & "i.", vbCritical, "Thông báo"
        Else
            MsgBoxUni js("message")
        End If

        Me.lblThanhCong.Visible = False
        Me.lblThatBai.Visible = True
        GoTo CleanExit
    Else
        sessionToken = GdtRequiredAuthText(js, "token")
        cToken = sessionToken
        bearer = cToken
        
        Me.lblThanhCong.Visible = True
        Me.lblThatBai.Visible = False
        
        ' Luu User/pass neu chua co trong danh sach, va Token
        On Error GoTo SaveErr
        Call saveUserPassToken
        
        Application.OnTime Time + TimeValue("00:00:01"), "closeForm"
    End If
    
    Application.ScreenUpdating = True
    
    GoTo CleanExit
    
HttpErr:
    'LogMsg "Loi mang khi dang nhap: " & err.Description
    MsgBox "Loi mang: " & err.Description, vbCritical
    GoTo CleanExit

ResponseErr:
    Me.lblThanhCong.Visible = False
    Me.lblThatBai.Visible = True
    MsgBoxUni UniConvert("Pharn hoofi ddawng nhaajp khoong howjp leej. Vui lofng thuwr laji."), vbCritical
    GoTo CleanExit

SaveErr:
    MsgBoxUni UniConvert("DDax ddawng nhaajp nhuwng khoong luwu dduwowjc file. Kieerm tra quyeen ghi vaf thuwr laji."), vbExclamation

CleanExit:
    Me.lblThatBai.Visible = Not Me.lblThanhCong.Visible
    Me.cmdDangNhap.Caption = originalCaption
    Me.cmdDangNhap.Enabled = True
    mLoginBusy = False
End Sub

Function layCaptcha() As Boolean
    InvalidateCaptcha
    layCaptcha = False
    'Kiem tra ket noi internet
    If GetInternetConnectedState = False Then
        MsgBoxUni "Kh" & ChrW(244) & "ng c" & ChrW(243) & " k" & ChrW(7871) & "t n" & ChrW(7889) & "i internet.", vbCritical, "No internet"
        layCaptcha = False
        Exit Function
    End If
    
    Dim http As Object, cContent As String, sMsg As String, newKey As String, captchaText As String
    On Error GoTo HttpErr
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 15000, 15000
    http.Open "GET", BASE_URL & CAPTCHA_URL, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Request-Id", CreateRequestID
    On Error GoTo HttpErr
    http.send
    
    If http.Status <> 200 Then
        sMsg = "C" & ChrW(243) & " l" & ChrW(7895) & "i k" & ChrW(7871) & "t n" & ChrW(7889) & "i v" & ChrW(224) & " l" & ChrW(7845) & "y d" & ChrW(7919) & " li" & ChrW(7879) & "u."
        sMsg = sMsg & " (HTTP " & http.Status & ")." & vbCrLf & UniConvert("Vui lofng baasm DDawng nhaajp ddeer thuwr laji.")
        MsgBoxUni sMsg, vbCritical
        layCaptcha = False
        Exit Function
    End If
    
    On Error GoTo ResponseErr
    Set js = ParseGdtAuthObject(http.responseText)
    newKey = GdtRequiredAuthText(js, "key")
    cContent = GdtRequiredAuthText(js, "content")
    
    captchaText = detectSVGCaptcha(http.responseText)
    If Len(Trim$(captchaText)) <> 6 Then
        MsgBoxUni UniConvert("Khoong nhaajn dajng dduwowjc CAPTCHA. Vui lofng baasm DDawng nhaajp ddeer laays max mowsi."), vbExclamation
        Exit Function
    End If
    Me.txtNhapCaptcha = captchaText
    ckey = newKey
    layCaptcha = True
    
    Exit Function
    
HttpErr:
    MsgBox "Loi mang khi lay CAPTCHA: " & err.Description
    layCaptcha = False
    Exit Function

ResponseErr:
    MsgBoxUni UniConvert("Pharn hoofi CAPTCHA khoong howjp leej. Vui lofng laays CAPTCHA mowsi."), vbExclamation
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
