VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCapNhatMK 
   Caption         =   "+++"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "frmCapNhatMK.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCapNhatMK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrMatKhau
Dim r As Long, khongtimthay As Boolean

Private Sub cmdCapNhat_Click()
    If khongtimthay = True Then
        MsgBoxUni "MST n" & ChrW(224) & "y ch" & ChrW(432) & "a c" & ChrW(243) & " trong file.", vbExclamation
        Exit Sub
    End If
    If Me.txtMKCu = "" Or Me.txtMKMoi = "" Or Me.txtNhapLaiMKMoi = "" Then
        MsgBoxUni "Ph" & ChrW(7843) & "i nh" & ChrW(7853) & "p " & ChrW(273) & ChrW(7847) & "y " & ChrW(273) & ChrW(7911) & " th" & ChrW(244) & "ng tin.", vbExclamation, "Thông báo"
        Me.txtMKCu.SetFocus
        Exit Sub
    End If
    If Me.txtNhapLaiMKMoi <> Me.txtMKMoi Then
        MsgBoxUni "M" & ChrW(7853) & "t kh" & ChrW(7849) & "u kh" & ChrW(244) & "ng tr" & ChrW(249) & "ng kh" & ChrW(7899) & "p!", vbExclamation, "Thông báo"
        Me.txtNhapLaiMKMoi = ""
        Me.txtNhapLaiMKMoi.SetFocus
        Exit Sub
    End If
    If Me.txtMKCu <> Me.txtPass Then
        MsgBoxUni "M" & ChrW(7853) & "t kh" & ChrW(7849) & "u c" & ChrW(361) & " kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(250) & "ng.", vbCritical
        Exit Sub
    End If
    
    Call luuMatKhau
    
End Sub

Sub luuMatKhau()
    ThisWorkbook.Sheets("MENU").Range("B" & r + 6).Value = Me.txtMKMoi
    MsgBoxUni "C" & ChrW(7853) & "p nh" & ChrW(7853) & "t m" & ChrW(7853) & "t kh" & ChrW(7849) & "u th" & ChrW(224) & "nh c" & ChrW(244) & "ng.", vbInformation, "Thành công"
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    arrMatKhau = ThisWorkbook.Sheets("MENU").Range("A7:B" & ThisWorkbook.Sheets("MENU").Cells(Rows.count, "A").End(xlUp).row).Value
    Me.txtMST = sMST
    For r = 1 To UBound(arrMatKhau)
        If arrMatKhau(r, 1) = Me.txtMST Then
            Me.txtPass = arrMatKhau(r, 2)
            khongtimthay = False
            Exit For
        Else
            khongtimthay = True
        End If
    Next
End Sub

