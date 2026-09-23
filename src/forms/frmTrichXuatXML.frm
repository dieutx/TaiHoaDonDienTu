VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrichXuatXML 
   Caption         =   "+++"
   ClientHeight    =   4186
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   OleObjectBlob   =   "frmTrichXuatXML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTrichXuatXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChonFolder_Click()
    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .Title = "Chon thu muc de luu file Zip XML"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
        Me.txtXMLFolderPath = .SelectedItems(1) & "\"
    End With
End Sub

Private Sub cmdTrichXuat_Click()
    Application.Cursor = xlWait
    If Me.txtXMLFolderPath <> "" Then
        Call parseXML(Me.txtXMLFolderPath, Me.optMua.Value)
    End If
    Application.Cursor = xlDefault
End Sub

Private Sub UserForm_Initialize()
    Me.optMua.Value = True
    Me.txtXMLFolderPath = Environ("USERPROFILE") & "\Documents\"
End Sub

