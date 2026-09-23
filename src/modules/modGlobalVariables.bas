Attribute VB_Name = "modGlobalVariables"
Option Explicit

#If VBA7 Or Win64 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public ckey As String, cToken As String, bearer As String, getStatus As String
Public js As Object
Public arrTrangThai As Variant, arrKQKTHoaDon As Variant   'Ket qua kiem tra hd
Public dicTenCotTC As Object    'Ten truong (cot) tra cuu
Public dicLink As Object            'Lay link tra cuu
Public sMST As String
Public oSaveUnzipFolder As Variant

Public Const ROWSTART As Long = 2   'Dong bat dau ghi du lieu

Public Const BASE_URL As String = "https://hoadondientu.gdt.gov.vn/api"
