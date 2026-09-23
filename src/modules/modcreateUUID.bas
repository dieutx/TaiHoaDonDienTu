Attribute VB_Name = "modcreateUUID"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef pbBuffer As Any, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long
#Else
    Private Declare Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByRef pbBuffer As Any, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long
#End If

Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2

'========================================================
' Tao UUID v4 tuong ung voi crypto.randomUUID() trong file js
' Vd: 8b7e4c91-3d25-4f6a-9b12-73c8a5e21d04
' Dung chuan UUIDv4 (nhóm 3 luôn bat dau la so 4, nhom 4:8,9,a,b)
'========================================================
Public Function CreateRequestID() As String

    Dim b(0 To 15) As Byte
    Dim ret As Long

    ret = BCryptGenRandom( _
            0, _
            b(0), _
            16, _
            BCRYPT_USE_SYSTEM_PREFERRED_RNG)

    If ret <> 0 Then
        err.Raise vbObjectError + 1001, _
                  "CreateRequestID", _
                  "BCryptGenRandom failed. Error = " & ret
    End If

    'UUID version 4
    'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx
    b(6) = (b(6) And &HF) Or &H40
    b(8) = (b(8) And &H3F) Or &H80

    CreateRequestID = _
        ByteHex(b(0)) & ByteHex(b(1)) & _
        ByteHex(b(2)) & ByteHex(b(3)) & "-" & _
        ByteHex(b(4)) & ByteHex(b(5)) & "-" & _
        ByteHex(b(6)) & ByteHex(b(7)) & "-" & _
        ByteHex(b(8)) & ByteHex(b(9)) & "-" & _
        ByteHex(b(10)) & ByteHex(b(11)) & _
        ByteHex(b(12)) & ByteHex(b(13)) & _
        ByteHex(b(14)) & ByteHex(b(15))

End Function

'========================================================
' Byte -> 2 ký tu HEX
'========================================================
Private Function ByteHex(ByVal b As Byte) As String

    Const HEX As String = "0123456789abcdef"

    ByteHex = _
        Mid$(HEX, (b \ 16) + 1, 1) & _
        Mid$(HEX, (b And &HF) + 1, 1)

End Function

