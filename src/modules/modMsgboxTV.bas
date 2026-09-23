Attribute VB_Name = "modMsgboxTV"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function MessageBoxW Lib "user32" _
        (ByVal hwnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal wType As Long) As Long
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Public Declare Function MessageBoxW Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal wType As Long) As Long
#End If

Public Function MsgBoxUni(ByVal sMsgUni As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal sTitleUni As String = "Thông báo") As VbMsgBoxResult
    MsgBoxUni = MessageBoxW(GetActiveWindow, StrPtr(sMsgUni), StrPtr(sTitleUni), Buttons)
End Function

Function UniConvert(text As String, Optional InputMethod As String = "Telex") As String
    Dim VNI_Type, Telex_Type, CharCode, Temp, i As Long
    UniConvert = text
    VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
        "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
        "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
        "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
        "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
    Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", _
        "ees", "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", _
        "ows", "owf", "owr", "owx", "owj", "uws", "uwf", "uwr", "uwx", "uwj", _
        "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", "e_s", "ef", _
        "e_r", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", _
        "of", "or", "ox", "oj", "oo", "ow", "us", "uf", "ur", "ux", _
        "uj", "uw", "ys", "yf", "yr", "yx", "yj", "AWS", "AWF", "AWR", _
        "AWX", "AWJ", "AAS", "AAF", "AAR", "AAX", "AAJ", "EES", "EEF", "EER", _
        "EEX", "EEJ", "OOS", "OOF", "OOR", "OOX", "OOJ", "OWS", "OWF", "OWR", _
        "OWX", "OWJ", "UWS", "UWF", "UWR", "UWX", "UWJ", "AS", "AF", "AR", _
        "AX", "AJ", "AW", "AA", "DD", "E_S", "EF", "E_R", "EX", "EJ", _
        "EE", "IS", "IF", "IR", "IX", "IJ", "OS", "OF", "OR", "OX", _
        "OJ", "OO", "OW", "US", "UF", "UR", "UX", "UJ", "UW", "YS", _
        "YF", "YR", "YX", "YJ")
    'Luu y: them "_" thanh "e_r", "E_R" de tranh loi "R" bien thanh dau hoi vai truong hop; E_S
    
    CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), ChrW(7849), ChrW(7851), ChrW(7853), _
        ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), _
        ChrW(7899), ChrW(7901), ChrW(7903), ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), _
        ChrW(225), ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
        ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), ChrW(243), _
        ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), ChrW(7911), ChrW(361), _
        ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925), ChrW(7854), ChrW(7856), ChrW(7858), _
        ChrW(7860), ChrW(7862), ChrW(7844), ChrW(7846), ChrW(7848), ChrW(7850), ChrW(7852), ChrW(7870), ChrW(7872), ChrW(7874), _
        ChrW(7876), ChrW(7878), ChrW(7888), ChrW(7890), ChrW(7892), ChrW(7894), ChrW(7896), ChrW(7898), ChrW(7900), ChrW(7902), _
        ChrW(7904), ChrW(7906), ChrW(7912), ChrW(7914), ChrW(7916), ChrW(7918), ChrW(7920), ChrW(193), ChrW(192), ChrW(7842), _
        ChrW(195), ChrW(7840), ChrW(258), ChrW(194), ChrW(272), ChrW(201), ChrW(200), ChrW(7866), ChrW(7868), ChrW(7864), _
        ChrW(202), ChrW(205), ChrW(204), ChrW(7880), ChrW(296), ChrW(7882), ChrW(211), ChrW(210), ChrW(7886), ChrW(213), _
        ChrW(7884), ChrW(212), ChrW(416), ChrW(218), ChrW(217), ChrW(7910), ChrW(360), ChrW(7908), ChrW(431), ChrW(221), _
        ChrW(7922), ChrW(7926), ChrW(7928), ChrW(7924))
    
    Select Case InputMethod
        Case Is = "VNI": Temp = VNI_Type
        Case Is = "Telex": Temp = Telex_Type
    End Select
    For i = 0 To UBound(CharCode)
        UniConvert = Replace(UniConvert, Temp(i), CharCode(i), , , vbBinaryCompare)
        'UniConvert = Replace(UniConvert, UCase(Temp(i)), UCase(CharCode(i)))
    Next i
End Function

