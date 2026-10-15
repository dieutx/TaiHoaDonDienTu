Attribute VB_Name = "modGdtJson"
Option Explicit

Public Sub RequireGdtDetailIdentity(ByVal invoice As Object, ByVal seller As String, ByVal template As String, _
    ByVal series As String, ByVal number As String)
    If GdtInvoiceIdentity(SafeJsonText(invoice, "nbmst"), SafeJsonText(invoice, "khmshdon"), _
        SafeJsonText(invoice, "khhdon"), SafeJsonText(invoice, "shdon")) <> _
        GdtInvoiceIdentity(seller, template, series, number) Then Err.Raise 5, , "API detail invoice identity mismatch"
End Sub

Public Function GdtInvoiceStatus(ByVal raw As String) As Long
    Dim value As Double
    GdtInvoiceStatus = -1
    If Not IsNumeric(raw) Then Exit Function
    On Error GoTo Unknown
    value = CDbl(raw)
    If value >= 0 And value <= 6 And value = Fix(value) Then GdtInvoiceStatus = CLng(value)
Unknown:
End Function

Public Function GdtJsonNumber(ByVal value As Variant, ByVal field As String) As Double
    Static re As Object
    If IsObject(value) Or IsArray(value) Or IsNull(value) Or IsEmpty(value) Then GoTo Invalid
    If VarType(value) = vbString Then
        If re Is Nothing Then
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([Ee][+-]?[0-9]+)?$"
        End If
        If Not re.Test(Trim$(CStr(value))) Then GoTo Invalid
        GdtJsonNumber = Val(Trim$(CStr(value)))
    ElseIf VarType(value) <> vbBoolean And IsNumeric(value) Then
        GdtJsonNumber = CDbl(value)
    Else
        GoTo Invalid
    End If
    Exit Function
Invalid:
    Err.Raise 5, "GdtJsonNumber", "Invalid JSON number: " & field
End Function

Public Sub GdtWriteJsonNumberCell(ByVal cell As Range, ByVal value As Variant, ByVal field As String)
    If IsNull(value) Or IsEmpty(value) Then Exit Sub
    If VarType(value) = vbString Then
        If Len(Trim$(CStr(value))) = 0 Then Exit Sub
    End If
    cell.Value2 = GdtJsonNumber(value, field)
End Sub

' Optional JSON arrays may be absent, null or empty. Never add missing keys
' by reading Dictionary.Item before Exists.
Public Function GdtJsonArray(ByVal value As Object, ByVal key As String) As Collection
    Set GdtJsonArray = New Collection
    If value Is Nothing Then Exit Function
    If Not value.Exists(key) Then Exit Function
    If Not IsObject(value(key)) Then Exit Function
    If TypeName(value(key)) <> "Collection" Then Err.Raise 5, "GdtJsonArray", "Expected array: " & key
    Set GdtJsonArray = value(key)
End Function

Public Function GdtJsonHasValue(ByVal value As Object, ByVal key As String) As Boolean
    If Not value.Exists(key) Then Exit Function
    If IsNull(value(key)) Or IsEmpty(value(key)) Then Exit Function
    If IsObject(value(key)) Then Exit Function
    GdtJsonHasValue = (Len(CStr(value(key))) > 0)
End Function

Public Function GdtJsonHasError(ByVal value As Object) As Boolean
    If TypeName(value) <> "Dictionary" Then Exit Function
    If Not value.Exists("error") Then Exit Function
    If IsNull(value("error")) Or IsEmpty(value("error")) Then Exit Function
    If IsObject(value("error")) Then
        GdtJsonHasError = True
    Else
        GdtJsonHasError = (Len(CStr(value("error"))) > 0 And CStr(value("error")) <> "False")
    End If
End Function

Public Function BuildGdtLoginJson(ByVal username As String, ByVal password As String, _
    ByVal captcha As String, ByVal captchaKey As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "username", username
    payload.Add "password", password
    payload.Add "cvalue", captcha
    payload.Add "ckey", captchaKey
    BuildGdtLoginJson = JsonConverter.ConvertToJson(payload)
End Function

' Never echo authentication response bodies: they can contain credentials.
Public Function ParseGdtAuthObject(ByVal jsonText As String) As Object
    On Error GoTo InvalidResponse
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJSON(jsonText)
    If TypeName(parsed) <> "Dictionary" Then GoTo InvalidResponse
    Set ParseGdtAuthObject = parsed
    Exit Function
InvalidResponse:
    Err.Raise vbObjectError + 2101, "ParseGdtAuthObject", "Invalid authentication JSON object"
End Function

Public Function GdtRequiredAuthText(ByVal parsed As Object, ByVal key As String) As String
    On Error GoTo InvalidField
    If Not parsed.Exists(key) Then GoTo InvalidField
    If VarType(parsed(key)) <> vbString Then GoTo InvalidField
    If Len(Trim$(CStr(parsed(key)))) = 0 Then GoTo InvalidField
    GdtRequiredAuthText = CStr(parsed(key))
    Exit Function
InvalidField:
    Err.Raise vbObjectError + 2102, "GdtRequiredAuthText", "Missing or invalid authentication field: " & key
End Function
