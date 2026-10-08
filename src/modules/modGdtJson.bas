Attribute VB_Name = "modGdtJson"
Option Explicit

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
