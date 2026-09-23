Attribute VB_Name = "modHTTPRequest"
Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwFlags As LongPtr, ByVal dwReserved As Long) As Boolean
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwFlags As Long, ByVal dwReserved As Long) As Boolean
#End If

Public xmlHttp As Object
Public cookieHeader As String, errMsg As String
Public LastGdtStatus As Long
Public LastGdtAttempts As Long
Public LastGdtRetryAfter As Long
Public LastGdtError As String
Public LastGdtAuthFailed As Boolean
Public LastGdtRetryable As Boolean
Public LastGdtShouldQueue As Boolean

Function createXmlHttpObj()
    Set xmlHttp = CreateObject("MSXML2.serverXMLHTTP.6.0")
End Function

Function httpPost(ByVal url As String, ByVal formData As String) As String
    On Error GoTo EH
    
    With xmlHttp 'CreateObject("MSXML2.serverXMLHTTP.6.0")
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        '.setRequestHeader "Content-type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Content-type", "application/json"
        .setRequestHeader "Request-Id", CreateRequestID
        .send (formData)
        httpPost = .responseText
        'Luu cookie header (session) de su dung tiep
        cookieHeader = .getResponseHeader("Set-Cookie")
        'Debug.Print cookieHeader
    End With
    
EH_Exit:
    Exit Function
    
EH:
    Select Case err.Number
        Case -2147012889
            MsgBox "Loi ket noi Server hoac sai dia chi Web", vbExclamation, "Thông báo - Http Post error"
        Case -2147012894
            MsgBoxUni "Qu" & ChrW(225) & " th" & ChrW(7901) & "i gian k" & ChrW(7871) & "t n" & ChrW(7889) & "i v" & ChrW(7899) & "i m" & ChrW(225) & "y ch" & ChrW(7911) & " Web." & vbCrLf _
                & "Ki" & ChrW(7875) & "m tra " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng truy" & ChrW(7873) & "n Internet v" & ChrW(224) & " th" & ChrW(7917) & " k" & ChrW(7871) & "t n" & ChrW(7889) & "i l" & ChrW(7841) & "i.", vbExclamation, "Thông báo - Http Post error"
        Case Else
            MsgBox "ErrNum: " & err.Number & vbCrLf & "Err description: " & err.Description, vbExclamation, "Thông báo"
    End Select
    httpPost = "Error"
    Resume EH_Exit
    
End Function

Function httpGet(ByVal url As String, bearer As String, ByRef res As String) As Boolean
    On Error GoTo EH
    
    'With xmlHttp 'CreateObject("MSXML2.serverXMLHTTP.6.0")
    With CreateObject("MSXML2.serverXMLHTTP.6.0")
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
        .setRequestHeader "Accept", "application/json, text/plain, */*"
        .setRequestHeader "Content-type", "application/gzip;application/json; application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Request-Id", CreateRequestID
        .setRequestHeader "Authorization", "Bearer " & bearer & ""
        .send
        'httpGet = .responseText
        'On Error Resume Next
        res = .responseText
        
        'Dim getStatus As String
        If err.Number <> 0 Then
            'sometimes there is no "asnwer" from server
        Else
            getStatus = .Status
            If getStatus <> "300" And getStatus <> "200" Then
                Select Case getStatus
                    Case 429
                        'Debug.Print "Status: " & .Status & "-" & "Too many request sent to server"
                        MsgBoxUni "L" & ChrW(7895) & "i 429 - G" & ChrW(7917) & "i qu" & ChrW(225) & " nhi" & ChrW(7873) & "u y" & ChrW(234) & "u c" & ChrW(7847) & "u t" & ChrW(7899) & "i m" & ChrW(225) & "y ch" & ChrW(7911) & " Web.", vbExclamation, "Http Get Error 429"
                    Case 500
                        'Debug.Print "Status: " & .Status & "-" & "Internal server error"
                        MsgBoxUni "L" & ChrW(7895) & "i web server t" & ChrW(7915) & " ch" & ChrW(7889) & "i k" & ChrW(7871) & "t n" & ChrW(7889) & "i.", vbExclamation, "Http request error 500"
                    Case Else
                        'Debug.Print "Status: " & .Status & "-" & "Not working"
                End Select
                httpGet = False
            Else
                httpGet = True
                'Debug.Print "Status: " & .Status & "-" & "Connected."
            End If
        End If
        On Error GoTo 0
    End With
    
EH_Exit:
    Exit Function
    
EH:
    Select Case err.Number
        Case -2147012889
            MsgBoxUni "Loi ket noi Server hoac sai dia chi Web", vbExclamation, "Thông báo - Http get error"
        Case -2147012894
            MsgBoxUni "Qu" & ChrW(225) & " th" & ChrW(7901) & "i gian k" & ChrW(7871) & "t n" & ChrW(7889) & "i v" & ChrW(7899) & "i m" & ChrW(225) & "y ch" & ChrW(7911) & " Web." & vbCrLf _
                & "Ki" & ChrW(7875) & "m tra " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng truy" & ChrW(7873) & "n Internet v" & ChrW(224) & " th" & ChrW(7917) & " k" & ChrW(7871) & "t n" & ChrW(7889) & "i l" & ChrW(7841) & "i.", vbExclamation, "Thông báo - Http get error"
        Case Else
            MsgBox "ErrNum: " & err.Number & vbCrLf & "Err description: " & err.Description, vbExclamation, "Thông báo - Http Get error"
    End Select
    httpGet = False
    Resume EH_Exit
    
End Function

Function httpOption(url As String) As String
    On Error GoTo EH
    
    With xmlHttp 'CreateObject("MSXML2.serverXMLHTTP.6.0")
        .Open "OPTIONS", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        .setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        '.setRequestHeader "Content-type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Request-Id", CreateRequestID
        .send
        httpOption = .responseText
        
        If err.Number <> 0 Then
            'sometimes there is no "asnwer" from server
        Else
            getStatus = .Status
            If getStatus <> "300" And getStatus <> "200" Then
                Debug.Print "Not working"
            Else
                Debug.Print "Connected."
            End If
        End If
        On Error GoTo 0
    End With
    
EH_Exit:
    Exit Function
    
EH:
    Select Case err.Number
        Case -2147012889
            MsgBox "Loi ket noi Server hoac sai dia chi Web", vbExclamation, "Thông báo - Http Option error"
        Case Else
            MsgBox "ErrNum: " & err.Number & vbCrLf & "Err description: " & err.Description, vbExclamation, "Thông báo - Http Option error"
    End Select
    httpOption = "Error"
    Resume EH_Exit
    
End Function

Function httpGet_ex(ByVal url As String, bearer As String) As Boolean
    'On Error GoTo EH
    
    With xmlHttp 'CreateObject("MSXML2.serverXMLHTTP.6.0")
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        .setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        .setRequestHeader "Accept-Encoding", "gzip"
        .setRequestHeader "Content-type", "application/json; application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Request-Id", CreateRequestID
        .setRequestHeader "Authorization", "Bearer " & bearer & ""
        .send
        
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' Binary
        stream.Open
        stream.Write .responseBody ' Write the binary response to the stream
        
        ' Save the stream to the specified file path
        stream.SaveToFile "C:\Temp\invoices_XMLHTML\file_" & Format(Now, "yymmdd_hhnnss") & ".zip", 2 ' Overwrite if the file already exists
        stream.Close
        
        If err.Number <> 0 Then
            'sometimes there is no "asnwer" from server
        Else
            getStatus = .Status
            If getStatus <> "300" And getStatus <> "200" Then
                Debug.Print "Not working"
                If getStatus = 429 Then
                    Debug.Print "Too many request sent to server"
                End If
                httpGet_ex = False
            Else
                httpGet_ex = True
                Debug.Print "Connected."
            End If
        End If
        On Error GoTo 0
    End With
    
EH_Exit:
    Exit Function
    
EH:
    Select Case err.Number
        Case -2147012889
            MsgBox "Loi ket noi Server hoac sai dia chi Web", vbExclamation, "Thông báo"
        Case Else
            MsgBox "ErrNum: " & err.Number & vbCrLf & "Err description: " & err.Description, vbExclamation, "HttpGet_ex Error"
    End Select
    httpGet_ex = False
    Resume EH_Exit
    
End Function


Public Function validateInternetConnection() As Boolean
    Dim internetStatus As Boolean
    
    internetStatus = InternetGetConnectedState(0&, 0&)
    If internetStatus = False Then
        MsgBoxUni "Ki" & ChrW(7875) & "m tra l" & ChrW(7841) & "i k" & ChrW(7871) & "t n" & ChrW(7889) & "i internet", vbExclamation, "Thông báo"
        validateInternetConnection = False
        Exit Function
    End If
    validateInternetConnection = True
End Function

Public Function GetInternetConnectedState() As Boolean
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function

Public Function ApiGet(ByVal urlPath As String) As String
    Dim requestResult As clsGdtRequestResult

    If GdtAuthenticationFailed Then
        LastGdtStatus = 401
        LastGdtAttempts = 0
        LastGdtRetryAfter = 0
        LastGdtError = UniConvert("Token heest hajn. DDax duwfng yeeu caafu mowsi.")
        LastGdtAuthFailed = True
        LastGdtRetryable = False
        LastGdtShouldQueue = False
        Exit Function
    End If

    Set requestResult = ExecuteGdtRequest("GET", urlPath, getToken(), _
                                          vbNullString, "application/json", _
                                          "application/json, text/plain, */*", False)
    PublishLastGdtResult requestResult

    If requestResult.AuthenticationFailure Then GdtAuthenticationFailed = True
    If requestResult.Success Then ApiGet = requestResult.ResponseText

End Function

Public Sub PublishLastGdtResult(ByVal requestResult As clsGdtRequestResult)
    LastGdtStatus = requestResult.StatusCode
    LastGdtAttempts = requestResult.Attempts
    LastGdtRetryAfter = requestResult.RetryAfterSeconds
    LastGdtError = requestResult.ErrorMessage
    LastGdtAuthFailed = requestResult.AuthenticationFailure
    LastGdtRetryable = IsRetryableGdtStatus(requestResult.StatusCode)
    LastGdtShouldQueue = requestResult.ShouldQueueFinalRetry
    If requestResult.AuthenticationFailure Then GdtAuthenticationFailed = True
End Sub

