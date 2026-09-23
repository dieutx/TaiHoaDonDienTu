Attribute VB_Name = "modGdtRetry"
Option Explicit

' Central retry configuration. These values are intentionally kept in one module.
Public Const GDT_MAX_RETRIES As Long = 5
Public Const GDT_RETRY_BASE_SECONDS As Long = 2
Public Const GDT_RETRY_MAX_SECONDS As Long = 60
Public Const GDT_FINAL_RETRY_ENABLED As Boolean = True
Public Const GDT_FINAL_RETRY_COOLDOWN_SECONDS As Long = 30
Public Const GDT_REQUEST_DELAY_MS As Long = 600
Public Const GDT_HTTP_TIMEOUT_SECONDS As Long = 90

Private Const GDT_USER_AGENT As String = "Mozilla/5.0 ExcelVBA-HDDT"
Private mRandomized As Boolean
Public GdtAuthenticationFailed As Boolean

Public Sub ResetGdtRequestSession()
    GdtAuthenticationFailed = False
    ResetGdtStatusBar
End Sub

Public Function ExecuteGdtRequest( _
    ByVal httpMethod As String, _
    ByVal requestUrl As String, _
    ByVal bearerToken As String, _
    Optional ByVal requestBody As String = vbNullString, _
    Optional ByVal contentType As String = "application/json", _
    Optional ByVal acceptType As String = "application/json, text/plain, */*", _
    Optional ByVal expectBinary As Boolean = False, _
    Optional ByVal maxAttempts As Long = GDT_MAX_RETRIES, _
    Optional ByVal statusPrefix As String = vbNullString, _
    Optional ByVal actionHeader As String = vbNullString) As clsGdtRequestResult

    Dim result As clsGdtRequestResult
    Dim http As Object
    Dim attempt As Long
    Dim waitSeconds As Long
    Dim transportError As Boolean
    Dim transportMessage As String
    Dim responseReadError As String
    Dim responseTextValue As String
    Dim responseBodyValue As Variant
    Dim readSucceeded As Boolean

    Set result = New clsGdtRequestResult
    If maxAttempts < 1 Then maxAttempts = 1

    For attempt = 1 To maxAttempts
        Set http = Nothing
        transportError = False
        transportMessage = vbNullString
        result.StatusCode = 0
        result.ResponseText = vbNullString
        result.ResponseBody = Empty
        result.RetryAfterSeconds = 0

        On Error GoTo RequestError
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts GDT_HTTP_TIMEOUT_SECONDS * 1000, _
                         GDT_HTTP_TIMEOUT_SECONDS * 1000, _
                         GDT_HTTP_TIMEOUT_SECONDS * 1000, _
                         GDT_HTTP_TIMEOUT_SECONDS * 1000
        http.Open UCase$(httpMethod), requestUrl, False
        http.setRequestHeader "User-Agent", GDT_USER_AGENT
        http.setRequestHeader "Accept", acceptType
        http.setRequestHeader "Request-Id", CreateRequestID
        If Len(contentType) > 0 Then http.setRequestHeader "Content-Type", contentType
        If Len(bearerToken) > 0 Then http.setRequestHeader "Authorization", "Bearer " & bearerToken
        If Len(actionHeader) > 0 Then
            http.setRequestHeader "Accept-Language", "vi"
            http.setRequestHeader "Action", actionHeader
            http.setRequestHeader "End-Point", "/tra-cuu/tra-cuu-hoa-don"
            http.setRequestHeader "Referer", "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don"
        End If

        If Len(requestBody) > 0 Then
            http.send requestBody
        Else
            http.send
        End If

        result.StatusCode = CLng(http.Status)
        result.Attempts = attempt
        readSucceeded = TryReadGdtResponse(http, _
                                           expectBinary And IsSuccessfulGdtStatus(result.StatusCode), _
                                           responseTextValue, responseBodyValue, responseReadError)
        result.ResponseText = responseTextValue
        result.ResponseBody = responseBodyValue
        If Not readSucceeded And IsSuccessfulGdtStatus(result.StatusCode) Then
            transportError = True
            result.ErrorMessage = UniConvert("Khoong ddojc dduwowjc pharn hoofi tuwf masy chur: ") & responseReadError
            If attempt < maxAttempts Then
                waitSeconds = CalculateGdtRetrySeconds(attempt, 0)
                UpdateGdtRetryStatus statusPrefix, attempt + 1, maxAttempts, 0, waitSeconds
                WaitGdtSeconds waitSeconds
            End If
            GoTo ContinueAttempt
        End If

        If IsSuccessfulGdtStatus(result.StatusCode) Then
            result.Success = True
            result.ErrorMessage = vbNullString
            result.ShouldQueueFinalRetry = False
            Set ExecuteGdtRequest = result
            Exit Function
        End If

        If result.StatusCode = 401 Or result.StatusCode = 403 Then
            result.AuthenticationFailure = True
            result.ErrorMessage = UniConvert("Token heest hajn hoawjc khoong cos quyeefn truy caajp.")
            Set ExecuteGdtRequest = result
            Exit Function
        End If

        If IsNoXmlResponse(requestUrl, result.StatusCode, result.ResponseText) Then
            result.NoXml = True
            result.ErrorMessage = UniConvert("Khoong cos XML.")
            Set ExecuteGdtRequest = result
            Exit Function
        End If

        result.ErrorMessage = BuildHttpErrorMessage(result.StatusCode, result.ResponseText)
        If Not IsRetryableGdtStatus(result.StatusCode) Then
            Set ExecuteGdtRequest = result
            Exit Function
        End If

        result.RetryAfterSeconds = ReadRetryAfterSeconds(http)
        If attempt < maxAttempts Then
            waitSeconds = CalculateGdtRetrySeconds(attempt, result.RetryAfterSeconds)
            UpdateGdtRetryStatus statusPrefix, attempt + 1, maxAttempts, result.StatusCode, waitSeconds
            WaitGdtSeconds waitSeconds
        End If

ContinueAttempt:
        On Error GoTo 0
        Set http = Nothing
    Next attempt

    result.ShouldQueueFinalRetry = ShouldQueueGdtFinalRetry( _
                                       result.StatusCode, transportError, result.Attempts, maxAttempts)
    Set ExecuteGdtRequest = result
    Exit Function

RequestError:
    transportError = True
    transportMessage = Err.Description
    Err.Clear
    On Error GoTo 0
    result.StatusCode = 0
    result.Attempts = attempt
    result.ErrorMessage = UniConvert("Looxi keest noosi hoawjc heest thowfi gian: ") & transportMessage
    If attempt < maxAttempts Then
        waitSeconds = CalculateGdtRetrySeconds(attempt, 0)
        UpdateGdtRetryStatus statusPrefix, attempt + 1, maxAttempts, 0, waitSeconds
        WaitGdtSeconds waitSeconds
        GoTo ContinueAttempt
    End If
    GoTo ContinueAttempt
End Function

Public Function ExecuteGdtBinaryRequest( _
    ByVal requestUrl As String, _
    ByVal bearerToken As String, _
    Optional ByVal maxAttempts As Long = GDT_MAX_RETRIES, _
    Optional ByVal statusPrefix As String = vbNullString) As clsGdtRequestResult

    Set ExecuteGdtBinaryRequest = ExecuteGdtRequest("GET", requestUrl, bearerToken, _
        vbNullString, "application/zip", "application/zip, application/octet-stream", _
        True, maxAttempts, statusPrefix)
End Function

Public Function IsSuccessfulGdtStatus(ByVal statusCode As Long) As Boolean
    IsSuccessfulGdtStatus = (statusCode >= 200 And statusCode < 300)
End Function

Public Function ShouldQueueGdtFinalRetry( _
    ByVal statusCode As Long, _
    ByVal transportError As Boolean, _
    ByVal attempts As Long, _
    ByVal maxAttempts As Long) As Boolean

    If Not GDT_FINAL_RETRY_ENABLED Then Exit Function
    If maxAttempts < 1 Then maxAttempts = 1
    If attempts < maxAttempts Then Exit Function
    ShouldQueueGdtFinalRetry = transportError Or IsRetryableGdtStatus(statusCode)
End Function

Public Function TryReadGdtResponse( _
    ByVal http As Object, _
    ByVal requireBinaryBody As Boolean, _
    ByRef responseText As String, _
    ByRef responseBody As Variant, _
    ByRef errorMessage As String) As Boolean

    responseText = vbNullString
    responseBody = Empty
    errorMessage = vbNullString

    On Error GoTo TextReadFailed
    responseText = CStr(http.responseText)

ReadBinary:
    On Error GoTo BinaryReadFailed
    If requireBinaryBody Then responseBody = http.responseBody
    On Error GoTo 0
    TryReadGdtResponse = True
    Exit Function

TextReadFailed:
    If Not requireBinaryBody Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    GoTo ReadBinary

BinaryReadFailed:
    errorMessage = Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Public Function IsRetryableGdtStatus(ByVal statusCode As Long) As Boolean
    Select Case statusCode
        Case 0, 429, 500, 502, 503, 504
            IsRetryableGdtStatus = True
    End Select
End Function

Public Function CalculateGdtRetrySeconds( _
    ByVal failedAttempt As Long, _
    Optional ByVal retryAfterSeconds As Long = 0, _
    Optional ByVal jitterSeconds As Double = -1) As Long

    Dim baseDelay As Double
    Dim computedDelay As Double

    If retryAfterSeconds > 0 Then
        baseDelay = retryAfterSeconds
    Else
        If failedAttempt < 1 Then failedAttempt = 1
        baseDelay = GDT_RETRY_BASE_SECONDS * (2 ^ (failedAttempt - 1))
    End If

    If jitterSeconds < 0 Then
        If Not mRandomized Then
            Randomize
            mRandomized = True
        End If
        jitterSeconds = Rnd
    End If

    If jitterSeconds < 0 Then jitterSeconds = 0
    If jitterSeconds > 1 Then jitterSeconds = 1
    computedDelay = baseDelay + jitterSeconds
    If computedDelay > GDT_RETRY_MAX_SECONDS Then computedDelay = GDT_RETRY_MAX_SECONDS
    CalculateGdtRetrySeconds = CLng(Int(computedDelay + 0.999999))
End Function

Public Sub WaitGdtSeconds(ByVal seconds As Long)
    Dim deadline As Date

    If seconds <= 0 Then Exit Sub
    If seconds > GDT_RETRY_MAX_SECONDS Then seconds = GDT_RETRY_MAX_SECONDS
    deadline = DateAdd("s", seconds, Now)
    Do While Now < deadline
        DoEvents
    Loop
End Sub

Public Sub WaitGdtMilliseconds(ByVal milliseconds As Long)
    Dim started As Double
    Dim elapsed As Double

    If milliseconds <= 0 Then Exit Sub
    started = Timer
    Do
        DoEvents
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
    Loop While elapsed * 1000# < milliseconds
End Sub

Public Sub WaitForFinalGdtRetry()
    If GDT_FINAL_RETRY_ENABLED Then WaitGdtSeconds GDT_FINAL_RETRY_COOLDOWN_SECONDS
End Sub

Public Sub ResetGdtStatusBar()
    Application.StatusBar = False
End Sub

Private Function ReadRetryAfterSeconds(ByVal http As Object) As Long
    Dim rawValue As String

    On Error GoTo MissingHeader
    rawValue = Trim$(CStr(http.getResponseHeader("Retry-After")))
    On Error GoTo 0
    If Len(rawValue) = 0 Then Exit Function
    If IsNumeric(rawValue) Then
        If CDbl(rawValue) > 0 Then ReadRetryAfterSeconds = CLng(CDbl(rawValue))
    End If
    Exit Function

MissingHeader:
    Err.Clear
    On Error GoTo 0
End Function

Private Function IsNoXmlResponse( _
    ByVal requestUrl As String, _
    ByVal statusCode As Long, _
    ByVal responseText As String) As Boolean

    Dim normalized As String

    If statusCode <> 404 And statusCode <> 500 Then Exit Function
    If InStr(1, requestUrl, "export-xml", vbTextCompare) = 0 Then Exit Function
    normalized = LCase$(responseText)
    IsNoXmlResponse = (InStr(1, normalized, "khong co", vbTextCompare) > 0) Or _
                      (InStr(1, normalized, LCase$(UniConvert("khoong cos")), vbTextCompare) > 0) Or _
                      (InStr(1, normalized, "no xml", vbTextCompare) > 0) Or _
                      (InStr(1, normalized, "not found", vbTextCompare) > 0)
End Function

Private Function BuildHttpErrorMessage(ByVal statusCode As Long, ByVal responseText As String) As String
    Dim safeText As String

    safeText = Trim$(responseText)
    If Len(safeText) > 500 Then safeText = Left$(safeText, 500)
    If Len(safeText) > 0 Then
        BuildHttpErrorMessage = "HTTP " & CStr(statusCode) & ": " & safeText
    Else
        BuildHttpErrorMessage = "HTTP " & CStr(statusCode)
    End If
End Function

Private Sub UpdateGdtRetryStatus( _
    ByVal statusPrefix As String, _
    ByVal nextAttempt As Long, _
    ByVal maxAttempts As Long, _
    ByVal statusCode As Long, _
    ByVal waitSeconds As Long)

    Dim message As String

    If Len(statusPrefix) > 0 Then message = statusPrefix & " - "
    message = message & "HTTP " & CStr(statusCode) & " - " & _
              UniConvert("thuwr laji ") & CStr(nextAttempt) & "/" & CStr(maxAttempts) & _
              UniConvert(" sau ") & CStr(waitSeconds) & UniConvert(" giaay")
    Application.StatusBar = message
    On Error Resume Next
    frmTaiHoaDon.ReportRetryProgress message
    On Error GoTo 0
End Sub
