VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsGdtXmlTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public InvoiceIndex As Long
Public Endpoint As String
Public ResponseKind As String
Public Result As clsGdtRequestResult
Public Running As Boolean
Public Finished As Boolean
Public ElapsedSeconds As Double
Private mHttp As Object
Private mToken As String
Private mDue As Double
Private mStarted As Double
Private mFirstStarted As Double
Private mMaxAttempts As Long
Private mTimeoutSeconds As Long
Private mExpectBinary As Boolean
Private mActionHeader As String

Public Sub Configure(ByVal index As Long, ByVal requestUrl As String, ByVal token As String, _
    Optional ByVal maxAttempts As Long = GDT_MAX_RETRIES, _
    Optional ByVal timeoutSeconds As Long = GDT_HTTP_TIMEOUT_SECONDS, _
    Optional ByVal expectBinary As Boolean = True, _
    Optional ByVal actionHeader As String = vbNullString)
    InvoiceIndex = index
    Endpoint = requestUrl
    mToken = token
    mExpectBinary = expectBinary
    mActionHeader = actionHeader
    mMaxAttempts = maxAttempts
    If mMaxAttempts < 1 Then mMaxAttempts = 1
    mTimeoutSeconds = timeoutSeconds
    If mTimeoutSeconds < 1 Then mTimeoutSeconds = 1
    Set Result = New clsGdtRequestResult
End Sub

Public Sub Tick(ByVal throttle As clsGdtXmlThrottle, ByVal activeCount As Long)
    Dim atTime As Double, readError As String, readOk As Boolean
    Dim responseTextValue As String, responseBodyValue As Variant
    atTime = GdtClockSeconds()
    If Finished Then Exit Sub
    If GdtStopRequested Or GdtAuthenticationFailed Then
        Cancel
        Exit Sub
    End If
    On Error GoTo TransportFailed
    If Not Running Then
        If atTime < mDue Then Exit Sub
        If Not throttle.CanStart(activeCount, atTime) Then Exit Sub
        Result.Attempts = Result.Attempts + 1
        Set mHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        mHttp.setTimeouts mTimeoutSeconds * 1000, mTimeoutSeconds * 1000, _
            mTimeoutSeconds * 1000, mTimeoutSeconds * 1000
        Result.StatusCode = 0
        Result.ResponseBody = Empty
        Result.ResponseText = vbNullString
        Result.RetryAfterSeconds = 0
        Result.NotBefore = 0
        Result.InvalidJsonBody = False
        mStarted = atTime
        If mFirstStarted = 0 Then mFirstStarted = atTime
        throttle.Started atTime
        mHttp.Open "GET", Endpoint, True
        mHttp.setRequestHeader "User-Agent", "Mozilla/5.0 ExcelVBA-HDDT"
        If mExpectBinary Then
            mHttp.setRequestHeader "Accept", "application/zip, application/octet-stream"
            mHttp.setRequestHeader "Content-Type", "application/zip"
        Else
            mHttp.setRequestHeader "Accept", "application/json, text/plain, */*"
            mHttp.setRequestHeader "Content-Type", "application/json"
        End If
        mHttp.setRequestHeader "Request-Id", CreateRequestID
        If Len(mToken) > 0 Then mHttp.setRequestHeader "Authorization", "Bearer " & mToken
        If Len(mActionHeader) > 0 Then
            mHttp.setRequestHeader "Accept-Language", "vi"
            mHttp.setRequestHeader "Action", mActionHeader
            mHttp.setRequestHeader "End-Point", "/tra-cuu/tra-cuu-hoa-don"
            mHttp.setRequestHeader "Referer", "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don"
        End If
        mHttp.send
        Running = True
        Exit Sub
    End If
    If mHttp.readyState <> 4 Then
        If atTime - mStarted > mTimeoutSeconds Then
            mHttp.abort
            Err.Raise vbObjectError + 711, , IIf(mExpectBinary, "XML", "JSON") & " request timed out"
        End If
        Exit Sub
    End If
    Result.StatusCode = CLng(mHttp.Status)
    readOk = TryReadGdtResponse(mHttp, mExpectBinary And IsSuccessfulGdtStatus(Result.StatusCode), _
        responseTextValue, responseBodyValue, readError)
    Result.ResponseText = responseTextValue
    Result.ResponseBody = responseBodyValue
    If IsSuccessfulGdtStatus(Result.StatusCode) And readOk Then
        If Not mExpectBinary And IsGdtMarkupResponse(Result.ResponseText) Then
            ' HTTP 200 with a web page body: the JSON API never answered, so this
            ' is a transport failure and the slow down policy applies.
            Result.InvalidJsonBody = True
            Result.ErrorMessage = BuildGdtMarkupResponseMessage(Result.StatusCode, _
                Result.ResponseText, ReadGdtResponseContentType(mHttp))
            throttle.RateLimited CalculateGdtRetrySeconds(Result.Attempts, Result.RetryAfterSeconds, 0), atTime
            RetryOrComplete atTime
        Else
            Result.Success = True
            Result.ErrorMessage = vbNullString
            Result.ShouldQueueFinalRetry = False
            Complete
        End If
    ElseIf Result.StatusCode = 401 Or Result.StatusCode = 403 Then
        GdtAuthenticationFailed = True
        Result.AuthenticationFailure = True
        Result.ErrorMessage = "HTTP " & Result.StatusCode & ": login required"
        Complete
    ElseIf mExpectBinary And IsNoXmlResponse(Endpoint, Result.StatusCode, Result.ResponseText) Then
        Result.NoXml = True
        Result.ErrorMessage = "No XML"
        Complete
    Else
        Result.RetryAfterSeconds = ReadRetryAfterSeconds(mHttp)
        If Result.RetryAfterSeconds > 0 Then Result.NotBefore = GdtClockSeconds() + Result.RetryAfterSeconds
        Result.ErrorMessage = BuildHttpErrorMessage(Result.StatusCode, Result.ResponseText)
        If Not readOk Then Result.StatusCode = 0: Result.ErrorMessage = readError
        If Result.StatusCode = 429 Then throttle.RateLimited Result.RetryAfterSeconds, atTime
        If Not mExpectBinary And (Result.StatusCode = 503 Or Result.StatusCode = 504) Then
            throttle.RateLimited CalculateGdtRetrySeconds(Result.Attempts, Result.RetryAfterSeconds, 0), atTime
        End If
        RetryOrComplete atTime
    End If
    Exit Sub

TransportFailed:
    Result.StatusCode = 0
    Result.ErrorMessage = Err.Description
    Err.Clear
    RetryOrComplete atTime
End Sub

Private Sub RetryOrComplete(ByVal atTime As Double)
    Dim retryable As Boolean

    Running = False
    Set mHttp = Nothing
    retryable = IsRetryableGdtStatus(Result.StatusCode) Or Result.InvalidJsonBody
    If retryable And Result.Attempts < mMaxAttempts Then
        mDue = atTime + CalculateGdtRetrySeconds(Result.Attempts, Result.RetryAfterSeconds)
    Else
        Result.ShouldQueueFinalRetry = retryable
        Complete
    End If
End Sub

Private Sub Complete()
    If mFirstStarted > 0 Then ElapsedSeconds = GdtClockSeconds() - mFirstStarted
    Running = False
    Finished = True
    Set mHttp = Nothing
    mToken = vbNullString
End Sub

Public Sub Cancel()
    On Error Resume Next
    If Not mHttp Is Nothing Then mHttp.abort
    Result.AuthenticationFailure = GdtAuthenticationFailed
    Result.ShouldQueueFinalRetry = False
    Result.ErrorMessage = IIf(mExpectBinary, "XML", "JSON") & " request cancelled"
    Complete
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mHttp Is Nothing Then mHttp.abort
End Sub
