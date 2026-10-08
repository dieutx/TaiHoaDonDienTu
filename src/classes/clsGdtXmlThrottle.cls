VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsGdtXmlThrottle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BaseConcurrency As Long
Public CurrentConcurrency As Long
Public BaseIntervalMs As Long
Public CurrentIntervalMs As Long
Public CooldownUntil As Double
Public NextStart As Double
Public LastRateLimit As Double
Private mLastDrop As Double

Public Sub Configure(ByVal concurrency As Long, ByVal intervalMs As Long)
    If concurrency < 1 Then concurrency = 1
    If concurrency > GDT_XML_MAX_CONCURRENCY Then concurrency = GDT_XML_MAX_CONCURRENCY
    If intervalMs < 0 Then intervalMs = 0
    BaseConcurrency = concurrency
    CurrentConcurrency = concurrency
    BaseIntervalMs = intervalMs
    CurrentIntervalMs = intervalMs
    CooldownUntil = 0
    NextStart = 0
    LastRateLimit = 0
    mLastDrop = 0
End Sub

Public Function CanStart(ByVal activeCount As Long, ByVal atTime As Double) As Boolean
    Recover atTime
    CanStart = Not GdtStopRequested And Not GdtAuthenticationFailed And Not GdtPauseRequested _
        And activeCount < CurrentConcurrency And atTime >= CooldownUntil And atTime >= NextStart
End Function

Public Sub Started(ByVal atTime As Double)
    NextStart = atTime + CurrentIntervalMs / 1000#
End Sub

Public Sub RateLimited(ByVal retryAfter As Long, ByVal atTime As Double)
    Dim delay As Long
    delay = retryAfter
    If delay <= 0 Then delay = 15 + Int(Rnd * 5) - 2
    If CooldownUntil < atTime + delay Then CooldownUntil = atTime + delay
    If mLastDrop = 0 Or atTime - mLastDrop >= 30 Then
        If CurrentConcurrency > 1 Then CurrentConcurrency = CurrentConcurrency - 1
        mLastDrop = atTime
    End If
    CurrentIntervalMs = CLng(CurrentIntervalMs * 1.5)
    If CurrentIntervalMs < BaseIntervalMs Then CurrentIntervalMs = BaseIntervalMs
    If CurrentIntervalMs > 5000 Then CurrentIntervalMs = 5000
    LastRateLimit = atTime
End Sub

Public Sub Recover(ByVal atTime As Double)
    If LastRateLimit = 0 Or atTime < CooldownUntil Then Exit Sub
    If atTime - LastRateLimit < 10 Then Exit Sub
    If CurrentIntervalMs > BaseIntervalMs Then
        CurrentIntervalMs = BaseIntervalMs + (CurrentIntervalMs - BaseIntervalMs) \ 2
        If CurrentIntervalMs - BaseIntervalMs < 20 Then CurrentIntervalMs = BaseIntervalMs
    ElseIf CurrentConcurrency < BaseConcurrency Then
        CurrentConcurrency = CurrentConcurrency + 1
    End If
    LastRateLimit = atTime
    If CurrentIntervalMs = BaseIntervalMs And CurrentConcurrency = BaseConcurrency Then LastRateLimit = 0
End Sub
