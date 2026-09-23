VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGdtRequestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Success As Boolean
Public StatusCode As Long
Public ResponseText As String
Public ResponseBody As Variant
Public ErrorMessage As String
Public Attempts As Long
Public RetryAfterSeconds As Long
Public ShouldQueueFinalRetry As Boolean
Public AuthenticationFailure As Boolean
Public NoXml As Boolean

