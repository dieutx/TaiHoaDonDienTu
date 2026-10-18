[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$fixtureRoot = Join-Path $env:TEMP ('hddt-login-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$fixtureBook = Join-Path $fixtureRoot 'LoginFixture.xlsm'
Copy-Item -LiteralPath $BuiltWorkbook -Destination $fixtureBook
$excel = $null; $book = $null; $formCode = $null; $module = $null; $fakeHttp = $null
$results = [ordered]@{}
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($fixtureBook, 0, $false)
    # Instrument only the disposable workbook. The real button handler, JSON
    # parser, SVG recognizer and account/token persistence all execute in Excel.
    $formCode = $book.VBProject.VBComponents.Item('frmDangNhap').CodeModule
    $source = $formCode.Lines(1, $formCode.CountOfLines)
    $source = $source.Replace('CreateObject("MSXML2.ServerXMLHTTP.6.0")', 'NewLoginFixtureHttp()')
    $source = $source.Replace('http.Open ', 'http.FixtureOpen ')
    $source = $source.Replace('GetInternetConnectedState = False', 'FixtureOffline')
    $source = $source.Replace('MsgBoxUni ', 'CaptureLoginMessage ').Replace('MsgBox ', 'CaptureLoginMessage ')
    $source = [regex]::Replace($source, '(?m)^\s*Application\.OnTime .+$', '        FixtureScheduled = True')
    $formCode.DeleteLines(1, $formCode.CountOfLines)
    $formCode.AddFromString($source)
    $formCode.AddFromString(@'
Public Sub FixtureClick()
    cmdDangNhap_Click
End Sub
Public Sub FixturePrimeCaptcha()
    blnCaptcha = layCaptcha()
End Sub
'@)
    $fakeHttp = $book.VBProject.VBComponents.Add(2)
    $fakeHttp.Name = 'clsLoginFixtureHttp'
    $fakeHttp.CodeModule.AddFromString(@'
Option Explicit
Private verb As String
Public Status As Long
Public responseText As String
Public Sub setTimeouts(ByVal resolve As Long, ByVal connect As Long, ByVal sending As Long, ByVal receiving As Long)
    If resolve <= 0 Or connect <= 0 Or sending <= 0 Or receiving <= 0 Then Err.Raise 5, , "Missing timeout"
End Sub
Public Sub FixtureOpen(ByVal requestMethod As String, ByVal requestUrl As String, ByVal isAsync As Boolean)
    verb = requestMethod
    If isAsync Then Err.Raise 5, , "Unexpected async login"
    If verb = "GET" And InStr(requestUrl, "/captcha") = 0 Then Err.Raise 5, , "CAPTCHA endpoint changed"
    If verb = "POST" And InStr(requestUrl, "/security-taxpayer/authenticate") = 0 Then Err.Raise 5, , "Login endpoint changed"
End Sub
Public Sub setRequestHeader(ByVal name As String, ByVal value As String)
End Sub
Public Sub send(Optional ByVal body As String = "")
    If frmDangNhap.cmdDangNhap.Enabled Then Err.Raise 5, , "Button enabled during request"
    If verb = "GET" Then
        If frmDangNhap.cmdDangNhap.Caption <> UniConvert("DDang laays max...") Then Err.Raise 5, , "Missing CAPTCHA status"
        FixtureGets = FixtureGets + 1
        If FixtureThrowGet Then Err.Raise 5, , "Fixture network failure"
        Status = FixtureGetStatus: responseText = FixtureCaptcha
    Else
        If frmDangNhap.cmdDangNhap.Caption <> UniConvert("DDang ddawng nhaajp...") Then Err.Raise 5, , "Missing login status"
        FixturePosts = FixturePosts + 1: FixtureBody = body
        If FixtureThrowPost Then Err.Raise 5, , "Fixture timeout"
        Status = FixturePostStatus: responseText = FixtureAuth
    End If
End Sub
'@)
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'modLoginFixture'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public FixtureGets As Long, FixturePosts As Long, FixtureGetStatus As Long, FixturePostStatus As Long
Public FixtureOffline As Boolean, FixtureThrowGet As Boolean, FixtureThrowPost As Boolean, FixtureScheduled As Boolean
Public FixtureCaptcha As String, FixtureAuth As String, FixtureBody As String, FixtureMessage As String
Public Function NewLoginFixtureHttp() As Object
    Set NewLoginFixtureHttp = New clsLoginFixtureHttp
End Function
Public Sub CaptureLoginMessage(ByVal message As String, Optional ByVal buttons As Long, Optional ByVal title As String)
    FixtureMessage = message
End Sub
Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise 5, , message
End Sub
Public Function LoginFormSmoke() As String
    Dim form As Object
    On Error GoTo Failed
    Set form = VBA.UserForms.Add("frmDangNhap")
    Unload form
    LoginFormSmoke = "PASS"
    Exit Function
Failed:
    LoginFormSmoke = "FAIL: " & Err.Description
End Function
Private Function CaptchaJson(ByVal count As Long) As String
    Dim svg As String, i As Long, value As Object
    svg = "<svg>"
    For i = 1 To count
        svg = svg & "<path d=""M" & i & " 0 QQQQQZMQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZMQQZ""/>"
    Next
    Set value = CreateObject("Scripting.Dictionary")
    value.Add "key", "fixture-key": value.Add "content", svg & "</svg>"
    CaptchaJson = JsonConverter.ConvertToJson(value)
End Function
Public Function LoginFlowCase(ByVal scenario As String) As String
    Dim originalCaption As String, payload As Object
    On Error GoTo Failed
    Unload frmDangNhap
    With ThisWorkbook.Sheets("MENU")
        .Range("A7:E1000").ClearContents
        .Range("A7").Value = "FIXTURE-USER"
        .Range("B7").Value = "fixture-password"
        .Range("C7").Value = "Synthetic company"
    End With
    FixtureGets = 0: FixturePosts = 0: FixtureMessage = "": FixtureBody = ""
    FixtureOffline = False: FixtureThrowGet = False: FixtureThrowPost = False: FixtureScheduled = False
    FixtureGetStatus = 200: FixturePostStatus = 200
    FixtureCaptcha = CaptchaJson(6): FixtureAuth = "{""token"":""fixture-token""}"
    cToken = "": bearer = "": ckey = "stale-key"
    frmDangNhap.lsbUsers.ListIndex = 0
    originalCaption = frmDangNhap.cmdDangNhap.Caption
    Require FixtureGets = 0 And FixturePosts = 0, "Account selection issued a request"
    Select Case scenario
    Case "MissingUsername": frmDangNhap.txtUser.Value = "  "
    Case "MissingPassword": frmDangNhap.txtPass.Value = ""
    Case "Offline": FixtureOffline = True
    Case "CaptchaHttp": FixtureGetStatus = 503
    Case "CaptchaNetwork": FixtureThrowGet = True
    Case "CaptchaMalformed": FixtureCaptcha = "<html>fixture</html>"
    Case "CaptchaMissingKey": FixtureCaptcha = "{""content"":""fixture""}"
    Case "CaptchaEmpty": FixtureCaptcha = CaptchaJson(0)
    Case "CaptchaPartial": FixtureCaptcha = CaptchaJson(3)
    Case "LoginHttpRetry": FixturePostStatus = 401
    Case "LoginTimeoutRetry": FixtureThrowPost = True
    Case "LoginMalformed": FixtureAuth = "<html>fixture</html>"
    Case "LoginMessageRetry": FixtureAuth = "{""message"":""captcha fixture rejected""}"
    Case "LoginMissingToken": FixtureAuth = "{}"
    End Select
    frmDangNhap.FixtureClick
    Select Case scenario
    Case "SavedAccount", "EditedAccount"
        Require FixtureGets = 1 And FixturePosts = 1, "Saved account click did not fetch CAPTCHA and submit"
    Case "MissingUsername", "MissingPassword", "Offline"
        Require FixtureGets = 0 And FixturePosts = 0, "Invalid input/offline issued a request"
    Case "CaptchaHttp", "CaptchaNetwork", "CaptchaMalformed", "CaptchaMissingKey", "CaptchaEmpty", "CaptchaPartial"
        Require FixtureGets = 1 And FixturePosts = 0, "Invalid CAPTCHA submitted"
    Case Else
        Require FixtureGets = 1 And FixturePosts = 1, "Login request missing"
    End Select
    If scenario <> "SavedAccount" And scenario <> "EditedAccount" Then
        Require Len(FixtureMessage) > 0, "Failure was silent"
        Require cToken = "" And Not FixtureScheduled, "Failure persisted token or advanced form"
        Require ThisWorkbook.Sheets("MENU").Range("D7").Value = "", "Failure saved token"
    End If
    If Right$(scenario, 5) = "Retry" Then
        FixturePostStatus = 200: FixtureThrowPost = False: FixtureAuth = "{""token"":""fixture-token""}"
        FixtureMessage = ""
        frmDangNhap.FixtureClick
        Require FixtureGets = 2 And FixturePosts = 2, "Retry reused submitted CAPTCHA"
    End If
    If scenario = "EditedAccount" Then
        ' Prime a challenge, then edit the password. Editing must clear it.
        FixtureScheduled = False
        frmDangNhap.cmdDangNhap.Enabled = False
        frmDangNhap.cmdDangNhap.Caption = UniConvert("DDang laays max...")
        frmDangNhap.FixturePrimeCaptcha
        frmDangNhap.cmdDangNhap.Enabled = True
        frmDangNhap.cmdDangNhap.Caption = originalCaption
        frmDangNhap.txtPass.Value = "changed-fixture-password"
        Require ckey = "" And frmDangNhap.txtNhapCaptcha.Value = "", "Password edit kept stale CAPTCHA"
        frmDangNhap.FixtureClick
        Require FixtureGets = 3 And FixturePosts = 2, "Password edit did not fetch new challenge"
    End If
    If FixtureScheduled Then
        Require cToken = "fixture-token" And bearer = "fixture-token", "Token not installed"
        Require ThisWorkbook.Sheets("MENU").Range("D7").Value = "fixture-token", "Token not saved"
        Require IsDate(ThisWorkbook.Sheets("MENU").Range("E7").Value), "Token time not saved"
        Set payload = JsonConverter.ParseJSON(FixtureBody)
        Require payload("username") = "FIXTURE-USER" And payload("cvalue") = "AAAAAA" And payload("ckey") = "fixture-key", "Login payload mismatch"
        If scenario = "EditedAccount" Then Require payload("password") = "changed-fixture-password", "Password edit ignored"
    End If
    Require frmDangNhap.cmdDangNhap.Enabled And frmDangNhap.cmdDangNhap.Caption = originalCaption, "Button not restored"
    LoginFlowCase = "PASS"
    Unload frmDangNhap
    Exit Function
Failed:
    LoginFlowCase = "FAIL: " & Err.Description
    Unload frmDangNhap
End Function
'@)
    $results['FormInitialization'] = [string]$excel.Run("'$($book.Name)'!LoginFormSmoke")
    Write-Output ("FormInitialization: " + $results['FormInitialization'])
    if ($results['FormInitialization'] -ne 'PASS') { throw $results['FormInitialization'] }
    foreach ($scenario in @('SavedAccount', 'MissingUsername', 'MissingPassword', 'Offline', 'CaptchaHttp', 'CaptchaNetwork', 'CaptchaMalformed', 'CaptchaMissingKey', 'CaptchaEmpty', 'CaptchaPartial', 'LoginHttpRetry', 'LoginTimeoutRetry', 'LoginMessageRetry', 'LoginMalformed', 'LoginMissingToken', 'EditedAccount')) {
        $result = [string]$excel.Run("'$($book.Name)'!LoginFlowCase", $scenario)
        $results[$scenario] = $result
        Write-Output "${scenario}: $result"
        if ($result -ne 'PASS') { throw "${scenario}: $result" }
    }
} catch {
    Write-Output $_.ScriptStackTrace
    try {
        $pane = $excel.VBE.ActiveCodePane
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        Write-Output ("Compile location=" + $pane.CodeModule.Parent.Name + ':' + $startLine)
        Write-Output $pane.CodeModule.Lines($startLine, 1)
        Release-ComObject $pane
    } catch {}
    throw
} finally {
    $results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $BuiltWorkbook) 'login-flow-runtime-result.json') -Encoding UTF8
    Release-ComObject $formCode; Release-ComObject $module; Release-ComObject $fakeHttp
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved) -match '^hddt-login-[a-f0-9]{32}$') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
if ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -ne $before) { throw 'Test changed the input workbook.' }
Write-Output 'LOGIN FLOW RUNTIME TEST PASSED'
