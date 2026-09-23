[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\App_v6.4-ui.xlsm')
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking

$templatePath = Join-Path $RepositoryRoot 'template\App_Template.xlsm'
$sourceForm = Join-Path $RepositoryRoot 'src\forms\frmTaiHoaDon.frm'
$sourceFrx = Join-Path $RepositoryRoot 'src\forms\frmTaiHoaDon.frx'
$formsDir = Join-Path $RepositoryRoot 'src\forms'
$stageDir = Join-Path $env:TEMP ('hddt-form-ui-' + [Guid]::NewGuid().ToString('N'))
$stageForm = Join-Path $stageDir 'frmTaiHoaDon.frm'
$stageFrx = Join-Path $stageDir 'frmTaiHoaDon.frx'

New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
$utf8Text = Get-Content -LiteralPath $sourceForm -Raw -Encoding utf8
[IO.File]::WriteAllText($stageForm, $utf8Text, [Text.Encoding]::GetEncoding(1258))
Copy-Item -LiteralPath $sourceFrx -Destination $stageFrx -Force
Copy-Item -LiteralPath $templatePath -Destination $OutputPath -Force

$excel = $null
$workbook = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($OutputPath, 0, $false)
    Write-Host 'Opened workbook.'
    if ([int]$workbook.VBProject.Protection -ne 0) { throw 'VBA project is locked.' }
    if ([bool]$workbook.VBASigned) { throw 'Workbook has a VBA digital signature; editing would invalidate it.' }

    $mainForm = $workbook.VBProject.VBComponents.Item('frmTaiHoaDon')
    $codeStart = $utf8Text.IndexOf('Option Explicit')
    if ($codeStart -lt 0) { throw 'Option Explicit not found in frmTaiHoaDon source.' }
    $mainCode = $utf8Text.Substring($codeStart)
    $codeModule = $mainForm.CodeModule
    if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
    $codeModule.AddFromString($mainCode)
    Release-ComObject $codeModule
    Write-Host 'Updated frmTaiHoaDon code.'
    $mainForm.Properties.Item('Height').Value = 550
    Write-Host 'Configured runtime progress and date-picker controls.'

    try { $oldHandler = $workbook.VBProject.VBComponents.Item('clsUiButtonHandler'); $workbook.VBProject.VBComponents.Remove($oldHandler); Release-ComObject $oldHandler } catch {}
    $handlerComponent = $workbook.VBProject.VBComponents.Add(2)
    $handlerComponent.Name = 'clsUiButtonHandler'
    $handlerComponent.CodeModule.AddFromString(@'
Option Explicit
Public WithEvents Button As MSForms.CommandButton
Public Owner As Object
Public ActionName As String
Private Sub Button_Click()
    CallByName Owner, ActionName, VbMethod
End Sub
'@)

    try {
        $oldPicker = $workbook.VBProject.VBComponents.Item('frmDatePicker')
        $workbook.VBProject.VBComponents.Remove($oldPicker)
        Release-ComObject $oldPicker
    } catch {}
    $picker = $workbook.VBProject.VBComponents.Add(3)
    $picker.Name = 'frmDatePicker'
    $picker.Properties.Item('Caption').Value = 'Chọn ngày'
    $picker.Properties.Item('Width').Value = 270
    $picker.Properties.Item('Height').Value = 150
    $pickerCode = @'
Option Explicit

Private mTarget As Object
Private mHandlers As Collection

Private Sub UserForm_Initialize()
    BuildPickerControls
    SetPickerDate Date
End Sub

Private Sub BuildPickerControls()
    Dim ctl As Object, i As Long, h As clsUiButtonHandler
    Set mHandlers = New Collection
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblDay", True): ctl.Caption = "Ngày": ctl.Left = 18: ctl.Top = 12: ctl.Width = 48: ctl.Height = 18
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblMonth", True): ctl.Caption = "Tháng": ctl.Left = 78: ctl.Top = 12: ctl.Width = 72: ctl.Height = 18
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblYear", True): ctl.Caption = "Năm": ctl.Left = 168: ctl.Top = 12: ctl.Width = 72: ctl.Height = 18
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", "cboDay", True): ctl.Left = 18: ctl.Top = 33: ctl.Width = 48: ctl.Height = 22
    For i = 1 To 31: ctl.AddItem Format$(i, "00"): Next i
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", "cboMonth", True): ctl.Left = 78: ctl.Top = 33: ctl.Width = 72: ctl.Height = 22
    For i = 1 To 12: ctl.AddItem Format$(i, "00"): Next i
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", "cboYear", True): ctl.Left = 168: ctl.Top = 33: ctl.Width = 72: ctl.Height = 22
    For i = Year(Date) - 10 To Year(Date) + 10: ctl.AddItem CStr(i): Next i
    AddPickerButton "cmdToday", "Hôm nay", 18, "SelectToday"
    AddPickerButton "cmdOK", "Chọn", 102, "AcceptDate"
    AddPickerButton "cmdCancel", "Hủy", 180, "CancelPicker"
End Sub

Private Sub AddPickerButton(ByVal controlName As String, ByVal caption As String, ByVal leftPos As Single, ByVal action As String)
    Dim ctl As Object, h As clsUiButtonHandler
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    ctl.Caption = caption: ctl.Left = leftPos: ctl.Top = 75: ctl.Width = 66: ctl.Height = 27
    Set h = New clsUiButtonHandler: Set h.Button = ctl: Set h.Owner = Me: h.ActionName = action: mHandlers.Add h
End Sub

Public Sub SetTarget(ByVal targetControl As Object)
    Set mTarget = targetControl
    If IsDate(targetControl.Value) Then SetPickerDate CDate(targetControl.Value)
End Sub

Private Sub SetPickerDate(ByVal selectedDate As Date)
    Me.Controls("cboDay").Value = Format$(Day(selectedDate), "00")
    Me.Controls("cboMonth").Value = Format$(Month(selectedDate), "00")
    Me.Controls("cboYear").Value = CStr(Year(selectedDate))
End Sub

Public Sub SelectToday()
    SetPickerDate Date
End Sub

Public Sub AcceptDate()
    Dim selectedDate As Date
    On Error GoTo InvalidDate
    selectedDate = DateSerial(CLng(Me.Controls("cboYear").Value), CLng(Me.Controls("cboMonth").Value), CLng(Me.Controls("cboDay").Value))
    If Day(selectedDate) <> CLng(Me.Controls("cboDay").Value) Or Month(selectedDate) <> CLng(Me.Controls("cboMonth").Value) Then GoTo InvalidDate
    mTarget.Value = Format$(selectedDate, "dd/mm/yyyy")
    frmTaiHoaDon.RefreshPickedDate mTarget.Name
    Unload Me
    Exit Sub
InvalidDate:
    MsgBox "Ngày đã chọn không hợp lệ.", vbExclamation, "Chọn ngày"
End Sub

Public Sub CancelPicker()
    Unload Me
End Sub
'@
    $picker.CodeModule.AddFromString($pickerCode)
    Write-Host 'Created frmDatePicker.'

    $workbook.SaveAs($OutputPath, 52)
    Write-Host 'Saved workbook.'
    $workbook.Close($true)
    Release-ComObject $picker; Release-ComObject $handlerComponent; Release-ComObject $mainForm
    Release-ComObject $workbook; $workbook=$null
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $workbook; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Reopen, export the two edited forms back to source control, and prove the file reopens cleanly.
$excel = $null; $workbook = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    $workbook=$excel.Workbooks.Open($OutputPath,0,$true)
    Write-Host 'Reopened workbook for export.'
    foreach($name in @('frmTaiHoaDon','frmDatePicker')) {
        $component=$workbook.VBProject.VBComponents.Item($name)
        $exportPath=Join-Path $stageDir ($name+'.frm')
        $component.Export($exportPath)
        Copy-Item -LiteralPath $exportPath -Destination (Join-Path $formsDir ($name+'.frm')) -Force
        $frx=[IO.Path]::ChangeExtension($exportPath,'.frx')
        if(Test-Path -LiteralPath $frx){Copy-Item -LiteralPath $frx -Destination (Join-Path $formsDir ($name+'.frx')) -Force}
        Release-ComObject $component
    }
    $workbook.Close($false); Release-ComObject $workbook; $workbook=$null
} finally {
    if($null -ne $workbook){try{$workbook.Close($false)}catch{}}
    if($null -ne $excel){try{$excel.Quit()}catch{}}
    Release-ComObject $workbook; Release-ComObject $excel
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

Remove-Item -LiteralPath $stageDir -Recurse -Force
Write-Host "UI upgrade created and reopened successfully: $OutputPath"
