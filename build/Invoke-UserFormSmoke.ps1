[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BuiltWorkbook,
    [Parameter(Mandatory)][string]$FormName
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$excel = $null
$workbook = $null
$testModule = $null
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $testModule = $workbook.VBProject.VBComponents.Add(1)
    $testModule.Name = 'modCodexBuildSmokeTest'
    $code = @(
        'Option Explicit',
        'Public Function CodexInstantiateForm(ByVal formName As String) As Boolean',
        'On Error GoTo Failed',
        'Dim frm As Object',
        'Set frm = VBA.UserForms.Add(formName)',
        'If formName = "frmTaiHoaDon" Then',
        '    If frm.Controls("cmdPauseResume").Enabled Then GoTo Failed',
        '    If frm.Controls("cmdStopDownload").Enabled Then GoTo Failed',
        '    If Not frm.Controls("cmdPauseResume").Visible Then GoTo Failed',
        '    If Not frm.Controls("cmdStopDownload").Visible Then GoTo Failed',
        '    If frm.Controls("cmdPauseResume").Left + frm.Controls("cmdPauseResume").Width > frm.InsideWidth Then GoTo Failed',
        '    If frm.Controls("cmdStopDownload").Left + frm.Controls("cmdStopDownload").Width > frm.InsideWidth Then GoTo Failed',
        'End If',
        'Unload frm',
        'CodexInstantiateForm = True',
        'Exit Function',
        'Failed:',
        'CodexInstantiateForm = False',
        'End Function'
    ) -join "`r`n"
    $testModule.CodeModule.AddFromString($code)
    $ok = [bool]$excel.Run("'$($workbook.Name)'!CodexInstantiateForm", $FormName)
    $workbook.VBProject.VBComponents.Remove($testModule)
    Release-ComObject $testModule; $testModule = $null
    $workbook.Close($false); Release-ComObject $workbook; $workbook = $null
    if (-not $ok) { exit 2 }
    exit 0
} catch {
    Write-Error $_
    exit 1
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $testModule; Release-ComObject $workbook; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
