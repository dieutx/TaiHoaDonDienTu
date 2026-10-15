[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$fixtureRoot = Join-Path $env:TEMP ('hddt-compat-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$xml = '<HDon><DLHDon Id="FIXTURE"><TTChung><KHHDon>FIXTURE</KHHDon><SHDon>42</SHDon><NLap>2026-10-13</NLap><DVTTe>VND</DVTTe><TGia>1.25</TGia></TTChung><NDHDon><NBan><Ten>Fixture Seller</Ten><MST>0000000001</MST></NBan><NMua><Ten>Fixture Buyer</Ten><MST>0000000002</MST></NMua><DSHHDVu><HHDVu><STT>1</STT><THHDVu>Fixture Item</THHDVu><SLuong>1.5</SLuong><ThTien>100.5</ThTien><TSuat>10%</TSuat><TTKhac><TTin><TTruong>fixture_tax</TTruong><DLieu>10.05</DLieu></TTin></TTKhac></HHDVu><HHDVu><STT>2</STT><ThTien>10</ThTien><TSuat>KCT</TSuat></HHDVu></DSHHDVu><TToan><TgTThue>10.05</TgTThue></TToan></NDHDon></DLHDon><DSCKS><NBan><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><Object><SignatureProperties><SignatureProperty><SigningTime>2026-10-12T17:00:00Z</SigningTime></SignatureProperty></SignatureProperties></Object></Signature></NBan></DSCKS></HDon>'
$fixtures = [ordered]@{
    Standard = $xml
    DefaultNamespace = $xml.Replace('<HDon>', '<HDon xmlns="urn:fixture:invoice">')
    PrefixNamespace = [regex]::Replace($xml.Replace(' xmlns="http://www.w3.org/2000/09/xmldsig#"', ''), '<(/?)([A-Za-z][A-Za-z0-9]*)', '<$1inv:$2').Replace('<inv:HDon>', '<inv:HDon xmlns:inv="urn:fixture:invoice">')
    ForeignField = $xml.Replace('<ThTien>100.5</ThTien>', '<ThTien>100.5</ThTien><foreign:ThTien xmlns:foreign="urn:fixture:foreign">999</foreign:ThTien>')
    LiteralText = $xml.Replace('<Ten>Fixture Seller</Ten>', '<Ten>=1+1</Ten>').Replace('<THHDVu>Fixture Item</THHDVu>', "<THHDVu>'quote</THHDVu>")
    BadTax = $xml.Replace('<DLieu>10.05</DLieu>', '<DLieu>invalid</DLieu>')
    BadSecondLine = $xml.Replace('<ThTien>10</ThTien>', '<ThTien>10,5</ThTien>')
    BadDate = $xml.Replace('<NLap>2026-10-13</NLap>', '<NLap>2026-02-31</NLap>')
    EmptyTax = $xml.Replace('<DLieu>10.05</DLieu>', '<DLieu/>')
}
foreach ($name in $fixtures.Keys) {
    $folder = Join-Path $fixtureRoot $name
    New-Item -ItemType Directory -Path $folder | Out-Null
    [IO.File]::WriteAllText((Join-Path $folder 'fixture.xml'), $fixtures[$name])
}
$recoveryFolder = Join-Path $fixtureRoot 'Recovery'
New-Item -ItemType Directory -Path $recoveryFolder | Out-Null
[IO.File]::WriteAllText((Join-Path $recoveryFolder '00-bad.xml'), $fixtures.BadSecondLine)
[IO.File]::WriteAllText((Join-Path $recoveryFolder '01-good.XML'), $xml)

function Assert-Condition([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Cell-Value([object]$Sheet, [int]$Row, [int]$Column) {
    $cell = $Sheet.Cells.Item($Row, $Column)
    try { return $cell.Value2 } finally { Release-ComObject $cell }
}
function Cell-Format([object]$Sheet, [int]$Row, [int]$Column) {
    $cell = $Sheet.Cells.Item($Row, $Column)
    try { return [string]$cell.NumberFormat } finally { Release-ComObject $cell }
}
$excel = $null; $workbook = $null; $module = $null; $sheet = $null; $report = $null; $lookup = $null
$results = [ordered]@{}
try {
    Write-Output 'Opening compatibility fixture workbook...'
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'$($workbook.Name)'!"
    $report = $workbook.Worksheets.Item('BaoCao_LoiTaiHD')
    $lookup = $workbook.Worksheets.Item('LinkTraCuu')
    $lookup.Range('O14').Value2 = 'fixture_tax'
    $lookup.Range('O15').Value2 = 'fixture_total'
    foreach ($name in @($fixtures.Keys) + @('Recovery')) {
        foreach ($purchase in @($true, $false)) {
            $label = $name + $(if ($purchase) { '-Purchase' } else { '-Sales' })
            Write-Output "Running $label"
            $sheet = $workbook.Worksheets.Item($(if ($purchase) { 'ChiTietHD_Mua_XML' } else { 'ChiTietHD_Ban_XML' }))
            $sheet.Range('A3:AI100').ClearContents()
            $sheet.Range('A3:AI100').NumberFormat = 'General'
            $report.Rows('2:100').ClearContents()
            $xmlPath = [string](Join-Path $fixtureRoot $name)
            [void]$excel.Run("${prefix}parseXML", $xmlPath, [bool]$purchase, [bool]$false)
            if ($name -in @('BadTax', 'BadSecondLine', 'BadDate')) {
                Assert-Condition ((Cell-Value $sheet 3 1) -eq $null -and (Cell-Value $sheet 4 1) -eq $null) "${label}: partial invoice retained"
                Assert-Condition ([string](Cell-Value $report 2 13) -ne '') "${label}: per-file error missing"
            } else {
                foreach ($row in @(3, 4)) {
                    Assert-Condition ((Cell-Value $sheet $row 1) -eq 'FIXTURE') "$label missing invoice"
                    Assert-Condition ((Cell-Value $sheet $row 8) -ceq '0000000001') "$label seller ID lost zeros"
                    Assert-Condition ((Cell-Value $sheet $row 14) -ceq '0000000002') "$label buyer ID lost zeros"
                    Assert-Condition ((Cell-Format $sheet $row 8) -eq '@' -and (Cell-Format $sheet $row 14) -eq '@') "$label tax ID format"
                    $dateSerial = ([datetime]'2026-10-13').ToOADate()
                    Assert-Condition ((Cell-Value $sheet $row 4) -eq $dateSerial) "$label invoice date value"
                    Assert-Condition ((Cell-Value $sheet $row 10) -eq $dateSerial) "$label signature namespace / UTC+7"
                    Assert-Condition ((Cell-Format $sheet $row 4) -eq 'dd/mm/yyyy' -and (Cell-Format $sheet $row 10) -eq 'dd/mm/yyyy') "$label date format"
                }
                Assert-Condition ([math]::Abs((Cell-Value $sheet 3 26) - 10.05) -lt 0.00001) "$label tax"
                Assert-Condition ([math]::Abs((Cell-Value $sheet 3 27) - 110.55) -lt 0.00001) "$label total"
                Assert-Condition ((Cell-Value $sheet 3 20) -eq 1.5) "$label quantity"
                Assert-Condition ((Cell-Value $sheet 4 26) -eq 0 -and (Cell-Value $sheet 4 27) -eq 10) "$label KCT"
                Assert-Condition ((Cell-Value $sheet 5 1) -eq $null) "$label unexpected row"
                if ($name -eq 'LiteralText') {
                    foreach ($row in @(3,4)) {
                        Assert-Condition ((Cell-Value $sheet $row 7) -ceq '=1+1') "$label common text changed during filldown"
                        $cell = $sheet.Cells.Item($row, 7)
                        try { Assert-Condition (-not $cell.HasFormula) "$label common text became formula" } finally { Release-ComObject $cell }
                    }
                    Assert-Condition ((Cell-Value $sheet 3 18) -ceq "'quote") "$label item apostrophe stripped"
                }
                if ($name -eq 'Recovery') {
                    Assert-Condition ([string](Cell-Value $report 2 13) -ne '') "$label error missing"
                } else {
                    Assert-Condition ((Cell-Value $report 2 1) -eq $null) "$label unexpected error"
                    $report.Cells.Item(2, 1).Value2 = 1
                    $report.Cells.Item(2, 3).Value2 = $(if ($purchase) { 'mua' } else { 'ban' })
                    $report.Cells.Item(2, 4).Value2 = 'XML'
                    $report.Cells.Item(2, 7).Value2 = 'fixture.xml'
                    $report.Cells.Item(2, 10).Value2 = 'Parse XML'
                    $report.Cells.Item(2, 13).Value2 = 'fixture'
                    [void]$excel.Run("${prefix}parseXML", $xmlPath, [bool]$purchase, [bool]$false)
                    Assert-Condition ((Cell-Value $report 2 1) -eq $null) "$label corrected file retained error"
                }
            }
            $results[$label] = 'PASS'
            Write-Output "$label=PASS"
            Release-ComObject $sheet; $sheet = $null
        }
    }
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'modCompatAuthTest'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function CompatAuth(ByVal jsonText As String, ByVal key As String, ByVal expected As String, ByVal shouldFail As Boolean) As Boolean
    On Error GoTo Rejected
    Dim parsed As Object, actual As String
    Set parsed = ParseGdtAuthObject(jsonText)
    actual = GdtRequiredAuthText(parsed, key)
    CompatAuth = (Not shouldFail And actual = expected)
    Exit Function
Rejected:
    CompatAuth = (shouldFail And InStr(Err.Description, "SENSITIVE-FIXTURE") = 0)
End Function
'@)
    $authCases = @(
        @('ValidToken', '{"token":"fixture-message-valid"}', 'token', 'fixture-message-valid', $false),
        @('ValidCaptcha', '{"key":"fixture-key","content":"fixture-svg"}', 'content', 'fixture-svg', $false),
        @('Malformed', '{SENSITIVE-FIXTURE', 'token', '', $true),
        @('Array', '[]', 'token', '', $true),
        @('Scalar', '"SENSITIVE-FIXTURE"', 'token', '', $true),
        @('Missing', '{}', 'token', '', $true),
        @('Null', '{"token":null}', 'token', '', $true),
        @('Numeric', '{"token":123}', 'token', '', $true),
        @('Blank', '{"token":"  "}', 'token', '', $true),
        @('Object', '{"token":{}}', 'token', '', $true),
        @('MissingCaptcha', '{"key":"fixture-key"}', 'content', '', $true)
    )
    foreach ($case in $authCases) {
        Assert-Condition ([bool]$excel.Run("${prefix}CompatAuth", $case[1], $case[2], $case[3], $case[4])) "Auth-$($case[0]) failed"
        $results['Auth-' + $case[0]] = 'PASS'
        Write-Output "Auth-$($case[0])=PASS"
    }
    $sheet = $workbook.Worksheets.Item('ChiTietHD_Mua')
    $sheet.Range('A450:AH453').ClearContents()
    [void]$excel.Run("${prefix}LinkTraCuu")
    [void]$excel.Run("${prefix}tenCotTraCuu")
    $guid = '00000000-0000-0000-0000-000000000001'
    $json = '{"id":"' + $guid + '","msttcgp":"0101360697","hdhhdvu":[{"thtien":100,"tsuat":0.1}]}'
    [void]$excel.Run("${prefix}ghiExcel_ChiTiet", $json, [int]450, 'mua')
    Assert-Condition ((Cell-Value $sheet 450 33) -eq "https://van.ehoadon.vn/Lookup?InvoiceGUID=$guid") 'BKAV detail link'
    [void]$excel.Run("${prefix}ghiExcel_ChiTiet", $json.Replace($guid, ''), [int]451, 'mua')
    Assert-Condition ([string](Cell-Value $sheet 451 33) -like '*GUID*' -and [string](Cell-Value $sheet 451 33) -notlike 'https://*') 'BKAV missing GUID'
    Assert-Condition ([string]$excel.Run("${prefix}GdtBkavLookupUrl", 'invalid&redirect=fixture') -eq '') 'BKAV invalid GUID'
    $results['BkavDetail'] = 'PASS'
    Write-Output 'BkavDetail=PASS'
} finally {
    $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $BuiltWorkbook) 'compatibility-fixes-runtime-result.json') -Encoding UTF8
    Release-ComObject $sheet; Release-ComObject $report; Release-ComObject $lookup; Release-ComObject $module
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $workbook; Release-ComObject $excel
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -match '^hddt-compat-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
Write-Output 'COMPATIBILITY FIXES RUNTIME TEST PASSED'
