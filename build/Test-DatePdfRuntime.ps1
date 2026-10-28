[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$output = Join-Path (Split-Path -Parent $BuiltWorkbook) 'date-pdf-test'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$xml = '<HDon><DLHDon Id="FIXTURE"><TTChung><THDon>Invoice fixture</THDon><KHMSHDon>1</KHMSHDon><KHHDon>FIXTURE</KHHDon><SHDon>42</SHDon><NLap>2026-10-28</NLap><DVTTe>VND</DVTTe><HTTToan>CK</HTTToan></TTChung><NDHDon><NBan><Ten>Fixture Seller &amp; Co</Ten><MST>0000000001</MST><DChi>Fixture address</DChi></NBan><NMua><Ten>Fixture Buyer</Ten><MST>0000000002</MST></NMua><DSHHDVu><HHDVu><STT>1</STT><TChat>1</TChat><THHDVu>Fixture Item &lt;script&gt; {{TOTAL}}</THHDVu><DVTinh>Item</DVTinh><SLuong>1.5</SLuong><DGia>10000</DGia><ThTien>15000</ThTien><TSuat>10%</TSuat></HHDVu></DSHHDVu><TToan><THTTLTSuat><LTSuat><TSuat>10%</TSuat><ThTien>15000</ThTien><TThue>1500</TThue></LTSuat></THTTLTSuat><TgTCThue>15000</TgTCThue><TgTThue>1500</TgTThue><TgTTTBSo>16500</TgTTTBSo><TgTTTBChu>Fixture amount in words</TgTTTBChu></TToan></NDHDon></DLHDon><DSCKS><NBan><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><Object><SignatureProperties><SignatureProperty><SigningTime>2026-10-28T10:00:00+07:00</SigningTime></SignatureProperty></SignatureProperties></Object></Signature></NBan></DSCKS></HDon>'
$utf8 = New-Object Text.UTF8Encoding($false)
$xmlPath = Join-Path $output 'fixture.xml'
[IO.File]::WriteAllText($xmlPath, $xml, $utf8)
$excel = $null; $book = $null; $module = $null; $form = $null
$results = [ordered]@{}
function Assert-Condition([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Invoke-TestMacro([string]$Name, [object[]]$MacroArguments) {
    # Pass literals inside VBA. Excel.Run can coerce string arguments (notably
    # numeric IDs); a no-argument entry point also avoids COM binder differences.
    if ($Name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw 'Invalid test macro name' }
    $literals = @($MacroArguments | ForEach-Object {
        $literal = '"' + ([string]$_).Replace('"','""') + '"'
        [regex]::Replace($literal, '[^\x00-\x7F]', { param($m) '" & ChrW(' + [int][char]$m.Value + ') & "' })
    })
    $code = $module.CodeModule
    try { $start = $code.ProcStartLine('DatePdfInvoke',0); $code.DeleteLines($start,$code.ProcCountLines('DatePdfInvoke',0)) } catch {}
    $code.AddFromString(('Public Function DatePdfInvoke() As Variant' + "`r`n" + 'DatePdfInvoke = ' + $Name + '(' + ($literals -join ', ') + ')' + "`r`nEnd Function"))
    return $excel.Run("${prefix}DatePdfInvoke")
}
try {
    $excel = New-ExcelApplication
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($BuiltWorkbook, 0, $false)
    $prefix = "'$($book.Name)'!"
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'modDatePdfTest'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function DatePeriodAudit() As String
    Dim months As Variant, p As Variant, sources As Variant, source As Variant, i As Long, last As Date
    Dim message As String, days As Variant, previousCount As Long
    ReDim months(1 To 2, 1 To 2)
    months(1, 1) = DateSerial(2026, 9, 28): months(1, 2) = DateSerial(2026, 9, 30)
    months(2, 1) = DateSerial(2026, 10, 1): months(2, 2) = DateSerial(2026, 10, 27)
    For Each source In Array("all", "query", "sco-query")
        p = GdtListPeriods(months, CStr(source))
        If source = "all" And UBound(p, 1) <> 6 Then Err.Raise 5, , "All-source period count"
        For Each sources In Array("query", "sco-query")
            last = 0
            For i = 1 To UBound(p, 1)
                If p(i, 3) = sources Then
                    If last = 0 Then last = months(1, 1) - 1
                    If p(i, 1) <> last + 1 Then Err.Raise 5, , "Gap/overlap in date periods"
                    If p(i, 2) < p(i, 1) Then Err.Raise 5, , "Inverted window"
                    If sources = "sco-query" And p(i, 2) - p(i, 1) + 1 > 10 Then Err.Raise 5, , "SCO window > 10"
                    last = p(i, 2)
                End If
            Next i
            If last <> 0 And last <> months(2, 2) Then Err.Raise 5, , "Missing tail date"
            If source <> "all" And source <> sources And last <> 0 Then Err.Raise 5, , "Unselected source included"
        Next sources
    Next source
    For Each days In Array(10, 5, 3, 1, 17)
        message = "HTTP 400: {""message"":""" & UniConvert("Khoarng thowfi gian tifm kieesm khoong dduwowjc lowsn hown ") & days & UniConvert(" ngafy.") & """}"
        If GdtReportedDayLimit(400, message) <> days Then Err.Raise 5, , "Dynamic limit parser"
        If GdtReportedDayLimit(500, message) <> 0 Then Err.Raise 5, , "Wrong status accepted"
        p = GdtListPeriods(months, "all")
        previousCount = UBound(p, 1)
        If Not GdtSplitListPeriod(p, 3, CLng(days)) Then Err.Raise 5, , "Rejected window not split"
        If p(1, 3) <> "query" Or p(1, 2) <> months(1, 2) Then Err.Raise 5, , "Earlier window changed"
        If p(2, 3) <> "sco-query" Or p(2, 2) <> months(1, 2) Then Err.Raise 5, , "Other source changed"
        last = months(2, 1) - 1
        For i = 3 To UBound(p, 1) - (previousCount - 3)
            If p(i, 3) <> "query" Or p(i, 1) <> last + 1 Then Err.Raise 5, , "Split continuity"
            If p(i, 2) - p(i, 1) + 1 > days Then Err.Raise 5, , "Split exceeds limit"
            last = p(i, 2)
        Next i
        If last <> months(2, 2) Then Err.Raise 5, , "Split omitted tail"
        If GdtSplitListPeriod(p, 1, 0) Or GdtSplitListPeriod(p, 1, 999) Then Err.Raise 5, , "Invalid limit accepted"
    Next days
    If GdtReportedDayLimit(400, "HTTP 400: invalid date 2026") <> 0 Then Err.Raise 5, , "Other 400 treated as date limit"
    DatePeriodAudit = "PASS"
End Function
Public Function PdfRejects(ByVal xmlPath As String, Optional ByVal identity As String = "") As Boolean
    Dim html As String
    On Error GoTo Rejected
    html = GdtBuildInvoiceHtml(xmlPath, identity)
    Exit Function
Rejected:
    PdfRejects = True
End Function
Public Function PdfFixtureIdentity() As String
    PdfFixtureIdentity = GdtInvoiceIdentity("0000000001", "1", "FIXTURE", "42")
End Function
Public Function PdfBuildSafe(ByVal xmlPath As String, ByVal identity As String) As String
    Dim html As String, stream As Object
    On Error GoTo Failed
    html = GdtBuildInvoiceHtml(xmlPath, identity)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText html: stream.SaveToFile xmlPath & ".html", 2: stream.Close
    PdfBuildSafe = "PASS"
    Exit Function
Failed:
    PdfBuildSafe = "ERROR: " & Err.Number & " " & Err.Description
End Function
Public Function PdfExportSafe(ByVal xmlPath As String, ByVal pdfPath As String, Optional ByVal identity As String = "") As String
    On Error GoTo Failed
    GdtExportInvoicePdf xmlPath, pdfPath, identity
    PdfExportSafe = "PASS"
    Exit Function
Failed:
    PdfExportSafe = "ERROR: " & Err.Number & " " & Err.Description
End Function
'@)
    $form = $book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $form.AddFromString(@'
Public Function PdfUiAudit() As String
    Dim check As Object
    Set check = Me.fraLoaiDuLieu.Controls("chkPdf")
    If check.Value Then Err.Raise 5, , "PDF should be opt-in"
    If check.Left + check.Width > Me.fraLoaiDuLieu.InsideWidth + 1 Then Err.Raise 5, , "PDF checkbox clipped"
    check.Value = True
    If Not Me.txtXMLFolderPath.Enabled Or Not Me.cmdChonFolder.Enabled Then Err.Raise 5, , "PDF-only folder disabled"
    If Me.chkXmlZip.Value Then Err.Raise 5, , "PDF selected an extra option"
    check.Value = False
    If Me.txtXMLFolderPath.Enabled Then Err.Raise 5, , "Folder stale after PDF deselected"
    If UiCtl("cmdXmlToPdf").Left < Me.cmdChiTietXML.Left + Me.cmdChiTietXML.Width Then Err.Raise 5, , "PDF button overlap"
    PdfUiAudit = "PASS"
End Function
'@)
    $module.CodeModule.AddFromString(@'
Public Function PdfUiTest() As String
    Dim f As frmTaiHoaDon
    Set f = New frmTaiHoaDon
    PdfUiTest = f.PdfUiAudit()
    Unload f
End Function
'@)
    $excel.VBE.ActiveVBProject = $book.VBProject
    $excel.VBE.CommandBars.FindControl(1, 578).Execute()
    Assert-Condition ([string]$excel.Run("${prefix}DatePeriodAudit") -eq 'PASS') 'Date period audit failed'
    Write-Output 'Date periods: PASS'
    $results.DatePeriods = 'PASS: source isolation, continuity, 10/5/3/1/17-day limits'
    Assert-Condition ([string]$excel.Run("${prefix}PdfUiTest") -eq 'PASS') 'PDF UI audit failed'
    Write-Output 'PDF UI: PASS'
    $results.PdfUi = 'PASS'
    $identity = [string]$excel.Run("${prefix}PdfFixtureIdentity")
    foreach ($variant in @('Standard','DefaultNamespace','PrefixNamespace','ForeignField','PaddedNumber','BadNumber','DTD','MissingItems')) {
        Write-Output "Checking $variant"
        $fixture = switch ($variant) {
            'DefaultNamespace' { $xml.Replace('<HDon>', '<HDon xmlns="urn:fixture:invoice">') }
            'PrefixNamespace' { [regex]::Replace($xml.Replace(' xmlns="http://www.w3.org/2000/09/xmldsig#"',''), '<(/?)([A-Za-z][A-Za-z0-9]*)', '<$1inv:$2').Replace('<inv:HDon>', '<inv:HDon xmlns:inv="urn:fixture:invoice">') }
            'ForeignField' { $xml.Replace('<ThTien>15000</ThTien>', '<ThTien>15000</ThTien><f:ThTien xmlns:f="urn:foreign">999999999</f:ThTien>') }
            'PaddedNumber' { $xml.Replace('<SHDon>42</SHDon>','<SHDon>00000042</SHDon>') }
            'BadNumber' { $xml.Replace('<SLuong>1.5</SLuong>', '<SLuong>1,5</SLuong>') }
            'DTD' { '<!DOCTYPE HDon [<!ENTITY external SYSTEM "file:///C:/fixture-secret">]>' + $xml }
            'MissingItems' { [regex]::Replace($xml, '<DSHHDVu>.*?</DSHHDVu>', '') }
            default { $xml }
        }
        $path = Join-Path $output "$variant.xml"
        [IO.File]::WriteAllText($path, $fixture, $utf8)
        if ($variant -in @('BadNumber','DTD','MissingItems')) {
            Assert-Condition ([bool](Invoke-TestMacro 'PdfRejects' @($path))) "$variant accepted"
        } else {
            $buildResult = [string](Invoke-TestMacro 'PdfBuildSafe' @($path, $identity))
            Assert-Condition ($buildResult -eq 'PASS') $buildResult
            $html = [IO.File]::ReadAllText($path + '.html', [Text.Encoding]::UTF8)
            Assert-Condition (-not $html.StartsWith('ERROR:')) $html
            Assert-Condition ($html.Contains('Fixture Seller &amp; Co') -and $html.Contains('0000000001')) "$variant lost seller data"
            Assert-Condition ($html.Contains('15.000') -and $html.Contains('16.500') -and $html.Contains('1,5')) "$variant numeric formatting"
            Assert-Condition ($html.Contains('&lt;script&gt;') -and $html.Contains('&#123;&#123;TOTAL&#125;&#125;') -and $html -notmatch '<script|999999999|\{\{') "$variant unsafe text / foreign field / unresolved token"
            Assert-Condition ($html.Contains('2026-10-28T10:00:00+07:00')) "$variant signature timestamp"
            if ($variant -eq 'Standard') { [IO.File]::WriteAllText((Join-Path $output 'fixture.html'), $html, $utf8) }
        }
        $results[$variant] = 'PASS'
    }
    Assert-Condition ([bool](Invoke-TestMacro 'PdfRejects' @($xmlPath, 'wrong-identity'))) 'Identity mismatch accepted'
    $pdfPath = Join-Path $output 'fixture.pdf'
    $exportResult = [string](Invoke-TestMacro 'PdfExportSafe' @($xmlPath, $pdfPath, $identity))
    Assert-Condition ($exportResult -eq 'PASS') $exportResult
    Assert-Condition ([bool](Invoke-TestMacro 'GdtIsPdfFile' @($pdfPath))) 'Exported file is not PDF'
    $pdfBefore = (Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash
    $exportResult = [string](Invoke-TestMacro 'PdfExportSafe' @((Join-Path $output 'BadNumber.xml'), $pdfPath))
    Assert-Condition ($exportResult.StartsWith('ERROR:')) 'Invalid XML export unexpectedly succeeded'
    Assert-Condition ((Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash -eq $pdfBefore) 'Existing PDF was damaged'
    $results.PdfExport = 'PASS: browser renderer, signature, preserve previous PDF on failure'
    $longXml = $xml.Replace('</DSHHDVu>', (([regex]::Match($xml,'<HHDVu>.*?</HHDVu>').Value * 60) + '</DSHHDVu>'))
    $longPath = Join-Path $output 'long.xml'
    [IO.File]::WriteAllText($longPath, $longXml, $utf8)
    $exportResult = [string](Invoke-TestMacro 'PdfExportSafe' @($longPath, (Join-Path $output 'long.pdf')))
    Assert-Condition ($exportResult -eq 'PASS') $exportResult
    $results.LongInvoice = 'PASS: 61 line items rendered'
    $specialFolder = Join-Path $output (([char]0x0111).ToString() + ' space#%&')
    New-Item -ItemType Directory -Path $specialFolder -Force | Out-Null
    $specialXml = Join-Path $specialFolder 'fixture.xml'
    [IO.File]::WriteAllText($specialXml, $xml, $utf8)
    $exportResult = [string](Invoke-TestMacro 'PdfExportSafe' @($specialXml, (Join-Path $specialFolder 'fixture.pdf')))
    Assert-Condition ($exportResult -eq 'PASS') $exportResult
    $results.UnicodePath = 'PASS: Unicode, spaces, #, % and &'
} finally {
    $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'result.json') -Encoding UTF8
    Release-ComObject $form; Release-ComObject $module
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $excel
}
Assert-Condition ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -eq $before) 'Runtime test changed the distributable workbook'
$results | ConvertTo-Json -Depth 5
Write-Output 'DATE/PDF RUNTIME TESTS PASSED'
