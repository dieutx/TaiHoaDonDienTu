[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$output = Join-Path (Split-Path -Parent $BuiltWorkbook) 'pdf-fallback-test'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$before = (Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash
$utf8 = New-Object Text.UTF8Encoding($false)
$detail = [ordered]@{
    nbmst='0000000001'; khmshdon='1'; khhdon='FIXTURE'; shdon=42; tdlap='2026-10-28'; nbten='Fixture Seller & Co'; nmten='Fixture Buyer'
    dvtte='VND'; tgtcthue=$null; tgtthue=$null; tgtttbso=$null
    hdhhdvu=@([ordered]@{stt=1;tchat=1;ten='Fixture <script> {{TOTAL}}';dvtinh='Unit';sluong=1.5;dgia=10000.5;thtien=15000.75;tsuat=0.1;tthue=1500.075})
    thttltsuat=@()
}
$summary = [ordered]@{
    nbmst='0000000001';khmshdon='1';khhdon='FIXTURE';shdon=42;tgtcthue=15000.75;tgtthue=1500.075;tgtttbso=16500.825
    thttltsuat=@([ordered]@{tsuat=0.1;thtien=15000.75;tthue=1500.075})
}
$detailPath = Join-Path $output 'detail.json'; $summaryPath = Join-Path $output 'summary.json'
[IO.File]::WriteAllText($detailPath, ($detail | ConvertTo-Json -Depth 10 -Compress), $utf8)
[IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 10 -Compress), $utf8)
$excel=$null; $book=$null; $module=$null; $form=$null; $results=[ordered]@{}
function Assert-Condition([bool]$Condition,[string]$Message) { if(-not $Condition){throw $Message} }
function Invoke-Case([string]$DetailFile,[string]$SummaryFile,[string]$OutputName,[bool]$Render=$false) {
    $code=$module.CodeModule
    try {$start=$code.ProcStartLine('FallbackEntry',0);$code.DeleteLines($start,$code.ProcCountLines('FallbackEntry',0))} catch {}
    $args=@($DetailFile,$SummaryFile,(Join-Path $output $OutputName)) | ForEach-Object {'"'+$_.Replace('"','""')+'"'}
    $code.AddFromString(('Public Function FallbackEntry() As String' + "`r`nFallbackEntry = RunPdfFallback(" + ($args -join ',') + ',' + $Render.ToString() + ")`r`nEnd Function"))
    return [string]$excel.Run("'$($book.Name)'!FallbackEntry")
}
try {
    $excel=New-ExcelApplication
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($BuiltWorkbook,0,$false)
    $module=$book.VBProject.VBComponents.Add(1);$module.Name='modPdfFallbackTest'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public DetailFixture As String, DetailHits As Long, LastDetailEndpoint As String
Public Function ReadFixture(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile path
    ReadFixture = stream.ReadText: stream.Close
End Function
Public Function RunPdfFallback(ByVal detailFile As String, ByVal summaryFile As String, ByVal outputPath As String, ByVal render As Boolean) As String
    Dim detail As String, summary As String, html As String, stream As Object, identity As String
    On Error GoTo Failed
    detail = ReadFixture(detailFile)
    If Len(summaryFile) > 0 Then summary = ReadFixture(summaryFile)
    identity = GdtPdfIdentity("0000000001", "1", "FIXTURE", "42")
    html = GdtBuildInvoiceHtmlFromDetail(detail, summary, identity)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText html
    stream.SaveToFile outputPath & ".html", 2: stream.Close
    If render Then GdtExportDetailPdf detail, outputPath & ".pdf", summary, identity
    RunPdfFallback = "PASS"
    Exit Function
Failed:
    RunPdfFallback = "ERROR: " & Err.Number & " " & Err.Description
End Function
Public Function PdfFallbackFixtureResult(ByVal endpoint As String) As clsGdtRequestResult
    Dim result As New clsGdtRequestResult
    DetailHits = DetailHits + 1: LastDetailEndpoint = endpoint
    result.Success = True: result.StatusCode = 200: result.Attempts = 1: result.ResponseText = DetailFixture
    Set PdfFallbackFixtureResult = result
End Function
'@)
    Assert-Condition ((Invoke-Case $detailPath $summaryPath 'api' $true) -eq 'PASS') 'API PDF with null detail totals failed'
    $html=[IO.File]::ReadAllText((Join-Path $output 'api.html'),[Text.Encoding]::UTF8)
    Assert-Condition ($html.Contains('15.000,75') -and $html.Contains('16.500,825') -and $html.Contains('10%')) 'API amounts/rate changed'
    Assert-Condition ($html.Contains('Fixture Seller &amp; Co') -and $html.Contains('&lt;script&gt;') -and $html.Contains('&#123;&#123;TOTAL&#125;&#125;')) 'API HTML escaping'
    Assert-Condition ($html.Contains('API') -and $html -notmatch '\{\{|2026-10-28T10:00') 'API source/signature disclosure'
    $results.NullDetailTotals='PASS: list totals + detail lines, decimal amounts, escaped text'
    $pdfPath=Join-Path $output 'api.pdf';$pdfHash=(Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash
    foreach($case in @('MissingItems','MissingTotal','WrongDetailIdentity','WrongSummaryIdentity','BadNumber','ListOnly')) {
        $d=$detail|ConvertTo-Json -Depth 10 -Compress|ConvertFrom-Json
        $s=$summary|ConvertTo-Json -Depth 10 -Compress|ConvertFrom-Json
        switch($case) {
            'MissingItems' {$d.hdhhdvu=@()}
            'MissingTotal' {$s.tgtttbso=$null}
            'WrongDetailIdentity' {$d.nbmst='0000000999'}
            'WrongSummaryIdentity' {$s.shdon=99}
            'BadNumber' {$d.hdhhdvu[0].thtien='bad'}
            'ListOnly' {$d.PSObject.Properties.Remove('hdhhdvu')}
        }
        $dPath=Join-Path $output ($case+'-detail.json');$sPath=Join-Path $output ($case+'-summary.json')
        [IO.File]::WriteAllText($dPath,($d|ConvertTo-Json -Depth 10 -Compress),$utf8)
        [IO.File]::WriteAllText($sPath,($s|ConvertTo-Json -Depth 10 -Compress),$utf8)
        $result=Invoke-Case $dPath $sPath 'api' $true
        Assert-Condition ($result.StartsWith('ERROR:')) "$case unexpectedly produced PDF"
        Assert-Condition ((Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash -eq $pdfHash) "$case damaged previous PDF"
        $results[$case]='PASS: rejected; previous PDF retained'
    }
    # Exercise the actual form phase. Only its HTTP transport is replaced;
    # summary-row validation, missing-XML branch and PDF rendering stay real.
    $form=$book.VBProject.VBComponents.Item('frmTaiHoaDon').CodeModule
    $start=$form.ProcStartLine('ExportInvoicePdfFromApi',0);$count=$form.ProcCountLines('ExportInvoicePdfFromApi',0)
    $text=$form.Lines($start,$count)
    $pattern='Set result = ExecuteGdtRequest\("GET", endpoint, RunToken\(\), vbNullString, "application/json", _\r?\n\s*"application/json, text/plain, \*/\*", False, GDT_MAX_RETRIES, "PDF DETAIL"\)'
    Assert-Condition ([regex]::IsMatch($text,$pattern)) 'PDF detail transport seam missing'
    $text=[regex]::Replace($text,$pattern,'Set result = PdfFallbackFixtureResult(endpoint)')
    $form.DeleteLines($start,$count);$form.InsertLines($start,$text)
    $form.AddFromString(@'
Public Function AuditPdfNoXml(ByVal folder As String, ByVal invoiceType As Long, ByVal source As Long) As String
    On Error GoTo Failed
    Dim ws As Worksheet, record As Object, path As String, fso As Object
    ResetGdtOperationControl: ResetGdtRequestSession
    mRunPurchase = (invoiceType = 1): mRunSource = IIf(source = 1, "query", "sco-query")
    Set ws = ThisWorkbook.Worksheets(IIf(mRunPurchase, "TongHopHD_Mua", "TongHopHD_Ban"))
    ws.Range("A3:BN3").ClearContents
    ws.Cells(3,3).Value2 = "1": ws.Cells(3,4).Value2 = "FIXTURE": ws.Cells(3,5).Value2 = 42
    ws.Cells(3,9).Value2 = "Fixture Seller & Co": ws.Cells(3,10).NumberFormat = "@": ws.Cells(3,10).Value2 = "0000000001"
    ws.Cells(3,42).Value2 = 15000.75: ws.Cells(3,44).Value2 = 1500.075: ws.Cells(3,49).Value2 = 16500.825
    Set record = JsonConverter.ParseJSON(DetailFixture)
    GdtReportAccount = "FIXTURE": GdtReportSession = "PDF-FALLBACK-TEST"
    EnsureGdtErrorReportSheet.Range("A2:S100").ClearContents
    ReDim arrHDChiTiet(1,7)
    arrHDChiTiet(0,0) = "0000000001": arrHDChiTiet(0,1) = "FIXTURE": arrHDChiTiet(0,2) = "42": arrHDChiTiet(0,3) = "1"
    arrHDChiTiet(0,4) = source: arrHDChiTiet(0,5) = DateSerial(2026,10,28): arrHDChiTiet(0,6) = 3
    oSaveUnzipFolder = folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(folder, GdtInvoiceFileBase("0000000001", "1", "FIXTURE", "42"))
    If fso.FileExists(path & ".xml") Then Err.Raise 5, , "Fixture must have no XML"
    DetailHits = 0
    ProcessInvoicePdfs 1
    If DetailHits <> 1 Or InStr(LastDetailEndpoint, "/api/" & mRunSource & "/invoices/detail?") = 0 Then Err.Raise 5, , "Wrong detail source/request count"
    If Not GdtIsPdfFile(path & ".pdf") Then Err.Raise 5, , "No fallback PDF saved"
    If CountGdtSessionErrors(CurrentDirectionName()) <> 0 Then Err.Raise 5, , "Unexpected PDF error"
    If fso.FileExists(path & ".xml") Then Err.Raise 5, , "Fabricated original XML"
    ws.Cells(3,5).Value2 = 99
    On Error Resume Next
    Dim rejected As String
    rejected = GdtPdfSummaryFromRow(ws,3,GdtPdfIdentity("0000000001","1","FIXTURE","42"))
    If Err.Number = 0 Then AuditPdfNoXml = "ERROR: wrong summary row accepted": Exit Function
    Err.Clear: On Error GoTo Failed
    AuditPdfNoXml = "PASS"
    Exit Function
Failed:
    AuditPdfNoXml = "ERROR: " & Err.Number & " " & Err.Description
End Function
'@)
    $dLiteral=$detailPath.Replace('"','""');$folderLiteral=$output.Replace('"','""')
    $module.CodeModule.AddFromString("Public Function FormFallbackEntry() As String`r`nDim f As frmTaiHoaDon, direction As Long, source As Long, result As String`r`nDetailFixture = ReadFixture(`"$dLiteral`")`r`nFor direction = 1 To 2`r`nFor source = 1 To 2`r`nSet f = New frmTaiHoaDon`r`nresult = f.AuditPdfNoXml(`"$folderLiteral`", direction, source)`r`nUnload f`r`nIf result <> `"PASS`" Then FormFallbackEntry = result: Exit Function`r`nNext source`r`nNext direction`r`nFormFallbackEntry = `"PASS`"`r`nEnd Function")
    $result=[string]$excel.Run("'$($book.Name)'!FormFallbackEntry")
    Assert-Condition ($result -eq 'PASS') $result
    $results.FormPipeline='PASS: no XML, purchase/sales, query/SCO, summary-row identity'
} finally {
    $results|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $output 'result.json') -Encoding UTF8
    Release-ComObject $form;Release-ComObject $module
    if($book){try{$book.Close($false)}catch{}}
    if($excel){try{$excel.Quit()}catch{}}
    Release-ComObject $book;Release-ComObject $excel
}
Assert-Condition ((Get-FileHash -LiteralPath $BuiltWorkbook -Algorithm SHA256).Hash -eq $before) 'Test changed workbook'
$results|ConvertTo-Json
Write-Output 'PDF FALLBACK RUNTIME TESTS PASSED'
