[CmdletBinding()]
param([Parameter(Mandatory)][string]$BuiltWorkbook)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'ExcelBuild.Common.psm1') -Force -DisableNameChecking
$BuiltWorkbook = (Resolve-Path -LiteralPath $BuiltWorkbook).Path
$expectedHeaders = Get-GdtDetailInvoiceHeaders -RepositoryRoot $repositoryRoot

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

    $layoutResults = @()
    foreach ($sheetName in @('ChiTietHD_Mua', 'ChiTietHD_Ban')) {
        $sheet = $workbook.Worksheets.Item($sheetName)
        $layoutResults += [ordered]@{
            Sheet = $sheetName
            TemplateHeader = [string]$sheet.Cells.Item(2, 1).Value2
            SeriesHeader = [string]$sheet.Cells.Item(2, 2).Value2
            InvoiceDateFormat = [string]$sheet.Columns.Item(4).NumberFormat
            SignatureDateFormat = [string]$sheet.Columns.Item(10).NumberFormat
            CodeDateFormat = [string]$sheet.Columns.Item(12).NumberFormat
            SellerTaxIdFormat = [string]$sheet.Columns.Item(8).NumberFormat
            BuyerTaxIdFormat = [string]$sheet.Columns.Item(14).NumberFormat
        }
        Release-ComObject $sheet
    }
    $layoutOk = @($layoutResults | Where-Object {
        $_.TemplateHeader -ne $expectedHeaders['1'] -or
        $_.SeriesHeader -ne $expectedHeaders['2'] -or
        $_.InvoiceDateFormat -ne 'dd/mm/yyyy' -or
        $_.SignatureDateFormat -ne 'dd/mm/yyyy' -or
        $_.CodeDateFormat -ne 'dd/mm/yyyy' -or
        $_.SellerTaxIdFormat -ne '@' -or
        $_.BuyerTaxIdFormat -ne '@'
    }).Count -eq 0

    $summaryLayoutResults = @()
    foreach ($sheetName in @('TongHopHD_Mua', 'TongHopHD_Ban')) {
        $sheet = $workbook.Worksheets.Item($sheetName)
        $summaryLayoutResults += [ordered]@{
            Sheet = $sheetName
            InvoiceDateFormat = [string]$sheet.Columns.Item(6).NumberFormat
            SignatureDateFormat = [string]$sheet.Columns.Item(12).NumberFormat
            CodeDateFormat = [string]$sheet.Columns.Item(14).NumberFormat
            OriginalInvoiceDateFormat = [string]$sheet.Columns.Item(62).NumberFormat
        }
        Release-ComObject $sheet
    }
    $summaryLayoutOk = @($summaryLayoutResults | Where-Object {
        $_.InvoiceDateFormat -ne 'dd/mm/yyyy' -or
        $_.SignatureDateFormat -ne 'dd/mm/yyyy' -or
        $_.CodeDateFormat -ne 'dd/mm/yyyy' -or
        $_.OriginalInvoiceDateFormat -ne 'dd/mm/yyyy'
    }).Count -eq 0
    if (-not ($layoutOk -and $summaryLayoutOk)) { throw 'Invoice headers or column formats are incorrect.' }

    $testModule = $workbook.VBProject.VBComponents.Add(1)
    $testModule.Name = 'modCodexDetailInvoiceTest'
    $testModule.CodeModule.AddFromString(@'
Option Explicit

Private Function RunDetailInvoiceFixture(ByVal sheetName As String, ByVal direction As String, _
    ByVal targetRow As Long, ByVal sellerTaxId As String, ByVal buyerTaxId As String) As Boolean
    On Error GoTo Failed
    Dim ws As Worksheet
    Dim sampleJson As String
    Dim writeRow As Long

    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Range("A" & targetRow & ":AH" & (targetRow + 1)).ClearContents
    sampleJson = "{""khmshdon"":""1"",""khhdon"":""C26TST"",""shdon"":42," & _
        """tdlap"":""2026-09-12T00:00:00"",""dvtte"":""VND"",""tgia"":1," & _
        """nbten"":""SELLER-FIXTURE"",""nbmst"":""" & sellerTaxId & """," & _
        """nbdchi"":""FIXTURE"",""nky"":""2026-09-12T00:00:00""," & _
        """mhdon"":""MCCQT-FIXTURE"",""ncma"":""2026-09-12T00:00:00""," & _
        """nmten"":""BUYER-FIXTURE"",""nmmst"":""" & buyerTaxId & """," & _
        """nmdchi"":""FIXTURE"",""msttcgp"":"""",""hdhhdvu"":[" & _
        "{""stt"":1,""tchat"":1,""mhhdvu"":""ITEM-1"",""ten"":""ITEM FIXTURE 1""," & _
        """dvtinh"":""UNIT"",""sluong"":2,""dgia"":100,""tlckhau"":0,""stckhau"":0," & _
        """ltsuat"":""10%"",""tsuat"":0.1,""thtien"":200,""tthue"":20,""thtcthue"":220}," & _
        "{""stt"":2,""tchat"":1,""mhhdvu"":""ITEM-2"",""ten"":""ITEM FIXTURE 2""," & _
        """dvtinh"":""UNIT"",""sluong"":1,""dgia"":200,""tlckhau"":0,""stckhau"":0," & _
        """ltsuat"":""10%"",""tsuat"":0.1,""thtien"":200,""tthue"":20,""thtcthue"":220}]," & _
        """tgtthue"":20}"

    writeRow = targetRow
    ghiExcel_ChiTiet sampleJson, writeRow, direction
    RunDetailInvoiceFixture = _
        (CStr(ws.Cells(2, 1).Value2) = UniConvert("Maaxu soos hosa ddown")) And _
        (CStr(ws.Cells(2, 2).Value2) = UniConvert("Kys hieeju HDD")) And _
        (CStr(ws.Cells(targetRow, 1).Value2) = "1") And _
        (CStr(ws.Cells(targetRow, 2).Value2) = "C26TST") And _
        (CLng(ws.Cells(targetRow, 3).Value2) = 42) And _
        (CDbl(ws.Cells(targetRow, 4).Value2) = CDbl(DateSerial(2026, 9, 12))) And _
        (ws.Cells(targetRow, 4).NumberFormat = "dd/mm/yyyy") And _
        (CDbl(ws.Cells(targetRow, 10).Value2) = CDbl(DateSerial(2026, 9, 12))) And _
        (CDbl(ws.Cells(targetRow, 12).Value2) = CDbl(DateSerial(2026, 9, 12))) And _
        (CStr(ws.Cells(targetRow, 8).Value2) = sellerTaxId) And _
        (CStr(ws.Cells(targetRow, 14).Value2) = buyerTaxId) And _
        (CStr(ws.Cells(targetRow + 1, 8).Value2) = sellerTaxId) And _
        (CStr(ws.Cells(targetRow + 1, 14).Value2) = buyerTaxId) And _
        (ws.Cells(targetRow, 8).NumberFormat = "@") And _
        (ws.Cells(targetRow, 14).NumberFormat = "@") And _
        (CLng(ws.Cells(targetRow, 16).Value2) = 1) And _
        (CLng(ws.Cells(targetRow + 1, 16).Value2) = 2) And _
        (CStr(ws.Cells(targetRow, 19).Value2) = "ITEM FIXTURE 1") And _
        (CStr(ws.Cells(targetRow + 1, 19).Value2) = "ITEM FIXTURE 2") And _
        (CDbl(ws.Cells(targetRow, 28).Value2) = 20) And _
        (CDbl(ws.Cells(targetRow, 29).Value2) = 220) And _
        (CDbl(ws.Cells(targetRow, 30).Value2) = 20) And _
        (CStr(ws.Cells(targetRow, 31).Value2) = "Tien thue ok") And _
        (CStr(ws.Cells(targetRow, 33).Value2) = "Khong co link tra cuu")
    Exit Function
Failed:
    RunDetailInvoiceFixture = False
End Function

Public Function CodexTestPurchaseDetail() As Boolean
    CodexTestPurchaseDetail = RunDetailInvoiceFixture("ChiTietHD_Mua", "mua", 200, "0000000001", "0000000002")
End Function

Public Function CodexTestSalesDetail() As Boolean
    CodexTestSalesDetail = RunDetailInvoiceFixture("ChiTietHD_Ban", "ban", 200, "0000000003", "0000000004")
End Function

Private Function RunSummaryDateFixture(ByVal sheetName As String, ByVal invoiceType As Long, _
    ByVal targetRow As Long) As Boolean
    On Error GoTo Failed
    Dim ws As Worksheet
    Dim sampleJson As String
    Dim writeRow As Long, sequenceNumber As Long

    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Range("A" & targetRow & ":BL" & targetRow).ClearContents
    ReDim arrTrangThai(1 To 1, 1 To 1)
    ReDim arrKQKTHoaDon(1 To 2, 1 To 1)
    arrTrangThai(1, 1) = "STATUS-FIXTURE"
    arrKQKTHoaDon(2, 1) = "CHECK-FIXTURE"
    Set dicLink = CreateObject("Scripting.Dictionary")
    Set dicTenCotTC = CreateObject("Scripting.Dictionary")

    sampleJson = "{""datas"":[{""tlhdon"":1,""khmshdon"":""1"",""khhdon"":""C26TST""," & _
        """shdon"":42,""tdlap"":""2026-09-12T00:00:00"",""dvtte"":""VND"",""tgia"":1," & _
        """nbten"":""SELLER-FIXTURE"",""nbmst"":""0000000001"",""nbdchi"":""FIXTURE""," & _
        """nky"":""2026-09-12T00:00:00"",""mhdon"":""MCCQT-FIXTURE""," & _
        """ncma"":""2026-09-12T00:00:00"",""nmten"":""BUYER-FIXTURE"",""nmmst"":""0000000002""," & _
        """nmdchi"":""FIXTURE"",""tgtcthue"":0,""tgtkcthue"":0,""tgtthue"":0," & _
        """ttcktmai"":0,""tgtkhac"":0,""tgtttbso"":0,""tgtttbchu"":"""",""gchu"":""""," & _
        """msttcgp"":"""",""thttltsuat"":[],""thttlphi"":[],""tthai"":0,""ttxly"":0," & _
        """cttkhac"":[],""ttkhac"":[]}]}"

    writeRow = targetRow
    sequenceNumber = 1
    ghiExcel_TongHop sampleJson, writeRow, invoiceType, sequenceNumber
    RunSummaryDateFixture = _
        (CDbl(ws.Cells(targetRow, 6).Value2) = CDbl(DateSerial(2026, 9, 12))) And _
        (ws.Cells(targetRow, 6).NumberFormat = "dd/mm/yyyy") And _
        (CDbl(ws.Cells(targetRow, 12).Value2) = CDbl(DateSerial(2026, 9, 12))) And _
        (CDbl(ws.Cells(targetRow, 14).Value2) = CDbl(DateSerial(2026, 9, 12)))
    Exit Function
Failed:
    RunSummaryDateFixture = False
End Function

Public Function CodexTestPurchaseSummaryDate() As Boolean
    CodexTestPurchaseSummaryDate = RunSummaryDateFixture("TongHopHD_Mua", 1, 200)
End Function

Public Function CodexTestSalesSummaryDate() As Boolean
    CodexTestSalesSummaryDate = RunSummaryDateFixture("TongHopHD_Ban", 2, 200)
End Function
'@)

    $purchaseOk = [bool]$excel.Run("'$($workbook.Name)'!CodexTestPurchaseDetail")
    $salesOk = [bool]$excel.Run("'$($workbook.Name)'!CodexTestSalesDetail")
    $purchaseSummaryOk = [bool]$excel.Run("'$($workbook.Name)'!CodexTestPurchaseSummaryDate")
    $salesSummaryOk = [bool]$excel.Run("'$($workbook.Name)'!CodexTestSalesSummaryDate")

    $fixtureResults = @()
    foreach ($fixture in @(
        @{ Sheet = 'ChiTietHD_Mua'; Seller = '0000000001'; Buyer = '0000000002'; Pass = $purchaseOk },
        @{ Sheet = 'ChiTietHD_Ban'; Seller = '0000000003'; Buyer = '0000000004'; Pass = $salesOk }
    )) {
        $sheet = $workbook.Worksheets.Item($fixture.Sheet)
        $fixtureResults += [ordered]@{
            Sheet = $fixture.Sheet
            Pass = $fixture.Pass
            TemplateCode = [string]$sheet.Cells.Item(200, 1).Value2
            Series = [string]$sheet.Cells.Item(200, 2).Value2
            SellerTaxId = [string]$sheet.Cells.Item(200, 8).Value2
            BuyerTaxId = [string]$sheet.Cells.Item(200, 14).Value2
            SellerTaxIdRepeated = [string]$sheet.Cells.Item(201, 8).Value2
            BuyerTaxIdRepeated = [string]$sheet.Cells.Item(201, 14).Value2
            ExpectedSellerTaxId = $fixture.Seller
            ExpectedBuyerTaxId = $fixture.Buyer
        }
        Release-ComObject $sheet
    }

    $result = [ordered]@{
        Pass = ($layoutOk -and $summaryLayoutOk -and $purchaseOk -and $salesOk -and $purchaseSummaryOk -and $salesSummaryOk)
        Workbook = $BuiltWorkbook
        Layout = $layoutResults
        SummaryLayout = $summaryLayoutResults
        SummaryFixtures = [ordered]@{
            Purchase = $purchaseSummaryOk
            Sales = $salesSummaryOk
        }
        Fixtures = $fixtureResults
    }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $repositoryRoot 'tests\detail-invoice-runtime-result.json') -Encoding utf8

    $workbook.VBProject.VBComponents.Remove($testModule)
    Release-ComObject $testModule; $testModule = $null
    $workbook.Close($false); Release-ComObject $workbook; $workbook = $null
    if (-not $result.Pass) { throw 'Detail-invoice runtime fixture failed.' }
    Write-Host 'DETAIL-INVOICE RUNTIME TEST PASSED'
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $testModule; Release-ComObject $workbook; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
