[CmdletBinding()]
param([string]$RepositoryRoot = '')
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    if ($PSScriptRoot) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
    else { $RepositoryRoot = (Get-Location).Path }
}

$retry = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGdtRetry.bas')
$report = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGdtErrorReport.bas')
$form = Get-Content -Raw (Join-Path $RepositoryRoot 'src\forms\frmTaiHoaDon.frm')
$http = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modHTTPRequest.bas')
$writeExcel = Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modGhiExcel.bas')
$build = Get-Content -Raw (Join-Path $RepositoryRoot 'build\Build-Excel.ps1')

function Result([int]$id, [string]$name, [string]$status, [string]$evidence) {
    [pscustomobject]@{ id=$id; name=$name; status=$status; evidence=$evidence }
}
function Has([string]$text, [string]$pattern) { [bool][regex]::IsMatch($text, $pattern, 'IgnoreCase,Multiline') }

# --- UniConvert (Telex) reconstruction -------------------------------------
# The build and the runtime must agree on Vietnamese text without putting any
# non-ASCII literal in a script, so the converter is rebuilt from its own table.
function Get-VbaLogicalLines([string]$Text) {
    $logicalLines = @()
    $buffer = ''
    foreach ($line in @($Text -split "`r?`n")) {
        $trimmed = $line.Trim()
        if ($trimmed.EndsWith('_')) { $buffer += $trimmed.Substring(0, $trimmed.Length - 1) + ' '; continue }
        $logicalLines += ($buffer + $trimmed)
        $buffer = ''
    }
    return $logicalLines
}
function Get-UniConvertTables([string]$MsgBoxSource) {
    $telex = @()
    $codes = @()
    foreach ($line in @(Get-VbaLogicalLines $MsgBoxSource)) {
        if ($line -match '^Telex_Type = Array\((.*)\)$') { $telex = @([regex]::Matches($Matches[1], '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }) }
        elseif ($line -match '^CharCode = Array\((.*)\)$') { $codes = @([regex]::Matches($Matches[1], 'ChrW\((\d+)\)') | ForEach-Object { [int]$_.Groups[1].Value }) }
    }
    return New-Object PSObject -Property @{ Telex = $telex; Codes = $codes }
}
function Convert-TelexToUnicode([string]$Text, [object]$Tables) {
    $result = $Text
    for ($index = 0; $index -lt $Tables.Codes.Count -and $index -lt $Tables.Telex.Count; $index++) {
        $result = $result.Replace($Tables.Telex[$index], [string][char]$Tables.Codes[$index])
    }
    return $result
}
function Get-UniConvertLiterals([string]$Path, [object]$Tables) {
    $literals = @()
    $lineNumber = 0
    foreach ($line in @(Get-Content -LiteralPath $Path)) {
        $lineNumber++
        foreach ($match in @([regex]::Matches($line, 'UniConvert\("([^"]*)"\)'))) {
            $literals += [pscustomobject]@{
                File = (Split-Path -Leaf $Path)
                Line = $lineNumber
                Source = $match.Groups[1].Value
                Decoded = Convert-TelexToUnicode $match.Groups[1].Value $Tables
            }
        }
    }
    return $literals
}

# --- VBA block balance (stand-in evidence while Excel cannot compile) -------
function Remove-VbaComment([string]$Line) {
    $inString = $false
    for ($index = 0; $index -lt $Line.Length; $index++) {
        $character = $Line[$index]
        if ($character -eq '"') {
            if ($inString -and ($index + 1) -lt $Line.Length -and $Line[$index + 1] -eq '"') { $index++; continue }
            $inString = -not $inString
            continue
        }
        if (-not $inString -and $character -eq "'") { return $Line.Substring(0, $index) }
    }
    return $Line
}
function Get-VbaBlockBalance([string]$Path) {
    # Order matters only for the report.  A block opener is a statement, so a
    # one-line "If x Then: statement" is complete and must not open a block.
    $pairs = @(
        @{ Name = 'Procedure'; Open = '(?i)^(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]'; Close = '(?i)^End\s+(?:Sub|Function|Property)\s*$'; Multiline = $false },
        @{ Name = 'If'; Open = '(?i)^If\b.*\bThen\s*$'; Close = '(?i)^End\s+If\s*$'; Multiline = $true },
        @{ Name = 'With'; Open = '(?i)^With\b'; Close = '(?i)^End\s+With\s*$'; Multiline = $false },
        @{ Name = 'Select'; Open = '(?i)^Select\s+Case\b'; Close = '(?i)^End\s+Select\s*$'; Multiline = $false },
        @{ Name = 'For'; Open = '(?i)^For\b'; Close = '(?i)^Next\b'; Multiline = $false },
        @{ Name = 'Do'; Open = '(?i)^Do\b'; Close = '(?i)^Loop\b'; Multiline = $false },
        @{ Name = 'While'; Open = '(?i)^While\b'; Close = '(?i)^Wend\s*$'; Multiline = $false }
    )
    $balance = @()
    for ($index = 0; $index -lt $pairs.Count; $index++) { $balance += 0 }
    $withoutComments = @(Get-Content -LiteralPath $Path | ForEach-Object { Remove-VbaComment $_ })
    foreach ($logicalLine in @(Get-VbaLogicalLines ($withoutComments -join "`r`n"))) {
        # String contents are removed and statements separated so that one-line
        # forms such as "For i = 1 To 3: Next i" count as a balanced pair.
        $withoutStrings = [regex]::Replace($logicalLine, '"[^"]*"', '""')
        $fragments = @($withoutStrings -split ':' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 0 })
        for ($fragment = 0; $fragment -lt $fragments.Count; $fragment++) {
            $line = $fragments[$fragment]
            for ($index = 0; $index -lt $pairs.Count; $index++) {
                if ([regex]::IsMatch($line, $pairs[$index].Open)) {
                    if ($pairs[$index].Multiline -and $fragment -lt $fragments.Count - 1) { continue }
                    $balance[$index]++
                }
                elseif ([regex]::IsMatch($line, $pairs[$index].Close)) { $balance[$index]-- }
            }
        }
    }
    $result = @()
    for ($index = 0; $index -lt $pairs.Count; $index++) { $result += [pscustomobject]@{ Name = $pairs[$index].Name; Balance = $balance[$index] } }
    return $result
}

$expectedPath = Join-Path $PSScriptRoot 'retry-ui-expected-text.json'
$expected = Get-Content -LiteralPath $expectedPath -Raw -Encoding UTF8 | ConvertFrom-Json
$tables = Get-UniConvertTables (Get-Content -Raw (Join-Path $RepositoryRoot 'src\modules\modMsgboxTV.bas'))

$declaredHeaders = @()
foreach ($line in @($report -split "`r?`n")) {
    if ($line -notmatch '^\s*headers\(1,\s*(\d+)\)\s*=\s*(.+?)\s*$') { continue }
    $expression = $Matches[2]
    if ($expression -match '^UniConvert\("([^"]*)"\)$') { $declaredHeaders += (Convert-TelexToUnicode $Matches[1] $tables) }
    elseif ($expression -match '^"([^"]*)"$') { $declaredHeaders += $Matches[1] }
    else { $declaredHeaders += '<unsupported expression>' }
}
$headerDifferences = @(Compare-Object -ReferenceObject @($expected.reportHeaders) -DifferenceObject $declaredHeaders -SyncWindow 0)
$headerStatus = if ($tables.Codes.Count -eq 0) { 'FAIL' } elseif ($declaredHeaders.Count -eq $expected.reportHeaders.Count -and $headerDifferences.Count -eq 0) { 'PASS_STATIC' } else { 'FAIL' }
$headerEvidence = if ($headerStatus -eq 'PASS_STATIC') {
    'UniConvert duoc tai tao tu modMsgboxTV.bas; 17 tieu de khop retry-ui-expected-text.json'
} else {
    'Lech tieu de: ' + (@($headerDifferences | ForEach-Object { $_.InputObject + '(' + $_.SideIndicator + ')' }) -join '; ')
}

$sourceFiles = @(Get-ChildItem -Path (Join-Path $RepositoryRoot 'src') -Recurse -File | Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') -and $_.Name -ne 'modMsgboxTV.bas' })
$literals = @()
foreach ($file in $sourceFiles) { $literals += @(Get-UniConvertLiterals $file.FullName $tables) }
$offendingLiterals = @()
foreach ($literal in $literals) {
    foreach ($fragment in @($expected.forbiddenFragments)) {
        if ($literal.Decoded.Contains($fragment)) { $offendingLiterals += "$($literal.File):$($literal.Line) [$($literal.Source)]"; break }
    }
}
$telexStatus = if ($offendingLiterals.Count -eq 0) { 'PASS_STATIC' } else { 'FAIL' }
$telexEvidence = if ($telexStatus -eq 'PASS_STATIC') {
    'Da decode ' + $literals.Count + ' literal UniConvert; khong con chuoi sai da biet.'
} else {
    'Literal sai: ' + ($offendingLiterals -join '; ')
}

$buildScriptErrors = @()
$nonAsciiScripts = @()
foreach ($script in @(Get-ChildItem -Path (Join-Path $RepositoryRoot 'build\*.ps1'), (Join-Path $RepositoryRoot 'build\*.psm1'))) {
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$null, [ref]$parseErrors)
    if (@($parseErrors).Count -gt 0) { $buildScriptErrors += "$($script.Name): $($parseErrors[0].Message)" }
    $bytes = [IO.File]::ReadAllBytes($script.FullName)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if (-not $hasBom -and @($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) { $nonAsciiScripts += $script.Name }
}
$scriptStatus = if ($buildScriptErrors.Count -eq 0 -and $nonAsciiScripts.Count -eq 0) { 'PASS_STATIC' } else { 'FAIL' }
$scriptEvidence = if ($scriptStatus -eq 'PASS_STATIC') {
    'Moi script build parse duoc va khong co literal non-ASCII thieu BOM.'
} else {
    'Loi parse: ' + ($buildScriptErrors -join '; ') + ' | Thieu BOM: ' + ($nonAsciiScripts -join ', ')
}

$vbaFiles = @(Get-ChildItem -Path (Join-Path $RepositoryRoot 'src') -Recurse -File | Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') })
$imbalancedFiles = @()
foreach ($file in $vbaFiles) {
    $unbalanced = @(Get-VbaBlockBalance $file.FullName | Where-Object { $_.Balance -ne 0 })
    if ($unbalanced.Count -gt 0) { $imbalancedFiles += "$($file.Name): " + (@($unbalanced | ForEach-Object { "$($_.Name)=$($_.Balance)" }) -join ',') }
}
$balanceStatus = if ($imbalancedFiles.Count -eq 0) { 'PASS_STATIC' } else { 'FAIL' }
$balanceEvidence = if ($balanceStatus -eq 'PASS_STATIC') {
    'Khoi Sub/Function/If/With/Select/For/Do can bang trong ' + $vbaFiles.Count + ' file nguon.'
} else {
    'File lech khoi: ' + ($imbalancedFiles -join '; ')
}

# The build report is the Excel-side evidence.  It is produced by
# build/Build-Excel.ps1 on a machine where Excel COM works.
$comparisonPath = Join-Path $PSScriptRoot 'comparison-report.json'
$comparison = $null
if (Test-Path -LiteralPath $comparisonPath) { $comparison = Get-Content -LiteralPath $comparisonPath -Raw | ConvertFrom-Json }
$reopenStatus = 'NOT_VERIFIED'
$reopenEvidence = 'Chua co tests/comparison-report.json. Chay build/Build-Excel.ps1 de lay bang chung mo lai workbook.'
if ($null -ne $comparison -and $comparison.ReopenVerified) {
    $failedChecks = @($comparison.Checks | Where-Object { -not $_.Pass })
    $reopenStatus = if ($failedChecks.Count -eq 0) { 'PASS_RUNTIME' } else { 'FAIL' }
    $reopenEvidence = 'Build da mo lai workbook trong Excel COM instance moi: ' + $comparison.Status + ', ' + $comparison.Checks.Count + ' check, ' + $failedChecks.Count + ' check loi, file ' + [IO.Path]::GetFileName([string]$comparison.BuiltWorkbook) + '.'
}
$compileStatus = 'NOT_VERIFIED'
$compileEvidence = 'Excel khong co API compile chinh thuc va chua co bang chung runtime. Chay build/Test-Build.ps1 -RunUserFormInstantiation.'
if ($null -ne $comparison -and $comparison.UserFormSmokeTest.Attempted) {
    $formResults = @($comparison.UserFormSmokeTest.Results)
    $formFailures = @($formResults | Where-Object { -not $_.Pass })
    $compileStatus = if ($formFailures.Count -eq 0) { 'PASS_RUNTIME' } else { 'FAIL' }
    $compileEvidence = if ($formFailures.Count -eq 0) {
        'Excel khong co API compile; bang chung runtime: ' + $formResults.Count + ' UserForm instantiate thanh cong trong tien trinh rieng (' + ((@($formResults | ForEach-Object { $_.Name })) -join ', ') + ').'
    } else {
        'UserForm instantiate loi: ' + ((@($formFailures | ForEach-Object { $_.Name })) -join ', ')
    }
}

$relatedRuntimePath = Join-Path $PSScriptRoot 'related-runtime-result.json'
$relatedRuntime = $null
if (Test-Path -LiteralPath $relatedRuntimePath) {
    $relatedRuntime = Get-Content -LiteralPath $relatedRuntimePath -Raw | ConvertFrom-Json
}
$relatedRuntimeStatus = if ($null -ne $relatedRuntime -and $relatedRuntime.Pass -and $relatedRuntime.ResolvedErrorRemoved -and $relatedRuntime.RelationErrorsVisible -and $relatedRuntime.LegacyRelatedColumnRemoved -and $relatedRuntime.DetailFillDownWithoutClipboard -and $relatedRuntime.FirstLineIsNewest -and $relatedRuntime.LastLineIsCurrent -and $relatedRuntime.EmptyRelatedHandled -and $relatedRuntime.StatusSixSkipsRelative -and $relatedRuntime.QueryNoticeLineCount -eq 1 -and $relatedRuntime.ScoNoticeLineCount -eq 3) { 'PASS_RUNTIME' } else { 'NOT_VERIFIED' }
$relatedRuntimeEvidence = if ($relatedRuntimeStatus -eq 'PASS_RUNTIME') {
    'Mau 18510 tao 6 dong; loi API hien thi dung cot; cot cu da bo; FillDown chi tiet khong dung clipboard.'
} else {
    'Chay build/Test-RelatedInvoiceRuntime.ps1 de kiem tra chuoi mau 18510.'
}

$results = @(
    (Result 1 'Thanh cong o lan dau' $(if(Has $retry 'IsSuccessfulGdtStatus\(result\.StatusCode\) Then'){'PASS_STATIC'}else{'FAIL'}) 'Nhanh 2xx tra Success ngay.'),
    (Result 2 '429 co Retry-After' $(if((Has $retry 'Case 0, 429, 500, 502, 503, 504') -and (Has $retry 'getResponseHeader\("Retry-After"\)')){'PASS_STATIC'}else{'FAIL'}) '429 retryable va uu tien header so.'),
    (Result 3 '429 khong co header' $(if(Has $retry 'GDT_RETRY_BASE_SECONDS \* \(2 \^ \(failedAttempt - 1\)\)'){'PASS_STATIC'}else{'FAIL'}) 'Backoff 2^n va jitter/cap.'),
    (Result 4 '500 roi thanh cong' $(if((Has $retry 'Case 0, 429, 500, 502, 503, 504') -and (Has $retry 'For attempt = 1 To maxAttempts')){'PASS_STATIC'}else{'FAIL'}) '500 nam trong vong retry.'),
    (Result 5 'Timeout roi thanh cong' $(if((Has $retry 'result\.StatusCode = 0') -and (Has $retry 'GoTo ContinueAttempt')){'PASS_STATIC'}else{'FAIL'}) 'Transport error/status 0 quay lai vong retry.'),
    (Result 6 'Het luot dua vao queue' $(if((Has $retry 'ShouldQueueGdtFinalRetry') -and (Has $form 'QueueFinalRetry')){'PASS_STATIC'}else{'FAIL'}) 'Exhausted retry duoc danh dau va queue khong trung.'),
    (Result 7 'Retry cuoi thanh cong' $(if((Has $form 'ProcessQueued(List|Detail|Xml|Relation)Retries') -and (Has $form 'MarkGdtRetrySuccess')){'PASS_STATIC'}else{'FAIL'}) 'List/detail/XML/relative/related deu co nhanh success.'),
    (Result 8 'Retry cuoi that bai' $(if(Has $form 'queued\.Attempts \+ requestResult\.Attempts[\s\S]*Khoong tari dduwowjc'){'PASS_STATIC'}else{'FAIL'}) 'That bai cuoi upsert Khong tai duoc.'),
    (Result 9 '401/403 dung request' $(if((Has $retry 'StatusCode = 401 Or result\.StatusCode = 403') -and (Has $http 'If GdtAuthenticationFailed Then')){'PASS_STATIC'}else{'FAIL'}) 'Co auth toan phien chan request moi.'),
    (Result 10 'Upsert khong trung' $(if((Has $report 'BuildGdtErrorKey') -and (Has $report 'FindGdtErrorRow')){'PASS_STATIC'}else{'FAIL'}) 'Tim khoa truoc khi ghi.'),
    (Result 11 'Bao cao mua/ban' $(if((Has $form 'Mua vafo') -and (Has $form 'Basn ra')){'PASS_STATIC'}else{'FAIL'}) 'Direction dung Mua vao/Ban ra.'),
    (Result 12 'Context query/sco-query' $(if((Has $form '"query"') -and (Has $form '"sco-query"') -and (Has $form 'ApiSource')){'PASS_STATIC'}else{'FAIL'}) 'Queue giu nguon API va dinh danh.'),
    (Result 19 'Bang ky thoi gian tao lai luc tai' $(if((Has $form 'Function lietKeThoiGian\(\) As Boolean') -and (Has $form 'Sub taiHoaDon_Total[\s\S]*?If Not lietKeThoiGian\(\) Then[\s\S]*?For k = 1 To UBound\(arrDate\)') -and (Has $form 'On Error GoTo InvalidDate')){'PASS_STATIC'}else{'FAIL'}) 'taiHoaDon_Total tao lai bang ky thoi gian truoc vong lap; arrDate khong con phu thuoc AfterUpdate cua o ngay.'),
    (Result 20 'O ngay trong khong lam loi form' $(if(Has $form 'If Len\(Trim\(Me\.txtTuNgay\)\) = 0 Or Len\(Trim\(Me\.txtDenNgay\)\) = 0 Then Exit Sub[\s\S]*?CDate\(Me\.txtDenNgay\) < CDate\(Me\.txtTuNgay\)'){'PASS_STATIC'}else{'FAIL'}) 'txtDenNgay_AfterUpdate chi so sanh ngay khi ca hai o da co gia tri.'),
    (Result 21 'Loi chi tiet khong truy cap arrDate(k)' $(if((Has $form 'errMsg = errMsg & UniConvert\("Looxi tari hoas ddown chi tieest: "\) & arrHDChiTiet\(j, 1\)') -and (-not (Has $form 'Looxi tari hoas ddown chi tieest ngafy:.*arrDate\(k'))){'PASS_STATIC'}else{'FAIL'}) 'Loi tai chi tiet ghi nhan ky hieu, so va ngay hoa don tu arrHDChiTiet(j), khong truy cap arrDate(k).'),
    (Result 22 'Tu dong mo rong buffer hoa don' $(if((Has $form 'EnsureInvoiceBufferCapacity') -and (Has $form 'GetSleepDelayMs') -and (Has $form 'Dim row As Long, row_ct As Long, stt As Long, n As Long')){'PASS_STATIC'}else{'FAIL'}) 'Co ham EnsureInvoiceBufferCapacity mo rong buffer dong, GetSleepDelayMs an toan va dat lai n luc tai.'),
    (Result 23 'Tien do theo so hoa don thuc te' $(if((Has $form 'mInvoiceWorkTotal = \(invoiceCount \* workKinds\) \+ relatedCallCount') -and (Has $form '10 \+ CLng\(\(mInvoiceWorkDone / mInvoiceWorkTotal\) \* 88\)') -and (-not (Has $form '30 \+ CLng'))){'PASS_STATIC'}else{'FAIL'}) 'Danh sach chiem 0-10%; 88% tiep theo gom ca relative, related, chi tiet va XML.'),
    (Result 24 'Goi API hoa don lien quan' $(if((Has $form 'If invoiceStatus = 6 Then[\s\S]*?WriteNoRelativeInvoiceData[\s\S]*?Else[\s\S]*?"relative"[\s\S]*?"related"') -and (Has $form 'apiSource = IIf\(invoiceBuffer\(invoiceIndex, 4\) = 1, "query", "sco-query"\)')){'PASS_STATIC'}else{'FAIL'}) 'Trang thai 2-5 goi relative va related; trang thai 6 chi goi related.'),
    (Result 25 'Header API lien quan' $(if((Has $retry 'setRequestHeader "Action", actionHeader') -and (Has $retry 'setRequestHeader "End-Point", "/tra-cuu/tra-cuu-hoa-don"') -and (Has $retry 'setRequestHeader "Authorization", "Bearer " & bearerToken')){'PASS_STATIC'}else{'FAIL'}) 'Request dung token phien hien tai, action, end-point va khong luu cookie vao source.'),
    (Result 26 'Bao cao chuoi lien quan' $(if((Has $writeExcel 'WriteRelativeInvoiceData') -and (Has $writeExcel 'Hosa ddown cos lieen quan') -and (Has $writeExcel 'Hosa ddown ddang tra cuwsu') -and (Has $writeExcel 'WriteRelatedInformationResponse')){'PASS_STATIC'}else{'FAIL'}) 'Ket qua relative thanh chuoi nhieu dong; related ghi rieng va chap nhan response rong.'),
    (Result 27 'Chuoi mau 18510' $relatedRuntimeStatus $relatedRuntimeEvidence),
    (Result 28 'Dinh dang thong bao related' $(if((Has $writeExcel 'mtthdtbssrs') -and (Has $writeExcel 'hdtbssrses') -and (Has $writeExcel 'kqtnhan') -and (Has $writeExcel 'Cow quan thuees khoong tieesp nhaajn')){'PASS_STATIC'}else{'FAIL'}) 'Parse ca response query va sco-query thanh cau thong bao tieng Viet.'),
    (Result 33 'Loi API lien quan hien thi tai dong hoa don' $(if((Has $writeExcel 'Public Sub WriteRelationRequestError') -and (Has $form 'WriteRelationRequestError queued\.ResponseKind[\s\S]*queued\.Attempts \+ requestResult\.Attempts')){'PASS_STATIC'}else{'FAIL'}) 'Relative ghi loi vao Chuoi hoa don lien quan; related ghi loi vao Thong tin lien quan sau lan thu cuoi.'),
    (Result 34 'Bo cot Co HD lien quan' $(if((Has $writeExcel 'Cells\(2, 64\)\.Value = UniConvert\("Thoong tin lieen quan"\)') -and (-not (Has $writeExcel 'Cos HDD lieen quan|tthdclquan|Cells\(targetRow, 65\)')) -and (Has $build 'Columns\.Item\(65\)\.Clear\(\)')){'PASS_STATIC'}else{'FAIL'}) 'Thong tin lien quan chuyen sang cot 64; build xoa cot cu va khong con du lieu Co HD lien quan.'),
    (Result 35 'Khong chiem dung clipboard khi ghi chi tiet' $(if((Has $writeExcel 'Set fillRange = ws\.Range\("A" & frow & ":N" & lrow\)[\s\S]*fillRange\.FillDown') -and (-not (Has $writeExcel 'sourceRange\.Copy|Application\.CutCopyMode'))){'PASS_STATIC'}else{'FAIL'}) 'Du lieu chung duoc FillDown noi bo, khong Copy va khong xoa clipboard Windows.'),
    (Result 29 'Ngay ISO khong phu thuoc locale' $(if(Has (Get-Content -LiteralPath (Join-Path $RepositoryRoot 'src\modules\modParseIso.bas') -Raw) 'UTCToLocalTime = DateSerial\(outsys\.wYear, outsys\.wMonth, outsys\.wDay\)'){'PASS_STATIC'}else{'FAIL'}) 'Tao ngay bang DateSerial/TimeSerial, khong ghep chuoi roi CDate theo locale.'),
    (Result 30 'Chi giu loi chua xu ly' $(if((Has $report 'ws\.Rows\(targetRow\)\.Delete Shift:=xlUp') -and (Has $report 'RenumberGdtErrorRows ws') -and (Has $form 'MarkGdtRetrySuccess CurrentDirectionName\(\)')){'PASS_STATIC'}else{'FAIL'}) 'Tai thanh cong xoa dong loi cung khoa va danh lai STT; ap dung ca lan dau va retry.'),
    (Result 31 'Log chi tiet phan trang danh sach' $(if((Has $form 'ReportListPageStart') -and (Has $form 'ReportListPageComplete') -and (Has $form 'queryPageNumber') -and (Has $form 'scoPageNumber') -and (Has $form 'LastGdtStatus') -and (Has $form 'TryRegisterNextListState')){'PASS_STATIC'}else{'FAIL'}) 'Log neu ky, nguon query/sco-query, trang, HTTP, so hoa don, tong luy ke, thoi gian va chan state lap.'),
    (Result 32 'Tam dung tiep tuc va dung an toan' $(if((Has $form 'cmdPauseResume') -and (Has $form 'cmdStopDownload') -and (Has $form 'TogglePauseDownload') -and (Has $form 'RequestStopDownload') -and (Has $retry 'Public GdtStopRequested As Boolean') -and (Has $retry 'Public GdtPauseRequested As Boolean') -and (Has $retry 'WaitForGdtControl')){'PASS_STATIC'}else{'FAIL'}) 'Hai nut runtime dieu khien co-operative; cac vong request va thoi gian cho deu kiem tra pause/stop.'),
    (Result 15 'Tieu de bao cao dung tieng Viet' $headerStatus $headerEvidence),
    (Result 16 'Literal Telex khong con sai' $telexStatus $telexEvidence),
    (Result 17 'Script build parse duoc' $scriptStatus $scriptEvidence),
    (Result 18 'Khoi VBA can bang' $balanceStatus $balanceEvidence),
    (Result 13 'Mo lai khong repair' $reopenStatus $reopenEvidence),
    (Result 14 'Compile VBAProject' $compileStatus $compileEvidence)
)

$jsonPath = Join-Path $PSScriptRoot 'retry-ui-test-results.json'
$results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding utf8

$lines = @('# Bao cao kiem tra retry/UI','',('Ngay chay: ' + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')),'', '| # | Tinh huong | Ket qua | Bang chung |','|---:|---|---|---|')
foreach($item in $results) { $lines += "| $($item.id) | $($item.name) | $($item.status) | $($item.evidence) |" }
$lines += @('','Luu y: `PASS_STATIC` chi xac nhan nhanh code va cau hinh hien dien; `PASS_RUNTIME` lay tu tests/comparison-report.json do build/Build-Excel.ps1 tao ra. Muc 13 va 14 la bang chung Excel; muc 17 va 18 la bang chung tinh khong can Excel.')
$lines | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'retry-ui-report.md') -Encoding utf8

if ($results.status -like 'FAIL*') { exit 1 }
