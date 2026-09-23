# Rà soát lỗi giao diện và luồng tải hóa đơn

Ngày rà soát: 22/09/2026

## 1. Luồng nghiệp vụ hiện tại

- `frmTaiHoaDon.taiHoaDon_Total` duyệt từng khoảng ngày, gọi lần lượt API `query` và `sco-query`, phân trang theo trường `state`, ghi dữ liệu tổng hợp bằng `ghiExcel_TongHop`, rồi gom định danh hóa đơn để gọi API `detail` và ghi bằng `ghiExcel_ChiTiet`.
- `frmTaiHoaDon.taiXML_zip` đang tải `export-xml` trực tiếp bằng `MSXML2.XMLHTTP`, ghi `responseBody` thành ZIP rồi gọi `Unzip`. Luồng này tách rời hoàn toàn khỏi `ApiGet` và không có retry dùng chung.
- `modHTTPRequest.ApiGet` đang dùng `MSXML2.ServerXMLHTTP.6.0`, nhưng xử lý 401/403/429/500/504 theo từng nhánh riêng, hiện thông báo rồi trả chuỗi rỗng. Chưa có timeout tập trung, `Retry-After`, exponential backoff, jitter, hàng đợi cuối phiên hoặc trạng thái lỗi chi tiết cho caller.

## 2. Các nhánh làm lỗi request bị bỏ qua

- Trong `taiHoaDon_Total`, lỗi hoặc response rỗng của `query` đi thẳng tới `nextPage_sco`; lỗi parse JSON cũng bị `On Error GoTo nextPage_sco` bỏ qua.
- Lỗi hoặc response rỗng của `sco-query` đi thẳng tới `nextDatePeriod`; lỗi parse JSON cũng bị `On Error GoTo nextDatePeriod` bỏ qua.
- Lỗi detail đi tới `nextInvoice`; vì vậy hóa đơn lỗi không có cơ chế retry cuối phiên và không được upsert vào báo cáo có cấu trúc.
- `taiXML_zip` gặp 429 sẽ `GoTo EH_Exit`, dừng toàn bộ phần XML; 500 được coi là không có hồ sơ XML; các lỗi khác đi tới hóa đơn kế tiếp. Không có retry/backoff.
- `cmdTaiHoaDon_Click` và `taiHoaDon_Total` reset `StatusBar` ở đường đi bình thường, nhưng chưa có một cleanup handler bao phủ mọi lỗi.

## 3. Nguyên nhân date picker không xuất hiện

- Hai nút runtime được thêm vào `Me.Controls` theo tọa độ tuyệt đối trong khi hai textbox ngày nằm trong frame `fraChonNgay`.
- Frame có thể che control runtime theo z-order; tọa độ của nút cũng không cùng hệ tọa độ với textbox trong frame.
- Collection `mUiHandlers` đã giữ handler sống trong vòng đời form, nên nguyên nhân chính không phải mất reference event handler.
- Cách sửa: thêm hai nút vào `fraChonNgay.Controls`, lấy tọa độ tương đối trực tiếp từ `txtTuNgay`/`txtDenNgay`, gọi `ZOrder 0`, và tiếp tục giữ handler trong `mUiHandlers`.

## 4. Encoding

- Source export hiện chứa mojibake ở các literal mới như `S?n s�ng`, `�ang ki?m tra...`, caption của `frmDatePicker`.
- Script cũ đọc UTF-8 rồi ghi Windows-1258 trước khi import, trong khi một phần source đã bị hỏng từ lần export trước; chuyển encoding lần nữa không thể khôi phục ký tự đã mất.
- Biện pháp áp dụng cho code mới: source VBA chỉ dùng ASCII cho chuỗi giao diện và chuyển sang Unicode lúc chạy bằng `UniConvert(...)` hoặc `ChrW(...)`. Báo cáo Markdown/JSON trong Git giữ UTF-8.

## 5. Phạm vi nâng cấp

- Tạo transport retry dùng chung cho JSON/text và binary ZIP.
- Cho `ApiGet` delegate sang transport mới và công bố trạng thái lần gọi gần nhất.
- Chuyển `taiXML_zip` sang transport binary mới.
- Tạo sheet `BaoCao_LoiTaiHD` 17 cột, upsert theo khóa nghiệp vụ.
- Tạo hàng đợi retry cuối phiên không trùng; 401/403 đặt cờ dừng toàn cục.
- Hiển thị phần trăm, hành động hiện tại, retry/wait/queue và log ngắn không chứa token.

## 6. Giới hạn kiểm chứng tại thời điểm rà soát

- Excel COM hiện trả lỗi `0x80070520` (logon session không còn tồn tại), nên chưa đọc được CodeModule trực tiếp từ workbook trong phiên này.
- Kết luận encoding ở trên dựa trên source export và binary form hiện có. Phần CodeModule trực tiếp phải được kiểm tra lại khi Excel COM hoạt động trước khi tuyên bố PASS.

## 7. Cập nhật sau khi xử lý (22/09/2026)

Đã sửa và kiểm chứng lại:

- `build/Build-Excel.ps1` là UTF-8 không BOM nhưng chứa literal tiếng Việt, nên Windows PowerShell 5.1 đọc bằng code page ANSI 1252, byte `0x92` biến thành dấu nháy cong và script không parse được (`The string is missing the terminator`). Phần tạo sheet thủ công trong script đã bỏ; thay bằng gọi `EnsureGdtErrorReportSheet` nên tiêu đề, định dạng và freeze panes đều do VBA sở hữu, và script build trở lại ASCII thuần.
- Literal Telex trong code mới đã rà bằng cách tái tạo `UniConvert`; năm chuỗi hiển thị sai đã được sửa: hậu tố 6 của báo cáo (`Ký hiệu mâũ số`), `Số mucj`, `Pảse dữ liệu` (giữ `Parse` ngoài `UniConvert`), `giải nesn XML` và `Ddã dừng yêeu cầu mơí`.
- `FormatGdtErrorReport` nhận thêm freeze panes dạng best effort để sheet được tạo lúc chạy có cùng hình thức với sheet do build tạo.
- `build/Test-Build.ps1` coi `BaoCao_LoiTaiHD` là phần thêm có chủ đích, chặn mọi worksheet mới khác, và đọc lại 17 tiêu đề để so với literal Telex đã giải mã.
- `tests/Test-RetryUiStatic.ps1` bổ sung bốn mục không cần Excel: đối chiếu 17 tiêu đề với `tests/retry-ui-expected-text.json`, quét toàn bộ literal `UniConvert` theo danh sách chuỗi sai đã biết, kiểm tra script build parse được và không thiếu BOM, và kiểm tra cân bằng khối VBA trên 21 file nguồn.
- `build/Upgrade-FormUI.ps1` được thêm BOM UTF-8 vì cũng chứa literal tiếng Việt.

## 8. Đóng hai mục Excel (22/09/2026)

Build đã chạy được và tạo `dist/App_v6.5.0-retry.xlsm` với `tests/comparison-report.json` đạt `Status=PASS`: mục 13 (mở lại không repair) và mục 14 (bằng chứng runtime cho compile) đều đạt. Bốn nguyên nhân đã xử lý:

- COM Excel trên máy thiếu đăng ký interface (`HKLM\SOFTWARE\Classes\Interface\{000208D5-0000-0000-C000-000000000046}` và các IID khác của Excel đều không tồn tại), nên `New-Object -ComObject Excel.Application` trả `0x80040155` cho mọi lời gọi và `Quit()` cũng thất bại, để lại tiến trình `EXCEL.EXE` mồ côi. Build, test và smoke test giờ tạo Excel bằng `[Type]::GetTypeFromCLSID` + `Activator.CreateInstance` (`New-ExcelApplication`), đường IDispatch không cần cast sang interop interface nên vẫn chạy được. Office vẫn nên được repair để sửa hẳn đăng ký COM.
- Template mang sẵn `Zone.Identifier` từ lần tải trên diễn đàn, nên mọi file build ra đều mở ở Protected View và bị chặn macro. Build gọi `Unblock-File` cho file đầu ra.
- Class module chỉ có LF khi được import làm workbook sau khi save không compile được, khiến mọi macro (kể cả macro có sẵn) báo `Cannot run the macro ... all macros may be disabled`. Build chuẩn hoá CRLF cho mọi file trước khi import.
- Tiêu đề báo cáo không còn là literal trong PowerShell: `Build-Excel.ps1` decode Telex từ `modGdtErrorReport` bằng `Get-GdtErrorReportHeaders`, nên vẫn một nguồn sự thật và vẫn không cần bật macro (`AutomationSecurity = 3` giữ nguyên để `Workbook_Open` không chạy khi build).

Kiểm chứng ngày 22/09/2026: `build/Build-Excel.ps1 -Version 6.5.0-retry` đạt (18 check), `build/Test-Build.ps1 -RunUserFormInstantiation` instantiate thành công cả 5 UserForm trong tiến trình riêng, và `tests/Test-RetryUiStatic.ps1` đạt cả 18 mục.

## 9. Sửa lỗi `Subscript out of range` khi tải hóa đơn (22/09/2026)

Log người dùng gửi dừng ngay sau dòng `Đang chuẩn bị tải hóa đơn...` (mốc 3% trong `cmdTaiHoaDon_Click`) và chưa có dòng `Đang lấy danh sách kỳ ...` (mốc 5% trong `taiHoaDon_Total`), tức lỗi xảy ra ở đầu `taiHoaDon_Total`, không phải trong vòng lặp tải. Chuỗi `Lỗi: Subscript out of range` là do nhãn `DownloadFailed` ghi `Err.Description`.

Vùng đó chỉ có bốn khả năng sinh lỗi 9. Ba khả năng đã bị loại bằng dữ liệu thật: mở `dist/App_v6.5.0-retry.xlsm` và đọc worksheet `LinkTraCuu` cho thấy sheet còn nguyên với vùng dữ liệu 118 dòng, đủ các bảng tra `B2:E`, `I2:I8`, `O2:O11`, nên `LinkTraCuu`, `tenCotTraCuu` và hai lệnh đọc range không lỗi. Còn lại `For k = 1 To UBound(arrDate)`.

Nguyên nhân: `arrDate` chỉ được tạo trong `txtDenNgay_AfterUpdate` và bị `Erase arrDate` ở nhãn `cleanup:` của chính `taiHoaDon_Total`, nên mọi đường đi không chạy qua sự kiện đó đều để lại mảng rỗng trong khi hai ô ngày vẫn có giá trị:

- bấm **Tải** lần thứ hai trong cùng form (lần trước đã `Erase` ở cleanup),
- chọn **Đến ngày** trước **Từ ngày**: `CDate(Me.txtTuNgay)` trên ô rỗng làm `txtDenNgay_AfterUpdate` lỗi 13, nên `lietKeThoiGian` không bao giờ chạy; lần chọn **Từ ngày** sau đó chỉ gọi `txtTuNgay_AfterUpdate`,
- sửa **Từ ngày** sau khi đã có bảng kỳ: bảng kỳ cũ còn nguyên và không được tạo lại.

Bằng chứng runtime lấy bằng cách chèn tạm hai hàm vào chính module form của bản build 6.5.0 rồi đọc `UBound(arrDate)`:

- form vừa load: `bound=-1` (lỗi 9), `lietKeThoiGian` khi đó cũng lỗi 13 `Type mismatch` vì ô ngày còn rỗng,
- đặt hai ô ngày bằng code (mô phỏng đúng trạng thái người dùng, không có `AfterUpdate`): `bound=-1` — chính là trạng thái làm `UBound(arrDate)` báo `Subscript out of range` trong log.

Bản sửa trong `src/forms/frmTaiHoaDon.frm`:

- `taiHoaDon_Total` gọi `lietKeThoiGian()` trước khi duyệt kỳ; `cmdTaiHoaDon_Click` cũng gọi hàm này ngay sau các kiểm tra đầu vào nên bảng kỳ luôn khớp hai ô ngày đang hiển thị và lỗi được báo bằng `AbortInvalidDateRange` (dùng đúng chuỗi `Kiểm tra lại ngày tìm kiếm.` cũ) thay vì `Subscript out of range`.
- `lietKeThoiGian` thành `Function ... As Boolean`, kiểm tra hai ô ngày bằng chính `correctDate`, bọc `On Error GoTo InvalidDate` cho chuỗi ngày hỏng, `Erase arrDate` trước khi tạo và chỉ gán `arrDate` khi đã tạo xong.
- Bỏ giới hạn cứng `ReDim arrDate(60, 2)`: `arrDate` không còn được dùng làm mảng tạm, kích thước được tính theo số tháng thực tế nên khoảng hơn 5 năm không còn lỗi.
- `txtDenNgay_AfterUpdate` chỉ so sánh hai mốc ngày khi cả hai ô đã có giá trị, nên chọn **Đến ngày** trước không còn lỗi 13.

Kiểm chứng bản 6.5.1: `build/Build-Excel.ps1 -Version 6.5.1-retry` đạt 18/18 check và `build/Test-Build.ps1 -RunUserFormInstantiation` đạt (5/5 UserForm), `tests/Test-RetryUiStatic.ps1` đạt 20 mục (thêm mục 19 và 20 cho đúng hai lỗi này). Probe chèn vào bản build 6.5.1 cho kết quả: đặt ngày bằng code (`bound=-1`) rồi `lietKeThoiGian=True`, `bound=1`, kỳ 1 `01/09/2026 -> 30/09/2026`; sau `Erase` (mô phỏng cuối lần tải trước) `lietKeThoiGian` vẫn `True`; `Đến ngày < Từ ngày` và chuỗi ngày rác đều trả `False` không lỗi; khoảng `01/12/2020 -> 31/12/2030` tạo 121 kỳ và kỳ cuối là `01/12/2030 -> 31/12/2030`.
