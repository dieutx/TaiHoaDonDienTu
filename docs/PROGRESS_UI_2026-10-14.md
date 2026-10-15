# Đồng hồ xử lý, log tổng số hóa đơn và thanh thời gian — 14/10/2026

Ba vấn đề giao diện được báo từ ảnh form khi đang tải:

1. **“Thời gian xử lý” đứng im.** Nhãn `lblStatus` chỉ được gán một lần ở cuối phiên tải (và ở đầu pha XML), nên trong suốt quá trình tải nó giữ nguyên giá trị cũ, không nhúc nhích theo thời gian thực.
2. **Log không cho biết tổng số hóa đơn cần xử lý.** Log chỉ có dòng tiến trình/HTTP; dòng “Đã lấy danh sách N hóa đơn” bị đẩy khỏi khung khi log dài, và các dòng chi tiết không kèm mẫu số.
3. **Thanh “Thời gian xử lý” nhìn chưa cân đối.** `lblStatus` rộng 318 điểm đặt ở `Left = 197,3` trong khi thanh tiến trình chạy từ 15 đến 767; mép trái 182,3 điểm so với mép phải 246,7 điểm nên lệch hẳn về bên trái.

Source đang phát triển, chưa phát hành phiên bản mới (`VERSION` vẫn là 6.7.6). Bằng chứng trước/sau nằm trong `dist/progress-ui-20261014/`.

## Thay đổi

### Đồng hồ “Thời gian xử lý” chạy liên tục

- Thêm `Public Sub TickElapsedTime` trong `frmTaiHoaDon`. Mỗi lần gọi, nếu đang tải, nhãn được cập nhật thành `Thời gian xử lý: hh:mm:ss` tính từ `mRunStarted`; chỉ gán khi chuỗi đổi nên không nhấp nháy.
- `GdtPumpWait` (`modGdtXmlScheduler.bas`) gọi `frmTaiHoaDon.TickElapsedTime` sau `DoEvents`, nên mọi vòng chờ — kể cả chờ `Retry-After`, cooldown 429 và giãn cách giữa hai request — đều làm đồng hồ nhảy.
- `SetProgress` cũng gọi `TickElapsedTime`, phòng trường hợp có cập nhật tiến độ mà không qua vòng chờ.
- `ResetProgressUI` đặt lại `mElapsedShown`, xóa `mInvoiceCount` và đưa nhãn về `Thời gian xử lý: 00:00:00`. Hàm này chạy ở đầu mỗi phiên tải, nên khi người dùng đổi tham số rồi bấm **Tải hóa đơn**, đồng hồ bắt đầu lại từ 0 thay vì giữ giá trị cũ.
- Dòng cũ ghi đè nhãn ở pha XML được chuyển thành `AppendSimpleLog` để không xóa đồng hồ đang chạy.

### Log hiển thị tổng số hóa đơn cần xử lý

- Thêm biến form `mInvoiceCount`, được gán trong `BeginInvoiceProgress` đúng lúc danh sách hóa đơn đã sẵn sàng.
- Ghi thẳng một dòng log: `Tổng số hóa đơn cần xử lý: N`.
- Mỗi dòng log tiến trình chi tiết giờ có mẫu số: `DETAIL 406/1234 | HTTP 200 | attempts=1 | 0,1s`.
- Thông báo trên thanh tiến độ cũng kèm mẫu số: `Đã xử lý chi tiết 1/3`. Khi resume bỏ qua hóa đơn đã xong, dòng `Đã có chi tiết n/N` cũng vậy.

### Căn lại thanh “Thời gian xử lý”

- Trong `EnsureRuntimeControls`, nhãn `lblStatus` được căn giữa theo đúng bề rộng thanh tiến trình và bật căn giữa chữ:
  ```vba
  Me.lblStatus.Left = Me.Controls("lblProgressBack").Left + (Me.Controls("lblProgressBack").Width - Me.lblStatus.Width) / 2
  Me.lblStatus.TextAlign = 2
  ```
- Không đổi tên control/component/sheet, không sửa `.frx`, không đổi cặp FRM/FRX.

## Kiểm thử

### Kiểm tra tĩnh

`tests/Test-RetryUiStatic.ps1` thêm ba mục:

- **36 — Đồng hồ thời gian xử lý chạy liên tục**: có `TickElapsedTime`, `mElapsedShown`, nhãn dựng từ `GdtElapsedText`, `GdtPumpWait` gọi `frmTaiHoaDon.TickElapsedTime`, và `mRunStarted` được đặt khi bắt đầu tải. `PASS_STATIC`.
- **37 — Log hiển thị tổng số hóa đơn**: `mInvoiceCount = invoiceCount`, dòng `Tổng số hóa đơn cần xử lý`, và dòng `DETAIL … "/" & mInvoiceCount`. `PASS_STATIC`.
- **38 — Thanh thời gian cần đối**: nhãn được căn giữa theo `lblProgressBack` và `TextAlign = 2`. `PASS_STATIC`.

### Kiểm tra runtime mới

`build/Test-ProgressClockRuntime.ps1` (đã thêm vào release gate) mở workbook đóng băng đối tượng, gắn tạm một hàm vào chính module `frmTaiHoaDon` để chạy đúng các thành viên thật rồi kiểm tra:

- Sau `ResetProgressUI` + `BeginInvoiceProgress 3`, nhãn đúng bằng `Thời gian xử lý: 00:00:00` (đồng hồ reset).
- `WaitGdtSeconds 3` **không** gọi `SetProgress`: chỉ vòng chờ `GdtPumpWait` cập nhật nhãn. Nhãn phải đổi thành `Thời gian xử lý: 00:00:02` → đồng hồ tự chạy.
- Log chứa `Tổng số hóa đơn cần xử lý: 3`.
- Thông báo tiến độ chứa `Đã xử lý chi tiết 1/3`.
- `lblStatus` cách đều hai mép thanh tiến trình (sai số ≤ 0,2 điểm) và `TextAlign = 2`.

Kết quả trong log gate: `PASS| Thời gian xử lý: 00:00:00| Thời gian xử lý: 00:00:02|Đã xử lý chi tiết 1/3`.

### Bố cục đo trên artifact cuối

`dist/fixes-20261014/Dump-AllControls.ps1` đọc toạ độ control sau `UserForm_Initialize` trên workbook build ra:

- `lblStatus` `L=232,0 W=318,0 R=550,0`; thanh tiến trình `lblProgressBack` `L=15,0 R=767,0`. Khoảng trống trái 217,0 điểm, phải 217,0 điểm — cân đối; trước sửa là 182,3 / 246,7.
- `dist/fixes-20261014/Inspect-Overlap.ps1`: **0** control chồng lấn.

### Release gate

`build/Test-Release.ps1 -Version 6.7.6` chạy trong `dist/progress-ui-20261014/verification-root` (bản archive của working tree, không ghi đè inventory/report của repository): **RELEASE CHECKS PASSED**, exit 0, 17 bước gồm bước `Test-ProgressClockRuntime` mới. Log: `dist/progress-ui-20261014/release-gate.log`.

### Artifact cuối

- `dist/progress-ui-20261014/verification-root/dist/TaiHoaDonDienTu_v6.7.6.xlsm`, 383.721 byte, `SHA256 = e38569539ba56647c07902da096bfef3e6579e7c9b11805e59d43c46d0737f7a`; hash không đổi sau mọi lần kiểm tra.
- Môi trường: Windows, Windows PowerShell 5.1.22621.6133, Excel COM 16.0 (64-bit). Excel 32-bit chưa được kiểm tra.
- Chưa gọi GDT thật; mọi bài runtime dùng workbook mở chỉ để đọc/đóng không lưu và máy chủ localhost giả lập.
