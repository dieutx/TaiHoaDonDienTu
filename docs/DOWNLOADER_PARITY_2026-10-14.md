# Phục hồi chi tiết sau XML 500/504 và giữ giãn cách cấu hình

Áp dụng hai thay đổi phù hợp với VBA từ vòng review
[hddt-downloader-windows PR #12](https://github.com/dieutx/hddt-downloader-windows/pull/12).
Source dựa trên branch `fix/core-data-resume-log`, commit `ee7e0b5` (PR #10);
các thay đổi chưa nằm trong release công khai 6.7.6.

## Phục hồi dữ liệu khi không tải được XML

- Sau các lượt thử XML, nếu HTTP status cuối là 500 hoặc 504, ứng dụng thử
  endpoint `/api/query/invoices/detail` hiện có bằng token của phiên chạy.
- Giữ nguyên dòng tổng hợp đã lấy từ danh sách. Detail được ghi vào
  `ChiTietHD_Mua` hoặc `ChiTietHD_Ban` theo chiều đang chạy, kể cả khi người
  dùng chỉ chọn XML và chưa chọn tải chi tiết.
- Phản hồi phải là object JSON, khớp MST người bán, mẫu số, ký hiệu và số hóa
  đơn. Parser/writer chi tiết hiện có kiểm tra cấu trúc và giữ rollback khi lỗi.
- Nếu resume đã xác minh đủ chi tiết, dùng phần đó, không tải/ghi thêm lần nữa.
- Không tạo XML, ZIP hoặc dữ liệu sheet `_XML` từ JSON. Công đoạn XML không
  được đánh dấu hoàn tất; lần chạy sau vẫn có thể thử tải XML gốc.
- `BaoCao_LoiTaiHD` giữ HTTP status, endpoint và số lượt XML đã thử; kết quả
  ghi `Recovered from API detail (no XML)` khi đã có dữ liệu chi tiết. Đây
  vẫn là một mục XML chưa hoàn tất, không được tính là tải đủ XML thành công.
- API detail lỗi thì báo nguyên nhân cùng lỗi XML, giữ dữ liệu cũ và cơ chế
  retry cuối phiên. HTTP status khác không kích hoạt fallback. Stop/auth
  failure chặn request phục hồi mới.

Fallback dùng `ExecuteGdtRequest` và khoảng nghỉ/retry hiện có. Request detail
phục hồi chạy tuần tự trên luồng điều khiển, có DoEvents/stop giữa các lượt;
không thêm worker ghi Excel hay đổi luồng đăng nhập. Các XML đã chạy có thể
chờ được thu kết quả trong lúc request detail đang chờ. Đây là lựa chọn patch
nhỏ; nếu tần suất fallback lớn cần benchmark trước khi đưa vào pipeline JSON.

## Giãn cách sau 429

`clsGdtXmlThrottle.RateLimited` áp trần điều tiết 5.000 ms trước, rồi giữ
giãn cách ít nhất bằng `BaseIntervalMs`. Trước sửa, mức cấu hình 5.001,
20.000 hoặc 60.000 ms lại bị giảm xuống 5.000 ms khi nhận 429. Mức mặc định
300 ms XML/600 ms JSON và số kết nối không đổi.

Các đối ứng khác đã có trong baseline VBA: khóa hóa đơn chứa người bán/mẫu,
chờ retry có DoEvents và cờ stop, metadata attempts/status, resume theo scope
và báo cáo lỗi có khóa riêng. Không copy lại PowerShell, WPF, OPC/sharedStrings.

## Kiểm thử

```powershell
.\build\Build-Excel.ps1 -Version '6.7.6' -OutputPath '.\dist\parity-test\TaiHoaDonDienTu_v6.7.6.xlsm'
.\build\Test-DownloaderParityRuntime.ps1 -BuiltWorkbook '.\dist\parity-test\TaiHoaDonDienTu_v6.7.6.xlsm'
```

Fixture mới có trong release gate. Chỉ dùng Excel COM, dữ liệu giả và server
HTTP localhost. Source/endpoint được instrument trong bộ nhớ, workbook đóng
không lưu và hash artifact phải giữ nguyên.

- Baseline: test throttle thất bại tại giãn cách 5.001 ms, trước khi sửa.
- Sau sửa: các mức 0/300/600/5.000/5.001/20.000/60.000 ms, rate-limit lặp
  và phục hồi đạt; khóa seller/template và stop trong backoff đạt.
- Fallback mua/bán, nguồn sco-query, XML 500/504, reuse detail không trùng,
  JSON hỏng, khóa không khớp, HTTP detail lỗi, 401, stop và status khác đạt.
- Ca end-to-end gọi scheduler XML thật với HTTP localhost 504 rồi chuyển sang
  detail; fixture giảm giới hạn retry còn 1 chỉ trong bộ nhớ để chạy nhanh.
- Kiểm tra request detail đúng endpoint/header/token synthetic, giữ metadata
  XML và không ghi sheet XML.

Chưa xác minh trên GDT thật hoặc Office 32-bit. Không tăng version, sửa
template hoặc publish workbook release trong thay đổi này.
