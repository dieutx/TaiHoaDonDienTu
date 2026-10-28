# Phát hành v6.7.8 — 28/10/2026

- Tra cứu máy tính tiền chia kỳ tối đa 10 ngày; khi API báo giới hạn khác, đọc số ngày và tự chia nhỏ lại kỳ bị từ chối. Giữ đúng nguồn đã chọn khi bắt đầu phiên tải và phân trang độc lập từng kỳ.
- Bổ sung CAPTCHA tra cứu SCO theo HAR và kiểm chứng live, dùng bộ đọc SVG hiện có. CAPTCHA không lưu trong URL báo lỗi hoặc hàng đợi retry.
- Thêm lựa chọn PDF khi tải hóa đơn, ưu tiên dựng từ XML gốc. Nếu không có XML, dùng chi tiết API và tổng hợp đúng cùng hóa đơn để dựng PDF khi đủ dữ liệu; PDF ghi rõ nguồn API và không có XML gốc.
- Thêm **XML -> PDF** cho thư mục XML có sẵn. Mẫu HTML/CSS nhúng trong workbook, in A4 bằng Edge/Chrome, hỗ trợ nhiều trang và lặp đầu bảng. XML/ZIP đã tải vẫn giữ để đối chiếu.
- Kiểm tra danh tính và dữ liệu trước khi dựng. Thiếu hàng hóa, tổng tiền hoặc sai hóa đơn thì báo lỗi và giữ PDF cũ; không tạo XML thay thế, không giả chữ ký số.

## Kiểm thử

Release gate chạy toàn bộ kiểm thử nền, cộng kiểm thử ngày/PDF và PDF không có XML. Fixture chỉ dùng dữ liệu giả; live chỉ chạy khi chủ tài khoản cho phép, xóa dữ liệu thử sau khi kết thúc. Chi tiết API và phạm vi: [ngày, CAPTCHA và PDF](DATE_LIMIT_AND_XML_PDF_2026-10-28.md).

```powershell
.\build\Test-Release.ps1 -Version '6.7.8' -OutputPath '.\dist\release-v6.7.8\TaiHoaDonDienTu_v6.7.8.xlsm'
```

Bản phát hành đã chạy `RELEASE CHECKS PASSED`, exit 0: cấu trúc/source, khởi tạo UserForm, workbook công khai, chi tiết/thông tin liên quan, 35 ca ngày, review/compatibility/core/resume, đồng bộ downloader, bộ lọc nguồn, Retry-After, đồng hồ/giao diện, JSON song song, ngày/PDF, PDF không có XML và cập nhật phiên bản đều PASS. `Test-DatePdfRegression.ps1` chạy lại 11 nhóm kiểm thử nền và ngày/PDF cũng PASS.

## Sử dụng

Chọn **PDF** trên form tải, hoặc **XML -> PDF** cho thư mục có sẵn. Cần Microsoft Edge hoặc Google Chrome. PDF là bản thể hiện dữ liệu; chữ ký số chưa được xác minh. Nếu XML không có và API chưa đủ dữ liệu, xem bước `PDF` trong `BaoCao_LoiTaiHD`.

## Bản tải

File: `TaiHoaDonDienTu_v6.7.8.xlsm`, 463428 byte. SHA256: `3D302381833DA8FC7FA50969203A83EC22A7F6279B9D3A062F2B619C02F526C4`.

[TaiHoaDonDienTu v6.7.8](https://github.com/dieutx/TaiHoaDonDienTu/releases/tag/v6.7.8).
