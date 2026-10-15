# Phát hành v6.7.7 — 15/10/2026

## Thay đổi

- Phản hồi JSON có HTTP 200 nhưng body là HTML/XML được nhận diện, thử lại và đưa vào retry cuối phiên nếu hết lượt. Báo cáo giữ HTTP thật và ghi rõ phản hồi không phải JSON. Không tự thay đổi luồng đăng nhập.
- Tất cả sheet tắt Wrap Text và căn giữa theo chiều dọc. Các writer tổng hợp, chi tiết, thông tin liên quan, báo cáo lỗi và sheet trạng thái tạo mới giữ cùng quy tắc. Giá trị, công thức, định dạng ngày/số và căn ngang được giữ nguyên.
- Tiến độ phân trang hiển thị số hóa đơn và số trang đã nhận khi chưa biết tổng; nhật ký lớn hơn, đọc được dòng chọn, giữ lựa chọn khi có log mới; ETA chi tiết bắt đầu từ pha chi tiết.
- Bộ lọc nguồn hóa đơn áp dụng cho mua/bán và xóa dữ liệu theo phạm vi nguồn đã xác định. Tải bổ sung/resume giữ dữ liệu và chỉ bỏ qua công đoạn đã xác minh.
- Fallback API detail khi XML lỗi 500/504; kiểm tra định danh trước ghi, chống ghi chi tiết hai lần trong phiên tải mới. Giữ lỗi XML nếu không có XML gốc.
- Sửa deadline Retry-After ở retry cuối phiên, tách lỗi parse theo endpoint/công đoạn và kiểm tra schema trước khi xóa dữ liệu.
- Giữ các bản sửa tương thích XML namespace, kiểu dữ liệu, đăng nhập và link tra cứu từ nhánh đang phát triển.

## Kiểm thử

Lệnh phát hành:

```powershell
.\build\Test-Release.ps1 -Version '6.7.7' -OutputPath '.\dist\release-v6.7.7-final-20261015\TaiHoaDonDienTu_v6.7.7.xlsm'
```

Runner chạy các tiến trình Excel COM bằng MTA để tránh tình trạng Windows PowerShell STA đã in PASS nhưng bị kẹt khi thoát. Fixture chỉ dùng dữ liệu giả và HTTP localhost, không dùng tài khoản hoặc hóa đơn thật.

Release gate hoàn tất với `RELEASE CHECKS PASSED`, exit 0: 18 bước đều PASS, gồm cấu trúc/source, khởi tạo sáu UserForm, public workbook, chi tiết/thông tin liên quan, 35 ca ngày, review/compatibility/core/resume, fallback XML, bộ lọc nguồn, Retry-After, đồng hồ/giao diện, JSON song song và cập nhật online/offline. Ca Retry-After 5 giây đo được khoảng cách 5,052 giây với cooldown fixture 1 giây. Cả 13 tình huống JSON song song và kiểm tra BOM/HTML → 400 đều PASS.

Kiểm tra bổ sung định dạng cố tình bật Wrap Text, đặt căn trên/dưới trước khi gọi writer tổng hợp/chi tiết/thông tin liên quan. Kiểm tra public workbook xét toàn bộ Cells của từng sheet. Fixture HTML kiểm tra cả BOM/khoảng trắng và HTML → HTTP 400 để không giữ nhầm cờ lỗi qua request sau.

Chưa xác minh tỷ lệ lỗi/tốc độ trên GDT thật hoặc đầy đủ Office 32-bit. Retry không bảo đảm thành công nếu phiên đăng nhập hoặc gateway tiếp tục trả HTML; người dùng cần đăng nhập lại khi phù hợp. Khi tắt Wrap Text, đọc nội dung dài trên thanh công thức hoặc mở rộng cột; dữ liệu không bị rút ngắn.

## Bản tải

File: `TaiHoaDonDienTu_v6.7.7.xlsm`, 421434 byte. SHA256: `A5F939FB1B1E7653C083B81319881B59B4B2255EEB343A2A808F27719FB50086`.

[TaiHoaDonDienTu v6.7.7](https://github.com/dieutx/TaiHoaDonDienTu/releases/tag/v6.7.7).
