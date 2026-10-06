# TaiHoaDonDienTu v6.7.5

## Ngày tháng và thiết lập vùng

- Chuẩn hóa việc nhập và kiểm tra ngày `dd/mm/yyyy`, không để Windows diễn giải theo thứ tự ngày của locale.
- Date picker và so sánh khoảng ngày dùng giá trị Date thay vì chuyển đổi chuỗi theo locale.
- Giữ ngày hóa đơn ở dạng Date thật trong các luồng xử lý và trong workbook.
- Giữ đúng định dạng ngày `dd/mm/yyyy` trong tham số lọc gửi API.

## Phạm vi kiểm thử

- Build workbook mới từ template với phiên bản nhúng `6.7.5`.
- Chạy kiểm tra cấu trúc workbook, dữ liệu công khai và runtime chi tiết hóa đơn qua Excel.
