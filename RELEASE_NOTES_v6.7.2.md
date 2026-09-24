# TaiHoaDonDienTu v6.7.2

## Hiển thị lỗi API hóa đơn liên quan

- Khi API `relative` vẫn lỗi sau các lượt retry, lỗi cuối cùng được ghi vào cột **Chuỗi hóa đơn liên quan** của đúng hóa đơn.
- Khi API `related` vẫn lỗi sau các lượt retry, lỗi cuối cùng được ghi vào cột **Thông tin liên quan**.
- Thông báo gồm tổng số lần thử, HTTP status và nội dung lỗi nếu có.
- Bỏ cột **Có HĐ liên quan** không còn ý nghĩa; chuyển **Thông tin liên quan** sang cột kế tiếp của bảng kết quả.

## Không chiếm dụng clipboard

- Thay thao tác `Range.Copy` bằng `FillDown` khi điền thông tin chung cho nhiều dòng chi tiết.
- Không còn xóa clipboard Windows trong quá trình tải hóa đơn, nên có thể tiếp tục dùng `Ctrl+C`/`Ctrl+V` ở ứng dụng khác.
- Sửa luôn trường hợp hóa đơn chỉ có một dòng chi tiết để không tác động nhầm sang dòng kế tiếp.

## Kiểm thử

- Bổ sung regression test cho lỗi `relative`/`related`, cấu trúc cột mới và thao tác điền dữ liệu không dùng clipboard.
- Build/reopen workbook thành công, 35 kiểm tra đạt và cả 5 UserForm khởi tạo thành công trong Excel.
