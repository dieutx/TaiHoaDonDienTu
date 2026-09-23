# TaiHoaDonDienTu v6.7.0

## Thay đổi chính

- Log giai đoạn lấy danh sách hiển thị rõ kỳ ngày, nguồn `query`/`sco-query`, số trang, mã HTTP, số hóa đơn của trang, tổng lũy kế và thời gian phản hồi.
- Phát hiện `state` rỗng hoặc bị lặp để dừng phân trang, tránh vòng lặp vô hạn khi API trả con trỏ bất thường.
- Sửa nhánh lỗi của `query` để luôn bắt đầu đúng endpoint `sco-query`, không tái sử dụng nhầm URL trước đó.
- Thêm nút **Tạm dừng/Tiếp tục** và **Dừng**. Lệnh dừng có hiệu lực sau request đồng bộ hiện tại; dữ liệu đã tải được giữ lại.
- Các khoảng chờ retry, xử lý hóa đơn liên quan, chi tiết và XML đều phản hồi với trạng thái tạm dừng/dừng.

## Kiểm chứng

- Kiểm tra tĩnh xác nhận log phân trang, chặn `state` lặp và điều khiển pause/stop hiện diện trong các vòng xử lý.
- Kiểm tra runtime khởi tạo toàn bộ UserForm và xác nhận hai nút điều khiển mới tồn tại, mặc định bị vô hiệu hóa khi chưa tải.
- Workbook được mở lại bằng Excel sau build và kiểm tra không phát sinh repair.
- Template và file release không chứa token, cookie, thông tin đăng nhập hoặc danh sách MST đã cache.
