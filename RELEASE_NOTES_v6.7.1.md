# TaiHoaDonDienTu v6.7.1

## Sửa lỗi hiển thị

- Đưa hai nút **Tạm dừng/Tiếp tục** và **Dừng** vào đúng vùng hiển thị của UserForm.
- Bổ sung kiểm tra runtime xác nhận hai nút đều `Visible` và mép phải không vượt quá `InsideWidth` của form.

## Tính năng từ v6.7.0

- Log chi tiết kỳ ngày, nguồn `query`/`sco-query`, trang, HTTP, số hóa đơn và thời gian phản hồi.
- Chặn `state` rỗng/lặp để tránh phân trang vô hạn.
- Dừng an toàn sau request hiện tại và giữ dữ liệu đã tải.
