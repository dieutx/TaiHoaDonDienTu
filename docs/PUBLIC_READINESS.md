# Trạng thái công khai repository

**Trạng thái: READY**

Trước khi phát hành công khai, bản v6.7.2 đã được xử lý như sau:

- Xóa token phiên và định danh đăng nhập được lưu trong template.
- Giữ lại MST tra cứu trong `LinkTraCuu!B:C` vì bảng này dùng để chọn link tra cứu hóa đơn; chỉ xóa dữ liệu phiên đăng nhập trong `MENU!D7:E7`.
- Thay tham chiếu GUID hóa đơn mẫu bằng dữ liệu giả.
- Dùng screenshot đã che MST và đường dẫn người dùng.
- Build lại workbook, mở lại bằng Excel và chạy kiểm thử UserForm/VBA.
- Quét source và workbook, không còn JWT, cookie hoặc Authorization header được hard-code.

Các mã số thuế trong `LinkTraCuu!B:C` và source là dữ liệu định tuyến tra cứu, không phải dữ liệu tài khoản đăng nhập. Bài kiểm tra workbook công khai xác nhận bảng tra cứu còn nguyên sau bước sanitize.

Repository chưa công bố giấy phép phần mềm. Việc công khai source không tự động cấp quyền sử dụng lại theo MIT, GPL, Apache hoặc giấy phép khác.
