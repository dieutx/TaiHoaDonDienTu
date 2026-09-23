# Trạng thái công khai repository

**Trạng thái: READY**

Trước khi phát hành công khai, bản v6.7.1 đã được xử lý như sau:

- Xóa token phiên và định danh đăng nhập được lưu trong template.
- Xóa danh sách đơn vị/MST được cache trong `LinkTraCuu`.
- Thay tham chiếu GUID hóa đơn mẫu bằng dữ liệu giả.
- Dùng screenshot đã che MST và đường dẫn người dùng.
- Build lại workbook, mở lại bằng Excel và chạy kiểm thử UserForm/VBA.
- Quét source và workbook, không còn JWT, cookie hoặc Authorization header được hard-code.

Các mã số thuế còn lại trong source là hằng số định tuyến công khai của nhà cung cấp hóa đơn, không phải dữ liệu tài khoản đăng nhập.

Repository chưa công bố giấy phép phần mềm. Việc công khai source không tự động cấp quyền sử dụng lại theo MIT, GPL, Apache hoặc giấy phép khác.
