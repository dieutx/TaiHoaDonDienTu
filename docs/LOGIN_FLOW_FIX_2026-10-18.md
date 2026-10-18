# Sửa nút Đăng nhập không phản hồi — 18/10/2026

## Nguyên nhân và phạm vi sửa

`cmdDangNhap_Click` trước đây thoát ngay nếu `blnCaptcha = False`, không thông báo và không tự lấy CAPTCHA. Cờ này chỉ được cập nhật khi ô mật khẩu phát sinh `AfterUpdate` hoặc `Exit`. Chọn tài khoản đã lưu rồi bấm Đăng nhập có thể chưa qua các event đó. Hai event cũng có thể lấy CAPTCHA hai lần và chuyển focus ngay trước khi nút nhận click.

Bản sửa chuyển việc lấy CAPTCHA vào handler của nút: kiểm tra tài khoản, lấy mã nếu thiếu, rồi gửi đăng nhập. Sửa tài khoản/mật khẩu làm mất hiệu lực CAPTCHA cũ; mã đã gửi không được dùng lại sau lỗi. Nút hiển thị trạng thái xử lý, tạm khóa trong request và được khôi phục ở mọi nhánh kết thúc. CAPTCHA rỗng/nhận dạng thiếu được báo lỗi thay vì gửi mã không hợp lệ.

Giữ endpoint, timeout, thuật toán nhận dạng, JSON đăng nhập, cách lưu tài khoản/token và bước chuyển sang form tải hóa đơn hiện có. Không sửa `.frx` hoặc đổi tên control/component.

## Kiểm chứng

- `Test-LoginFlowRuntime.ps1`: khởi tạo form và 16 ca đăng nhập đạt trong Excel COM. Bao gồm tài khoản đã lưu, thiếu tên/mật khẩu, mất mạng, HTTP/lỗi mạng CAPTCHA, JSON CAPTCHA sai/thiếu key, mã rỗng/thiếu ký tự, HTTP/timeout/thông báo từ chối đăng nhập rồi thử lại, JSON đăng nhập sai/thiếu token và thay mật khẩu. Test chạy handler thật, parser, nhận dạng SVG và lưu token trên bản sao tạm với toàn bộ dữ liệu giả; SHA-256 của workbook đầu vào không đổi.
- Cùng ca `SavedAccount` chạy trên bản build cũ ngày 15/10/2026 thất bại: không có request CAPTCHA hoặc đăng nhập. Bản sửa gửi đúng một request mỗi bước.
- `Test-Build.ps1`: toàn bộ 21 kiểm tra cấu trúc/source đạt. Báo cáo tổng thể là `PARTIAL` vì lần chạy này không bật smoke test tất cả UserForm; khởi tạo riêng form đăng nhập được kiểm tra trong bộ runtime phía trên.
- `Test-PublicWorkbook.ps1`, `Test-ReleaseGate.ps1` và `git diff --check`: đạt.
- Chưa thực hiện đăng nhập GDT bằng tài khoản thật; chưa chạy toàn bộ bộ kiểm thử phát hành.

Lệnh kiểm thử mới được thêm vào `Test-Release.ps1` và README. Bản build cục bộ nằm tại `dist/login-flow-20261018/TaiHoaDonDienTu_v6.7.7.xlsm`; đây không phải bản phát hành GitHub mới.

## Suggested GitHub Issue ngoài phạm vi

**Hiện tượng:** `Test-Build.ps1 -RunUserFormInstantiation` bị kẹt khi chạy `Invoke-UserFormSmoke.ps1` cho `frmCapNhatMK`; gọi riêng script smoke cho `frmDangNhap` với `-Mta` cũng không kết thúc. Khởi tạo cùng form qua bộ runtime mới thì đạt. Chưa xác định được vị trí kẹt trong script smoke hay quá trình dọn COM.

**Phạm vi/rủi ro:** harness kiểm thử Excel; có thể làm quy trình phát hành chờ vô hạn. Không sửa trong bản vá đăng nhập này.

**Kiểm thử dự kiến:** thêm log từng bước mở workbook, khởi tạo/unload form và giải phóng COM; giới hạn thời gian cho từng tiến trình con; thử riêng STA/MTA và bảo đảm chỉ dọn tiến trình do test tạo.
