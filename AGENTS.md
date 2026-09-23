# Hướng dẫn cho AI Coding Agent

File này dành cho Codex, Claude Code, Cursor, GitHub Copilot và các AI Coding Agent khác làm việc trong repository.

## Project Status

Core của project hiện đã hoạt động và được coi là stable baseline.

Không rewrite core hoặc thay đổi behavior nếu task không yêu cầu rõ ràng. Ưu tiên patch nhỏ, giữ tương thích và có bằng chứng kiểm thử.

## Trước khi thay đổi

Agent phải:

1. Đọc `README.md`.
2. Đọc `CONTRIBUTING.md` và `SECURITY.md`.
3. Đọc code, test và tài liệu liên quan trực tiếp đến task.
4. Hiểu behavior hiện tại trước khi đề xuất sửa.
5. Kiểm tra working tree và bảo toàn thay đổi của người dùng.

## AI Contribution Rules

- Chỉ sửa đúng phạm vi task.
- Tránh large refactor và thay đổi mang tính sở thích style.
- Giữ backward compatibility với workbook, VBA component, UserForm và Office hiện tại.
- Không đổi worksheet CodeName, component name, control name hoặc `.frm`/`.frx` pairing nếu không bắt buộc.
- Dùng `apply_patch` cho file text; không chỉnh `.frx` như văn bản.
- Chạy kiểm thử tương xứng với rủi ro thay đổi.
- Cập nhật README, CHANGELOG hoặc troubleshooting nếu behavior thay đổi.
- Giữ attribution trong `README.md` và `NOTICE.md`.
- File release phải có dạng `TaiHoaDonDienTu_vX.X.X.xlsm`.

## Không được làm

- Rewrite toàn bộ project.
- Đổi framework hoặc ngôn ngữ triển khai.
- Đổi API architecture hoặc endpoint khi chưa xác minh.
- Tự tạo undocumented API.
- Thay authentication flow ngoài phạm vi task.
- Hard-code credential.
- Commit token, cookie, Authorization header hoặc private key.
- Commit hóa đơn, XML, PDF, ZIP hoặc dữ liệu doanh nghiệp thật.
- Xóa attribution hoặc làm sai lệch nguồn gốc dự án.
- Xóa compatibility logic khi chưa hiểu và chưa có regression test.
- Tự thêm giấy phép phần mềm khi maintainer chưa quyết định.

## Contribution được khuyến khích

### Bug fix

- Lỗi parse JSON/XML.
- Lỗi download, pagination, timeout, request-id hoặc UI.
- Lỗi hiển thị ngày, trạng thái, progress hoặc log.

### Compatibility

- API/header/response thay đổi đã được xác minh.
- XML schema mới.
- Khác biệt giữa Windows hoặc Office 32-bit/64-bit.

### Documentation

- README, hướng dẫn sử dụng, troubleshooting và screenshot đã sanitize.

### UX

- Thông báo lỗi rõ hơn, progress, log, filter và export trong phạm vi nhỏ.

### Testing

- Parser test, response fixture giả lập và regression test.

### Code quality

- Refactor nhỏ, có test chứng minh không đổi behavior.

## Khi phát hiện vấn đề ngoài phạm vi

Không tự sửa. Ghi đề xuất vào `docs/ROADMAP.md` hoặc mô tả một Suggested GitHub Issue gồm hiện tượng, phạm vi, rủi ro và cách kiểm thử dự kiến.
