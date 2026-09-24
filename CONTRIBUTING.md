# Hướng dẫn đóng góp

Cảm ơn bạn muốn đóng góp cho TaiHoaDonDienTu. Core hiện tại đang hoạt động và là stable baseline, vì vậy Pull Request nên nhỏ, dễ kiểm tra và chỉ thay đổi phạm vi cần thiết.

## Quy trình

```text
Fork repository
    ↓
Clone fork
    ↓
Tạo branch
    ↓
Thay đổi và kiểm thử
    ↓
Commit
    ↓
Push
    ↓
Mở Pull Request
```

Ví dụ:

```bash
git clone https://github.com/<tai-khoan>/TaiHoaDonDienTu.git
cd TaiHoaDonDienTu
git switch -c fix/request-id
```

## Đặt tên branch

Sử dụng một trong các tiền tố sau:

- `feat/`: tính năng mới đã được thống nhất.
- `fix/`: sửa lỗi.
- `docs/`: tài liệu.
- `refactor/`: refactor nhỏ, giữ nguyên behavior.
- `test/`: kiểm thử hoặc fixture giả lập.
- `chore/`: build, release hoặc bảo trì repository.

Ví dụ: `feat/export-csv`, `fix/request-id`, `docs/update-install-guide`.

## Phạm vi thay đổi

- Đọc `README.md`, `AGENTS.md` và code liên quan trước khi sửa.
- Không rewrite core, đổi API architecture, endpoint hoặc authentication flow nếu Issue/PR không yêu cầu rõ ràng.
- Không đổi tên module, function, UserForm, control hoặc worksheet CodeName nếu không cần thiết.
- Không sửa `.frx` bằng text editor.
- Giữ tương thích với workbook/template và Office hiện tại.
- Ghi technical debt chưa xử lý vào `docs/ROADMAP.md` hoặc đề xuất Issue riêng.

## Kiểm thử

Chọn bài test phù hợp với thay đổi:

```powershell
.\tests\Test-RetryUiStatic.ps1
.\build\Build-Excel.ps1 -Version '6.7.2' -OutputPath '.\dist\TaiHoaDonDienTu_v6.7.2.xlsm'
.\build\Test-Build.ps1 -BuiltWorkbook '.\dist\TaiHoaDonDienTu_v6.7.2.xlsm' -RunUserFormInstantiation
.\build\Test-RelatedInvoiceRuntime.ps1 -BuiltWorkbook '.\dist\TaiHoaDonDienTu_v6.7.2.xlsm'
```

Nếu không thể chạy Excel COM, hãy ghi rõ bài test nào chưa chạy và lý do trong PR.

## Commit convention

Ưu tiên Conventional Commits nhưng không bắt buộc tuyệt đối:

```text
feat: thêm tính năng export CSV
fix: sửa lỗi request-id
docs: cập nhật README
refactor: tách hàm parse XML
test: thêm test parser
chore: cập nhật workflow
```

Một commit nên có mục đích rõ ràng và không trộn thay đổi không liên quan.

## Pull Request

PR nên:

- Tập trung vào một vấn đề chính.
- Mô tả bug/feature, giải pháp và phạm vi ảnh hưởng.
- Ghi rõ cách kiểm thử và kết quả.
- Giữ backward compatibility hoặc giải thích rõ thay đổi behavior.
- Cập nhật tài liệu khi behavior thay đổi.
- Có screenshot đã sanitize nếu thay đổi UI.

Tên file release phải theo mẫu `TaiHoaDonDienTu_vX.X.X.xlsm`, không thêm hậu tố mô tả.

## Dữ liệu tuyệt đối không được commit

- Username hoặc password.
- Token, cookie, JWT hoặc Authorization header.
- Mã số thuế thật dùng cho tài khoản thử nghiệm. MST định tuyến trong `LinkTraCuu!B:C` được giữ để tạo link tra cứu hóa đơn.
- XML/PDF/ZIP hóa đơn thật.
- Tên, địa chỉ, số điện thoại hoặc thông tin khách hàng/doanh nghiệp nhạy cảm.

Fixture phải là dữ liệu giả lập và không thể truy ngược đến người nộp thuế thật. Xem thêm `SECURITY.md`.
