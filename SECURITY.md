# Chính sách bảo mật

## Dữ liệu nhạy cảm

Không đăng lên Issue, Discussion hoặc Pull Request:

- Username, password hoặc dữ liệu CAPTCHA.
- Token, JWT, cookie, session ID hoặc Authorization header.
- Private key, API key hoặc client secret.
- Mã số thuế thật gắn với tài khoản thử nghiệm. MST định tuyến trong `LinkTraCuu!B:C` là dữ liệu cần thiết cho link tra cứu và được giữ trong workbook phát hành.
- XML/PDF/HTML/ZIP hóa đơn thật.
- Tên, địa chỉ, số điện thoại, tài khoản ngân hàng hoặc thông tin khách hàng/doanh nghiệp nhạy cảm.

## Báo cáo lỗ hổng

Không tạo Issue công khai nếu báo cáo chứa credential hoặc cho phép truy cập dữ liệu hóa đơn.

1. Sử dụng **Private vulnerability reporting / Security advisory** của GitHub nếu repository đã bật tính năng này.
2. Nếu chưa có kênh bảo mật riêng, liên hệ maintainer qua kênh riêng trên hồ sơ GitHub và chỉ mô tả tối thiểu để thiết lập kênh trao đổi an toàn.
3. Thu hồi token, đổi mật khẩu và kết thúc phiên đăng nhập ngay nếu credential đã bị chia sẻ nhầm.

## Sanitize log

Chỉ giữ thông tin cần để tái hiện lỗi.

```text
MST: 01********
Authorization: Bearer [REDACTED]
Cookie: [REDACTED]
Tên khách hàng: [REDACTED]
Số hóa đơn: 12***
```

Không dùng thao tác che một phần nếu phần còn lại vẫn có thể nhận diện cá nhân/doanh nghiệp. Với file đính kèm, hãy tạo fixture giả thay vì sửa trực tiếp hóa đơn thật.

## Nội dung báo cáo an toàn

- Phiên bản project.
- Phiên bản Windows và Microsoft Excel.
- Luồng thao tác gây lỗi.
- HTTP status và error message đã sanitize.
- Schema tối giản hoặc fixture giả lập.
- Kết quả mong đợi và kết quả thực tế.

## Phạm vi hỗ trợ

Project phụ thuộc hệ thống bên ngoài. Thay đổi endpoint, CAPTCHA, authentication, header và schema response có thể gây gián đoạn nhưng không mặc nhiên là lỗ hổng của repository.

Trạng thái rà soát trước khi public được ghi tại `docs/PUBLIC_READINESS.md`. Không chuyển repository sang public khi tài liệu này còn finding mức cao chưa xử lý.
