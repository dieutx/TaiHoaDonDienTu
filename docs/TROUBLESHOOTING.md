# Xử lý sự cố

Trước khi gửi Issue, ghi lại phiên bản project, Windows, Excel, bước gây lỗi và thông báo đã được sanitize. Không đính kèm workbook hoặc hóa đơn thật.

## Không đăng nhập được

- Kiểm tra mã số thuế, mật khẩu và CAPTCHA.
- Đóng form rồi mở lại để lấy CAPTCHA mới nếu mã đã hết hiệu lực.
- Kiểm tra kết nối đến `https://hoadondientu.gdt.gov.vn/`.
- Không đăng credential hoặc ảnh màn hình có thông tin đăng nhập lên Issue.

## CAPTCHA sai hoặc không hiển thị

- Mở lại form để tải CAPTCHA mới.
- Kiểm tra kết nối mạng và thời gian hệ thống Windows.
- Nếu giao diện nguồn thay đổi, ghi phiên bản và mô tả hiện tượng; không tự gửi ảnh CAPTCHA thật nếu ảnh chứa định danh phiên.

## Token hết hạn / HTTP 401 hoặc 403

- Đăng xuất rồi đăng nhập lại.
- Chạy lại tác vụ sau khi có phiên mới.
- Không sao chép Authorization header, token hoặc cookie vào Issue.

## HTTP 429

- Tăng khoảng nghỉ giữa các request.
- Chờ trước khi chạy lại; ứng dụng sẽ tôn trọng `Retry-After` nếu server cung cấp.
- Giảm khoảng thời gian tra cứu hoặc chia tác vụ thành nhiều lần.

## Timeout, HTTP 500/502/503/504

- Kiểm tra kết nối và trạng thái hệ thống nguồn.
- Để ứng dụng hoàn thành cơ chế retry/backoff.
- Với API `relative`/`related`, xem lỗi cuối cùng ngay tại cột **Chuỗi hóa đơn liên quan** hoặc **Thông tin liên quan**.
- Kiểm tra `BaoCao_LoiTaiHD` và chỉ chia sẻ log đã sanitize.

## Lỗi request-id hoặc API thay đổi

- Xác nhận đang dùng release mới nhất.
- Ghi endpoint dạng đường dẫn, HTTP status và schema tối giản; xóa query chứa định danh thật.
- Không tự thay endpoint bằng giá trị chưa được xác minh.

## Không tải được XML/HTML ZIP

- Kiểm tra đã chọn thư mục có quyền ghi.
- Kiểm tra dung lượng ổ đĩa và phần mềm bảo mật có chặn Excel hay không.
- Kiểm tra token còn hiệu lực và xem `BaoCao_LoiTaiHD`.

## Lỗi parse XML

- Không đăng file XML thật.
- Tạo fixture tối giản bằng dữ liệu giả, giữ lại cấu trúc node gây lỗi.
- Ghi rõ loại hóa đơn mua vào/bán ra và nhà cung cấp nếu thông tin đó không nhạy cảm.

## File bị chặn hoặc macro không chạy

- Xác minh file được tải từ GitHub Release chính thức.
- Trong Properties của file, chọn **Unblock** nếu Windows hiển thị tùy chọn này.
- Kiểm tra Excel Trust Center và chính sách macro của tổ chức.
- Không hạ mức bảo mật toàn hệ thống chỉ để chạy một file chưa xác minh.

## Build từ source thất bại

- Đảm bảo Excel đã kích hoạt và không có hộp thoại modal.
- Bật **Trust access to the VBA project object model**.
- Dùng Windows PowerShell 5.1 hoặc PowerShell 7 có Excel COM.
- Không chỉnh `.frx` bằng text editor.
- Nếu gặp lỗi COM `0x80040155`, chạy script hiện có trước; cân nhắc Repair Office nếu lỗi đăng ký COM vẫn còn.

## Export hoặc ghi file thất bại

- Chọn thư mục có quyền ghi và không bị khóa bởi ứng dụng khác.
- Đóng file đích nếu đang mở trong Excel.
- Không dùng đường dẫn quá dài hoặc tên file chứa ký tự bị Windows cấm.

## Thông tin cần có trong bug report

- Phiên bản project, Windows và Excel.
- Office 32-bit hay 64-bit nếu biết.
- Bước tái hiện tối giản.
- Kết quả mong đợi và thực tế.
- Error message/log đã sanitize.
- Xác nhận không có credential hoặc dữ liệu hóa đơn thật.
