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
- Bản sửa 6.7.6 không cắt `Retry-After` dạng số giây xuống 60 giây; thời gian chờ có thể dài hơn khi server yêu cầu.
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

## Chi tiết hoặc thông tin liên quan vẫn tải chậm

- Bản sửa 6.7.6 có log `DETAIL scheduler`/`RELATED scheduler` khi vào pha tải JSON song song, tối đa 4 kết nối. Khoảng nghỉ trên form là khoảng cách giữa các lần gửi; server phản hồi nhanh có thể không cần dùng đủ 4 kết nối.
- Danh sách `query`/`sco-query` vẫn phân trang tuần tự. Nếu log đang ở bước này và gặp 504, việc tăng tốc chi tiết/thông tin liên quan chưa giúp được.
- 429/503/504 ở pha JSON làm bộ tải giảm tốc và chờ retry. Chi tiết giữ thứ tự gốc và ngừng thêm tác vụ khi bộ đệm 12 kết quả đầy; thông tin liên quan được ghi ngay vào dòng hóa đơn tương ứng.
- Xem [logic tải JSON song song](JSON_PARALLEL_DOWNLOAD.md) và `BaoCao_LoiTaiHD`. Tốc độ đo trên localhost không phải tốc độ cam kết trên GDT.

## Không tải được XML/HTML ZIP

- Kiểm tra đã chọn thư mục có quyền ghi.
- Kiểm tra dung lượng ổ đĩa và phần mềm bảo mật có chặn Excel hay không.
- Kiểm tra token còn hiệu lực và xem `BaoCao_LoiTaiHD`.
- Bản sửa 6.7.6 đợi giải nén hoàn tất, hỗ trợ XML/HTML trong thư mục con của ZIP và lưu tên file có MST người bán để tránh trùng số hóa đơn. Lỗi lưu/giải nén của từng hóa đơn được ghi trong `BaoCao_LoiTaiHD`; các hóa đơn khác tiếp tục được xử lý.
- Nếu ZIP tải lại bị hỏng, ZIP hợp lệ đã có vẫn được giữ; chỉ thay ZIP đích sau khi kiểm tra giải nén thành công.

## Ngày nhập sai hoặc trang dữ liệu không ghi được

- Nhập ngày tồn tại theo `dd/mm/yyyy` và để ngày bắt đầu không vượt ngày kết thúc. Bản sửa 6.7.6 giữ nguyên ô sai/để trống, báo lỗi và không tự chọn ngày hôm nay.
- Nếu một trang tổng hợp có dữ liệu không hợp lệ, bản sửa khôi phục cả trang về trước khi ghi và giữ tác vụ lỗi để tải lại. Kiểm tra `BaoCao_LoiTaiHD`; không coi HTTP 200 là đã xử lý thành công nếu JSON hoặc dữ liệu ghi bị lỗi.

## Ngày hóa đơn bị lùi một ngày trên 6.7.5

- 6.7.5 lấy phần ngày của chuỗi API trước khi chuyển múi giờ. Ví dụ `2026-09-11T17:00:00Z` được ghi thành 11/09, trong khi tại Việt Nam là 12/09.
- Bản sửa 6.7.6 chuyển thời điểm có múi giờ về UTC+7, giữ nguyên ngày không có múi giờ và lưu ngày thật của Excel để sắp xếp tăng dần. Không cần đổi thiết lập vùng hoặc múi giờ Windows.
- Tải lại dữ liệu cũ bị sai ngày; sắp xếp không khôi phục được thời điểm gốc đã mất.
- Với link XanhSM của đơn vị có MST riêng, nhập MST dạng Text vào `LinkTraCuu!C:C` và link tra cứu vào cột D. Chi nhánh dạng `0110269067-xxx` có thể dùng cấu hình MST chính khi chưa có link riêng.

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
