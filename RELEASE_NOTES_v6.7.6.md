# TaiHoaDonDienTu v6.7.6

- Tải chi tiết hóa đơn và thông tin liên quan bằng tối đa 4 kết nối HTTP bất đồng bộ, dùng khoảng nghỉ trên form để giãn cách gửi và giảm tải khi gặp 429/503/504. Ghi chi tiết theo thứ tự gốc với bộ đệm tối đa 12 kết quả; thông tin liên quan ghi đúng dòng tổng hợp. Xem [logic và kiểm thử JSON song song](docs/JSON_PARALLEL_DOWNLOAD.md).
- Giữ đầy đủ response JSON/ZIP trong task bất đồng bộ bằng biến cục bộ trước khi gán kết quả; khi retry thành công, xóa thông báo lỗi tạm thời cũ. Kiểm thử kiểm tra nội dung response và số kết nối thực tế.

- Sửa các finding của review ngày 07/10/2026: JSON null/thiếu, lỗi retry báo thành công khi chưa ghi được, tổng thuế nhiều dòng, XPath và thuế dự phòng XML, trùng tên file, dọn Temp, phân biệt báo cáo theo kỳ/đơn vị, nhận diện API error, khóa UI và phiên token, escape JSON đăng nhập và fixture CI.
- Tải XML song song bằng tối đa 6 request MSXML, giãn cách khởi đầu 300 ms, cooldown 429 chung và phục hồi dần theo logic `hddt-downloader-windows`. Worksheet và file vẫn được ghi tuần tự. Chi tiết tại [docs/XML_PARALLEL_DOWNLOAD.md](docs/XML_PARALLEL_DOWNLOAD.md).
- File tải mới dùng tên gồm MST, mẫu số, ký hiệu và số hóa đơn. Giải nén và kiểm tra XML trong thư mục riêng, không xóa thư mục Temp của tác vụ khác.

- Chuyển thời điểm ISO có múi giờ về UTC+7 trước khi lấy ngày hóa đơn; ví dụ `2026-09-11T17:00:00Z` hiển thị `12/09/2026`. Ngày không có múi giờ giữ nguyên. Không phụ thuộc múi giờ của Windows.
- Ngày trong thông tin liên quan cũng dùng UTC+7. Từ chối ngày/giờ ISO sai thay vì để `DateSerial`/`TimeSerial` tự cuộn sang ngày khác. Đã tái hiện lỗi lùi ngày trên workbook 6.7.5 và kiểm tra 35 ca ngày tháng trong Excel; xem [báo cáo 08/10/2026](docs/REVIEW_2026-10-08.md).
- Ưu tiên link theo MST người bán trong `LinkTraCuu` cho cả tổng hợp và chi tiết. Đọc cả cột C (MST Người bán), kể cả dòng thêm bên dưới danh sách cột B. Chi nhánh dạng `0110269067-xxx` dùng link MST chính khi chưa có cấu hình riêng. Đơn vị có MST khác cần nhập MST dạng Text vào cột C và link vào cột D. Mã tra cứu được đọc cả khi không có MST nhà cung cấp giải pháp.
- Sắp xếp tổng hợp và chi tiết theo ngày tăng dần sau khi kết thúc tải hoặc dừng, kể cả dữ liệu tải bổ sung. Giữ thứ tự các dòng cùng ngày, thông tin liên quan và đánh lại STT tổng hợp.
- Tự điều chỉnh chiều cao dòng kết quả, tối thiểu 18 và tối đa 150 point, mở lại dòng bị ẩn. Bỏ bộ lọc đang áp dụng trước khi sắp xếp để Excel xử lý đủ các dòng; có thể lọc lại sau khi tải.

Review mở rộng trước phát hành bổ sung các bản sửa:

- Giữ nguyên nội dung ngày nhập sai hoặc để trống để người dùng sửa; không tự thay bằng ngày hôm nay.
- Ghi tổng hợp theo cả trang: nếu một hóa đơn gây lỗi, khôi phục dữ liệu trước đó, vị trí ghi và STT; giữ lỗi để tải lại. Kiểm tra trường định danh/ngày của danh sách trước khi ghi.
- Giữ số `0` đầu MST trong báo cáo lỗi để cập nhật và xóa đúng tác vụ. Chặn ghi lặp trang khi tải lại bắt đầu từ cursor đã có.
- Không để metadata `null`/thiếu ghi đè thuế và tổng tiền hợp lệ. Bỏ qua XML không có hàng hóa trước khi ghi, tránh dữ liệu lẫn sang file kế tiếp; nhận cả đuôi `.XML`.
- Sửa retry khi lỗi kết nối xảy ra liên tiếp và tôn trọng `Retry-After` lớn hơn 60 giây. Danh sách/chi tiết dùng token đã chụp đầu phiên.
- HTTP 200 có JSON sai, API error hoặc sai loại dữ liệu liên quan vẫn được giữ trong báo cáo lỗi; chỉ xóa lỗi khi xử lý thành công.
- Escape ký tự tên file bằng bốn chữ số hex để tránh hai định danh khác nhau tạo cùng tên. ZIP tải lại được kiểm tra trong staging trước khi thay ZIP hợp lệ đã có.
- Thêm bộ kiểm tra phát hành dừng ngay khi một script thất bại. Xem [bằng chứng review mở rộng](docs/REVIEW_EXPANDED_2026-10-08.md).

Danh sách hóa đơn vẫn phân trang tuần tự theo cursor; tải song song áp dụng cho chi tiết, thông tin liên quan và XML. Lỗi HTTP 504 ở bước lấy danh sách vẫn phụ thuộc gateway phía GDT.

Với dữ liệu cũ bị sai ngày, cần tải lại để lấy thời điểm gốc; việc sắp xếp không tự sửa ngày đã lưu.
