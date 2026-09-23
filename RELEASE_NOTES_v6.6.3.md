# TaiHoaDonDienTu v6.6.3

## Thay đổi chính

- `BaoCao_LoiTaiHD` chỉ giữ các lỗi chưa xử lý.
- Khi danh sách, hóa đơn liên quan, thông tin liên quan, chi tiết hoặc XML tải thành công, lỗi cũ đúng khóa nghiệp vụ được xóa tự động.
- Dòng STT trong báo cáo lỗi được đánh lại sau khi xóa.
- Danh sách hóa đơn tổng hợp vẫn chỉ tải một lần; trạng thái 2–6 được lọc trong bộ nhớ trước khi gọi `relative`/`related`, tránh tải và phân trang lại cùng dữ liệu.

## Kiểm chứng

- Kiểm tra tĩnh xác nhận mọi nhánh tải thành công đều gọi cơ chế dọn lỗi.
- Kiểm tra runtime xác nhận upsert một lỗi giả lập rồi xử lý thành công sẽ xóa dòng đó khỏi `BaoCao_LoiTaiHD`.
- Template và file release đã được xóa token phiên, định danh đăng nhập và danh sách MST được cache trước khi công khai.
- Không dùng token hoặc cookie được cung cấp trong nội dung trao đổi để build hay kiểm thử.
