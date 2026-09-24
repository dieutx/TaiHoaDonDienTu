# TaiHoaDonDienTu v6.7.3

## Thông báo cập nhật tùy chọn

- Khi mở workbook, kiểm tra metadata phiên bản trên GitHub. Nếu có bản mới, hiện phiên bản, ngày phát hành và ghi chú; người dùng có thể mở trang tải hoặc tiếp tục dùng bản hiện tại.
- Lỗi mạng hoặc metadata không hợp lệ không chặn việc mở workbook.

## Khôi phục bảng tra cứu hóa đơn

- Khôi phục dữ liệu `LinkTraCuu!B2:C122` từ workbook v6.3 gốc: 117 MST cột B và 4 MST người bán cột C.
- Bước sanitize chỉ xóa dữ liệu phiên đăng nhập và giữ bảng tra cứu. Build kiểm tra số lượng MST và đối chiếu từng ô B:C với template.

## Kiểm thử

- Build v6.7.3, kiểm tra cấu trúc và kiểm tra workbook công khai đạt; cả 6 UserForm khởi tạo và đóng thành công trong Excel.
- Kiểm tra runtime thông báo cập nhật đạt với bản mới, cùng phiên bản, JSON lỗi và mất kết nối.
- Kiểm tra runtime hóa đơn liên quan đạt.

## Giới hạn đã biết

Logic chọn link hiện tra cột B sang D; cột C chưa tham gia chọn link. Nếu cột B có giá trị trùng, dòng xuất hiện sau sẽ ghi đè link của dòng trước.
