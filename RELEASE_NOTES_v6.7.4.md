# TaiHoaDonDienTu v6.7.4

## Cột chi tiết hóa đơn

- Thêm cột `Mẫu số hóa đơn` trước `Ký hiệu HĐ` trên cả `ChiTietHD_Mua` và `ChiTietHD_Ban`.
- Đồng bộ lại toàn bộ vị trí cột dữ liệu chi tiết, tiền thuế, link tra cứu và mã tra cứu sau khi thêm cột.
- Định dạng MST người bán và MST người mua dưới dạng Text trước khi ghi để giữ nguyên số `0` ở đầu.
- Điền đủ 15 trường thông tin chung từ cột A đến O cho mọi dòng hàng hóa của cùng một hóa đơn.

## Kiểm thử

- Thêm kiểm tra tĩnh cho ánh xạ cột, định dạng MST và bước build workbook.
- Thêm fixture runtime giả lập cho cả hóa đơn mua vào và bán ra; xác nhận mẫu số, ký hiệu, các cột hàng hóa và MST có số `0` đầu sau khi Excel VBA thực thi.
- Build kiểm tra lại cấu trúc workbook, dữ liệu công khai và runtime chi tiết hóa đơn trước khi tạo artifact `.xlsm`.

## Trạng thái phát hành

Candidate v6.7.4 chỉ được phát hành sau khi maintainer kiểm tra trực tiếp file `.xlsm` và xác nhận đạt.
