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
- Ổn định build trên Windows PowerShell 5.1/Excel COM khi độ rộng cột chi tiết đã đúng hoặc Excel từ chối setter `ColumnWidth` trực tiếp.
- Sửa fixture chi tiết để biến dòng dùng kiểm tra không bị thay đổi bởi tham số VBA `ByRef`.
- Chuẩn hóa đường dẫn workbook tuyệt đối trước khi mở bằng Excel COM trong các bài runtime và UserForm smoke test.

## Trạng thái phát hành

v6.7.4 đã hoàn tất build và kiểm thử bằng Microsoft Excel desktop trước khi phát hành.
