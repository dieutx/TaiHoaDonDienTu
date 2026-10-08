# Kiểm thử bản sửa lỗi xuất hóa đơn v6.7.6

Ngày chạy: 07/10/2026. Workbook: `dist/fix-675/TaiHoaDonDienTu_v6.7.6.xlsm`.

- `Test-DetailInvoiceStatic.ps1`: đạt.
- `Test-DetailInvoiceRuntime.ps1`: đạt trong Excel, cả mua vào và bán ra. Kiểm tra ngày UTC 17:00, giao năm, offset dương/âm, ngày không có múi giờ, đọc MST người bán ở cột C dưới dòng cuối cột B, fallback chi nhánh, link theo MST trên dòng tổng hợp, sắp xếp tăng dần và giữ thứ tự cùng ngày, thông tin liên quan, dòng ẩn/chiều cao 1 point, STT sau sắp xếp.
- Bộ fixture chi tiết hiện có vẫn đạt: mẫu số, MST dạng Text, ngày và dữ liệu hàng hóa.
- Kiểm tra cấu trúc khi build: `PARTIAL` theo script hiện có; mở lại workbook và so sánh cấu trúc đạt, không khẳng định đã thực hiện toàn bộ kiểm tra UserForm.
- `Test-PublicWorkbook.ps1`: báo `PUBLIC WORKBOOK TEST PASSED`, nhưng tiến trình bị kẹt khi dọn dẹp Excel. Đã đóng đúng tiến trình kiểm thử và chạy bài runtime riêng thành công. Vì vậy tiến trình build tổng thể không kết thúc với mã thành công, dù workbook đã được lưu và các fixture runtime đã đạt.
- `Test-RetryUiStatic.ps1`: còn một mục FAIL ở kiểm tra ô ngày trống (mục 20), nằm ngoài các đoạn sửa của tác vụ này. Không sửa hành vi nhập ngày trong bản vá.

Chỉ dùng fixture giả lập, chưa đối chiếu trực tiếp hóa đơn XanhSM hoặc tài khoản thuế thật. Chưa phát hành lên GitHub.

## Kiểm tra tiếp ngày 08/10/2026

Bản mới: `dist/review-20261008/TaiHoaDonDienTu_v6.7.6.xlsm`. Đã tái hiện trực tiếp lỗi lùi ngày trên workbook 6.7.5; bản sửa đạt 35 ca ngày tháng, ngày thông tin liên quan và ghi tổng hợp/chi tiết cả mua/bán. Đã kiểm tra thêm link đơn vị thành viên/chi nhánh, ZIP lồng thư mục, tải lại, hai người bán trùng số và báo lỗi ZIP hỏng.

Các mục chưa đạt/chưa hoàn tất trong báo cáo 07/10 đã được kiểm tra lại: static retry/UI không còn FAIL; kiểm tra cấu trúc PASS và cả 6 UserForm khởi tạo thành công; public workbook PASS với exit code 0. Script chi tiết bị kẹt finalizer được sửa và chạy lại PASS với exit code 0. Báo cáo đầy đủ: [docs/REVIEW_2026-10-08.md](../docs/REVIEW_2026-10-08.md).
