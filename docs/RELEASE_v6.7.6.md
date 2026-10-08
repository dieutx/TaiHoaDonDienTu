# Xác minh phát hành v6.7.6 — 08/10/2026

Build sạch từ source và template, giữ các workbook ứng viên trước đó. File phát hành: `TaiHoaDonDienTu_v6.7.6.xlsm`, kích thước 334422 byte, SHA256 `233DC90D9F5EBB3D8A844CE75B91053AA09E1ADAC00C9614B8F261E1D51CBB70`.

```powershell
.\build\Test-Release.ps1 -Version '6.7.6' -OutputPath '.\dist\release-v6.7.6-20261008\TaiHoaDonDienTu_v6.7.6.xlsm'
```

Quy trình kết thúc với `RELEASE CHECKS PASSED`, exit code 0. Kiểm tra source VBA khớp workbook, cấu trúc baseline và khởi tạo sáu UserForm đều PASS. Dữ liệu công khai, chi tiết mua/bán, thông tin liên quan, 35 ca ISO date, JSON/XML/ZIP, retry, pause/stop/auth, 11 nhóm review mở rộng và cập nhật phiên bản đều PASS. VERSION, source VBA và update.json cùng phiên bản 6.7.6; metadata trỏ đúng asset phát hành.

Fixture JSON song song đạt cả 12 nhóm. Máy chủ localhost đo tối đa 4 request đang chờ; không sai header/token và không trùng request-id. Mốc 8 request HTTP tuần tự mất 7,598 giây; song song với khoảng nghỉ mặc định 600 ms mất 5,133 giây gồm ghi 16 dòng vào Excel, khoảng 1,48 lần trong fixture. Khoảng cách gửi nhỏ nhất đo được 600,199 ms. Cấu hình 1 ms chỉ dùng kiểm thử đạt 2,641 giây và khoảng 2,88 lần; không đổi mặc định của sản phẩm.

Các báo cáo runtime nằm cạnh workbook trong `dist/release-v6.7.6-20261008`; báo cáo cấu trúc, chi tiết và retry/UI tại `tests` được tạo lại cho file phát hành. Chỉ dùng dữ liệu giả và HTTP localhost, không đăng nhập/tải hóa đơn thật; chưa kiểm tra đầy đủ ma trận Office 32-bit/64-bit hoặc tốc độ GDT thật.

[Tải bản phát hành](https://github.com/dieutx/TaiHoaDonDienTu/releases/tag/v6.7.6). Lỗi 504 khi phân trang danh sách vẫn phụ thuộc gateway phía GDT; chi tiết, thông tin liên quan và XML là các pha tải song song.
