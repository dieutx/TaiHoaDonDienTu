# Tải chi tiết và thông tin liên quan song song

Từ phiên bản 6.7.6, ứng dụng dùng tối đa 4 request MSXML HTTP bất đồng bộ cho từng pha thông tin liên quan và chi tiết hóa đơn. Các pha chạy nối tiếp: danh sách → thông tin liên quan → chi tiết → XML. Danh sách vẫn dùng cursor `state` của trang trước để lấy trang tiếp theo; thay đổi này không tăng tốc bước phân trang hoặc loại bỏ lỗi 504 phía gateway.

`frmTaiHoaDon.ProcessJsonInvoiceRequests` dùng lại task/throttle của XML với chế độ JSON, giữ endpoint và header đang có. Không thêm worker ghi worksheet. `clsGdtXmlTask` giữ tên component và chế độ ZIP mặc định để tương thích; các tham số mới cho JSON nằm cuối lời gọi `Configure`.

- Tối đa 4 kết nối đồng thời; khoảng nghỉ trên form là khoảng cách tối thiểu giữa các lần gửi, mặc định 600 ms. Request còn chờ response không ngăn request khác khởi chạy khi còn slot. Response nhanh có thể không cần dùng đủ 4 slot.
- Thông tin liên quan được ghi ngay khi hoàn tất, vào dòng tổng hợp đã lưu cho hóa đơn. Trạng thái 2–5 lấy relative/related; trạng thái 6 chỉ lấy related; các trạng thái khác được bỏ qua như trước.
- Chi tiết được ghi theo thứ tự danh sách gốc, kể cả khi response về đảo thứ tự. Bộ đệm giới hạn 12 tác vụ/kết quả; khi đầy, ngừng thêm tác vụ cho đến khi ghi được kết quả đầu. Điều này giới hạn bộ nhớ khi một hóa đơn chậm hoặc đang retry.
- 429 áp dụng cooldown chung, tôn trọng `Retry-After` và giảm tải theo policy XML. JSON 503/504 cũng giảm tải và áp dụng cooldown theo backoff/`Retry-After`; XML giữ policy 429 hiện có. Sau thời gian ổn định, số kết nối và khoảng cách gửi phục hồi dần.
- Mỗi request giữ retry có giới hạn. Request hết lượt được ghi vào báo cáo lỗi và đưa vào retry cuối phiên nếu đủ điều kiện. Retry cuối phiên hiện vẫn tuần tự. Parse/ghi lỗi không được báo thành công.
- Token được chụp đầu phiên. Tạm dừng chặn request mới/retry và vẫn nhận, ghi kết quả đang chạy. Dừng hoặc 401/403 hủy request còn chờ và chặn request mới.

Task bất đồng bộ đọc response vào biến cục bộ rồi gán lại `ResponseText`/`ResponseBody`, tránh mất nội dung khi truyền trường đối tượng qua tham số ByRef. Regression kiểm tra cả nội dung JSON lẫn ZIP/binary.

## Kiểm thử

```powershell
.\build\Test-ParallelJsonRuntime.ps1 -BuiltWorkbook '.\dist\parallel-json-final\TaiHoaDonDienTu_v6.7.6.xlsm'
```

Fixture HTTP localhost và Excel kiểm tra kết nối đồng thời thực tế, đúng dòng mua/bán, thứ tự chi tiết, bộ đệm 12 tác vụ với 24 hóa đơn, MST Text, ngày UTC+7, trạng thái hóa đơn, 429/503/504 rồi hồi phục, lỗi hết retry và retry cuối phiên, JSON hỏng, pause/resume, stop/auth, header, request-id và token snapshot. Fixture đổi endpoint chỉ trong bộ nhớ, đóng workbook không lưu. Fixture đo thời gian so với HTTP tuần tự cùng độ trễ giả, kiểm tra cả khoảng nghỉ 1 ms và 600 ms; không dùng credential hay dữ liệu hóa đơn thật.

Chưa đo tốc độ hoặc tải chịu được trên GDT thật. Tăng kết nối giúp khi request độc lập mất thời gian chờ, nhưng tốc độ còn phụ thuộc khoảng nghỉ, server, Excel và số hóa đơn cần lấy thông tin liên quan. Kết quả đo localhost không phải cam kết tốc độ trên GDT.

## Kết quả bản build cuối ngày 08/10/2026

Workbook `dist/parallel-json-final/TaiHoaDonDienTu_v6.7.6.xlsm`, SHA256 `2538C55507E20565215FC941E0C436D60D6F902432C813E9214A16DF9134DFCA`. `build/Test-Release.ps1` hoàn tất với `RELEASE CHECKS PASSED`, exit code 0; retry/UI static chạy lại sau đó cũng trả 0. Source VBA khớp workbook, cấu trúc baseline và cả sáu UserForm PASS. Các hồi quy chi tiết, thông tin liên quan, 35 ca ISO date, JSON/XML/ZIP, scheduler, review mở rộng và cập nhật phiên bản đều PASS.

Fixture JSON mới đạt cả 12 nhóm, bao gồm 24 hóa đơn với bộ đệm tối đa 12 tác vụ. Server đo tối đa 4 request đang chờ, không trùng request-id và không sai header/token. Kết quả tại `dist/parallel-json-final/parallel-json-runtime-result.json`; bản sao báo cáo cấu trúc, chi tiết và retry/UI nằm cạnh workbook.

| Đo 8 request chi tiết với độ trễ giả giống nhau | Thời gian | Kết nối đồng thời đo được |
|---|---:|---:|
| HTTP tuần tự, không thêm khoảng nghỉ | 7,605 giây | 1 |
| JSON song song, khoảng nghỉ 600 ms | 5,160 giây | 3 |
| JSON song song, khoảng nghỉ 1 ms chỉ trong fixture localhost | 2,594 giây | 4 |

Hai phép đo song song bao gồm ghi 16 dòng chi tiết vào Excel; mốc tuần tự chỉ đo HTTP, nên đây là so sánh bảo thủ cho phần chờ mạng. Mức tăng tốc tương ứng khoảng 1,47 và 2,93 lần trong fixture này. Server ghi nhận khoảng cách gửi nhỏ nhất 607 ms ở cấu hình 600 ms. Giá trị 1 ms dùng để kiểm tra scheduler, không được đặt làm mặc định của bản bàn giao.

Các số đo trên thuộc bản kiểm thử trước phát hành. Workbook phát hành được build sạch và kiểm thử lại riêng; xem [bằng chứng phát hành](RELEASE_v6.7.6.md) và [bản phát hành v6.7.6](https://github.com/dieutx/TaiHoaDonDienTu/releases/tag/v6.7.6).
