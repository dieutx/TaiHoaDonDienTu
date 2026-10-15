# Đánh giá mở rộng core ngày 13/10/2026

Đây là bằng chứng trước sửa. Các patch sau đợt review đã được triển khai ở source đang phát triển: [CORE_FIXES_2026-10-13.md](CORE_FIXES_2026-10-13.md).

## Kết luận

Core còn có thể cải thiện rõ ở tính toàn vẹn dữ liệu, tải bổ sung và chi phí ghi Excel. Thứ tự nên làm là: giữ văn bản đúng nguyên bản → chịu được trạng thái chưa biết → chống trùng khi tải bổ sung → tổng kết kết quả theo công đoạn → giảm chi phí báo cáo lỗi/ghi worksheet → tải lại phần thiếu và phục hồi phiên.

Đợt này chỉ rà soát và tạo bằng chứng bằng fixture; **chưa áp dụng các đề xuất vào source VBA**. Các bản sửa XML/auth/BKAV của đợt trước được giữ nguyên. Không thay endpoint, kiến trúc API, authentication, component, control hoặc worksheet CodeName.

## Phạm vi và bằng chứng

- Đọc luồng danh sách `query`/`sco-query`, cursor, retry cuối phiên, JSON/XML scheduler, throttle, writer, báo cáo lỗi và test liên quan.
- Chạy fixture giả trên workbook đã vá ở `dist/compat-fixes-20261013/final/TaiHoaDonDienTu_v6.7.6.xlsm`, mở chỉ đọc, tắt workbook events và đóng không lưu. Module/class thử nghiệm chỉ tồn tại trong bộ nhớ.
- Script tái hiện: `dist/core-review-20261013/Measure-CoreAudit.ps1`; kết quả: `dist/core-review-20261013/core-audit-results.json`.
- Không gọi GDT, không dùng tài khoản thật hoặc hóa đơn thật. Chuỗi thử công thức duy nhất là `=1+1`, không có liên kết hay lệnh bên ngoài.
- Hash workbook trước/sau giữ nguyên: `761ABCA80795CA1E84921050723DFBEA9F7E32A6AE2C128E120382FD4E9A19F5`.
- Script trả exit code 0 nghĩa là tái hiện đúng các quan sát dưới đây; không có nghĩa những điểm yếu này đã được sửa. Không chạy lại bộ release vì source thực thi không thay đổi.

| Phép thử | Quan sát thực tế |
|---|---|
| Gọi writer tổng hợp hai lần với cùng một hóa đơn | Có hai dòng với cùng định danh |
| Tên người bán là `=1+1` | Ô có `HasFormula=True`, giá trị hiển thị `2` |
| Trang gồm một hóa đơn thường và một hóa đơn có `tthai=99` | Lỗi VBA 9, con trỏ dòng phục hồi và cả hai hóa đơn không được ghi |
| `Retry-After: 120` | Trả 120 giây, đúng behavior hiện có |
| `Retry-After: Wed, 14 Oct 2026 12:00:00 GMT` | Trả 0, không xử lý dạng HTTP-date |
| `Retry-After: 999999999999` | Lỗi VBA 6 do chuyển sang Long |
| 10 lần đánh dấu thành công, báo cáo không có lỗi | Hiển thị 0,000 giây; phép đo chưa đủ độ phân giải để định lượng thời gian rất ngắn |
| 10 lần đánh dấu thành công, báo cáo có 1.000 lỗi không khớp | Khoảng 0,645 giây dù không xóa dòng nào |

Số đo là một lượt fixture cục bộ, không phải tốc độ GDT hoặc cam kết tăng tốc. Trạng thái `99` là dữ liệu giả kiểm tra khả năng chịu schema mở rộng; chưa xác nhận GDT đang trả trạng thái đó.

## Các patch nên ưu tiên

### 1. Giữ văn bản từ JSON/XML là văn bản — ưu tiên cao

**Hiện tượng:** `modGhiExcel.ghiExcel_TongHop` ghi các trường như `nbten` bằng `.Value`; Excel có thể diễn giải chuỗi đầu vào thành công thức. Fixture xác nhận tên `=1+1` mất biểu diễn gốc. Các đường ghi khác cần được kiểm kê tương tự, chưa coi mọi cột đều đã tái hiện lỗi.

**Phạm vi:** thêm helper ghi Text cho tên, địa chỉ, mã hàng, ghi chú, ký hiệu và mã tra cứu. Phân biệt rõ cột Text, ngày và số; `.Value2` riêng lẻ không thay thế việc bảo đảm Text. Không làm biến dạng công thức chủ động của template.

**Rủi ro:** định dạng Text áp dụng quá rộng sẽ khiến số tiền/ngày không tính toán hoặc sắp xếp đúng.

**Kiểm thử:** chuỗi `=1+1`, `+...`, `-...`, `@...`, dấu nháy và số 0 đầu phải round-trip nguyên bản; `HasFormula=False` cho dữ liệu văn bản. Ngày vẫn là serial Excel; số tiền vẫn là số. Kiểm tra tổng hợp, chi tiết và XML cho cả mua/bán.

### 2. Chịu được mã trạng thái mới — ưu tiên cao

**Hiện tượng:** `modGhiExcel.bas` dòng 107–108 truy cập trực tiếp `arrTrangThai(CInt(tthai)+1,1)` và `arrKQKTHoaDon(CInt(ttxly)+2,1)`. Một mã ngoài bảng tra cứu làm writer rollback cả trang. Rollback hiện tại bảo vệ khỏi trang ghi dở; vấn đề là một mã mới có thể chặn cả các hóa đơn hợp lệ cùng trang.

**Phạm vi:** helper kiểm tra miền và kiểu dữ liệu; mã biết dùng nhãn hiện tại, mã chưa biết hiển thị `Chưa xác định (mã ...)` kèm cảnh báo. Giữ nguyên mã thô, không gán trạng thái chưa biết thành trạng thái nghiệp vụ đã biết; không tự suy ra cần gọi related/relative API.

**Rủi ro:** fallback sai có thể che schema lỗi hoặc tạo request liên quan sai. Chỉ nới phần hiển thị enum, giữ kiểm tra định danh/ngày bắt buộc.

**Kiểm thử:** toàn bộ mã hiện tại, mã 99, thiếu/null, chuỗi và số vượt miền; trang trộn mã thường/mã mới phải giữ đủ hóa đơn. Schema hỏng thật vẫn có báo cáo lỗi và rollback như trước.

### 3. Chống trùng khi ghi tiếp/tải bổ sung — ưu tiên cao

**Hiện tượng:** chọn ghi tiếp gọi `taiHoaDon_Total(lrTong+1, lrChiTiet+1)`; các vòng nhận `datas` ghi và xếp lịch mọi hóa đơn. Có bộ nhớ các cursor đã thấy nhưng chưa có bộ khóa hóa đơn chống trùng. Fixture writer xác nhận hai lần ghi cùng hóa đơn tạo hai dòng; rà soát caller chưa thấy bộ lọc ngăn trường hợp này.

**Phạm vi:** xây khóa theo tài khoản, chiều mua/bán, MST người bán, mẫu số, ký hiệu và số hóa đơn. Theo dõi riêng việc đã có tổng hợp, chi tiết, XML và trạng thái/phiên bản dữ liệu; bỏ trùng tổng hợp không được bỏ luôn phần chi tiết/XML còn thiếu. Giữ chế độ ghi tiếp hiện tại tương thích hoặc thêm lựa chọn tải bổ sung rõ ràng.

**Rủi ro:** hóa đơn thay thế/điều chỉnh và dữ liệu cập nhật cần được xử lý đúng; không mặc định mọi bản ghi trùng khóa đều bất biến. Kiểm tra ý nghĩa nguồn `query`/`sco-query` trước khi gộp hai nguồn.

**Kiểm thử:** cùng kỳ tải hai lần; kỳ giao nhau; cùng số nhưng khác seller/mẫu/ký hiệu; nhiều dòng hàng của một hóa đơn; tổng hợp đã có nhưng chi tiết/XML thiếu; retry một trang. Giữ đúng liên kết related với dòng tổng hợp sau khi lọc và sort.

### 4. Phân biệt hoàn tất công việc với tải đủ dữ liệu — ưu tiên cao

**Hiện tượng qua đọc code:** `cmdTaiHoaDon_Click` hiển thị 100%/Hoàn tất khi không stop hoặc auth failure, dù còn lỗi trên báo cáo. `taiHoaDon_Total` cũng hiện `Xong.` trong điều kiện này. Handler lỗi scheduler XML chỉ log/hủy task rồi thoát, không truyền trạng thái lỗi về caller. Chống cursor lặp hiện có giúp tránh vòng lặp vô hạn, nhưng chưa cho người dùng kết luận rõ có thể thiếu trang.

**Phạm vi:** kết quả từng pha gồm số thành công, thiếu/lỗi, không có XML, hủy/chưa chạy và tình trạng phân trang. Tổng kết `Hoàn tất, còn N mục lỗi` hoặc `Đã dừng, còn N mục chưa xử lý`; lỗi scheduler phải được ghi nhận. Đếm lỗi theo phiên/tài khoản/chiều đang chạy, không đếm toàn bộ lịch sử. Không xem 404 thật sự không có XML là lỗi transport.

**Rủi ro:** retry thành công phải trừ lỗi cũ; progress tính theo tác vụ đã xử lý cần tách với số hóa đơn tải đủ.

**Kiểm thử:** lỗi danh sách một kỳ, parse/ghi lỗi, hết retry, XML không có, cursor lặp, stop, 401/403, phục hồi thành công. Đây là finding từ đọc code, chưa tái hiện toàn form bằng fixture mới của đợt này.

### 5. Hoàn thiện `Retry-After` — ưu tiên vừa

**Hiện tượng:** `ReadRetryAfterSeconds` chỉ xử lý số, có thể overflow với giá trị lớn. Dạng HTTP-date cũng là giá trị hợp lệ theo [RFC 9110 mục 10.2.3](https://www.rfc-editor.org/rfc/rfc9110.html#name-retry-after).

**Phạm vi:** parser độc lập nhận thời gian UTC giả được để test; hỗ trợ số giây và HTTP-date, kiểm tra miền trước CLng/DateAdd. Giá trị không thể biểu diễn cần thông báo/dừng chờ có kiểm soát, không âm thầm rút ngắn thời gian server yêu cầu. Áp dụng helper chung cho HTTP tuần tự và task bất đồng bộ.

**Rủi ro:** GMT so với giờ địa phương, đồng hồ lệch, header đã quá hạn; không được phá behavior hiện đang tôn trọng Retry-After 120 giây.

**Kiểm thử:** số 0/120, HTTP-date tương lai/quá khứ, header thiếu/hỏng, số rất lớn; pause/stop trong cooldown và không gửi request trước hạn.

### 6. Giảm chi phí báo cáo lỗi và ghi worksheet — ưu tiên vừa

**Đã đo:** `MarkGdtRetrySuccess` gọi tìm kiếm rồi renumber ngay cả khi không tìm thấy mục cần xóa. `FindGdtErrorRow` đọc vùng lỗi rồi vẫn đọc cột tài khoản từng ô. 1.000 lỗi không khớp gây khoảng 0,645 giây cho 10 lời gọi thành công giả.

**Patch nhỏ nên làm trước:** chỉ renumber khi có xóa; đọc cột account cùng mảng A:R và ghi STT một lần bằng mảng. Đánh giá index trong bộ nhớ sau đó, vì row delete có thể làm index lỗi thời và sheet có thể được người dùng chỉnh.

**Tối ưu kế tiếp qua đọc code:** writer tổng hợp parse lại JSON đã parse để kiểm tra; nhiều trường được ghi từng ô; chi tiết đọc cấu hình metadata thuế trong từng dòng hàng; AutoFit từng dòng và cuối pha có thể làm lại. Cân nhắc cache cấu hình một lần mỗi phiên/hóa đơn và tạo mảng dữ liệu cho một trang hoặc một hóa đơn trước khi ghi `Range.Value2`. Microsoft khuyến nghị giảm số lần truyền dữ liệu bằng cách [đọc/ghi cả khối](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions#read-and-write-large-blocks-of-data-in-a-single-operation).

**Rủi ro:** mất format Text/ngày, công thức template, rollback hoặc thứ tự dòng hàng. Không tăng kết nối/giảm khoảng nghỉ để bù chi phí Excel.

**Kiểm thử:** so sánh giá trị, kiểu, format, link và thứ tự trước/sau; rollback có dữ liệu cũ; report nhiều tài khoản; lỗi không khớp và xóa nhiều dòng. Benchmark với 0/1.000/10.000 lỗi và số dòng đại diện trên cùng máy; chưa có số đo tăng tốc writer ở đợt này.

## Mở rộng tính năng sau các patch dữ liệu

### Tải lại riêng phần lỗi/thiếu

`mFinalRetryQueue` hiện chỉ tồn tại trong bộ nhớ của lần chạy; báo cáo lỗi có lưu nhưng không được tự nạp thành hàng đợi sau khi mở lại workbook. Thêm thao tác chọn công đoạn cần tải lại từ báo cáo sẽ hữu ích trước khi làm resume toàn bộ.

Phải lọc account/chiều; tái tạo request từ định danh đã kiểm tra, không tin endpoint bất kỳ trong sheet; yêu cầu phiên đăng nhập hiện tại và không lưu token vào checkpoint. Dòng target của related cần tìm lại bằng định danh vì sheet đã sort sau lần chạy cũ. Kiểm thử retry độc lập cho LIST/DETAIL/RELATED/XML, lỗi ghi file và hai tài khoản.

### Phục hồi phiên tải sau đóng Excel

Lưu checkpoint theo kỳ và công đoạn đã hoàn thành, có schema version, khóa tài khoản và tùy chọn tải. Cursor có thể hết hạn, nên cần fallback quét lại kỳ kết hợp chống trùng thay vì tin cursor luôn tiếp tục được. Việc lưu dữ liệu/trạng thái phải nhất quán; ghi checkpoint sau khi kết quả được lưu thành công.

Đây là thay đổi rủi ro cao hơn: nên triển khai sau chống trùng và tải lại lỗi. Kiểm thử đóng/mở tại từng pha, không mất/nhân đôi dòng, checkpoint lỗi/hết hạn, workbook chỉ đọc, tài khoản khác và đổi bộ lọc. Chưa có implementation hoặc benchmark resume.

## Phạm vi chưa nên mở rộng ngay

Không cần rewrite core hay đổi ngôn ngữ. Scheduler đã có giới hạn kết nối, cooldown, token snapshot, pause/stop và regression; tăng số kết nối chưa có bằng chứng GDT chịu được. PDF nhiều nhà cung cấp cần URL/định danh đã xác minh và test riêng; không suy ra endpoint tải PDF từ helper mẫu. Những mục roadmap bảo mật cũ không được coi là finding mới nếu chưa đối chiếu lại trạng thái hiện tại.

Đề xuất chia thành các Issue/patch độc lập theo thứ tự ở trên. Giữ tất cả tên workbook sheet, VBA component và Form hiện tại; dùng fixture giả và build cục bộ trước khi thử dữ liệu thực tế.
