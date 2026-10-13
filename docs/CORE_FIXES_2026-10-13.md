# Bản sửa dữ liệu, log và tải bổ sung — 13/10/2026

PR này xử lý các lỗi đã rà soát trong [CORE_REVIEW_2026-10-13.md](CORE_REVIEW_2026-10-13.md). Đây là source đang phát triển, chưa nằm trong release công khai 6.7.6. Không tăng số kết nối, đổi endpoint, authentication flow, worksheet CodeName, tên component/control hoặc cặp FRM/FRX. Các bản sửa XML/auth/BKAV trước đó được giữ và kiểm tra chung.

## Tải bổ sung và resume

1. Chọn đúng đơn vị, chiều mua/bán, kỳ và công đoạn cần tải.
2. Bấm **Tải hóa đơn**, chọn **No** tại câu hỏi xóa dữ liệu để tải bổ sung. **Yes** vẫn xóa dữ liệu của chiều đang chọn rồi tải mới.
3. Lưu workbook sau khi tải hoặc dừng nếu muốn dùng trạng thái đó sau khi đóng Excel. Đăng nhập lại khi phiên hết hạn.

Tải bổ sung luôn lấy lại danh sách của kỳ đã chọn. Không lưu hoặc dùng lại cursor cũ. Khóa gồm tài khoản đang chọn, chiều mua/bán, nguồn `query`/`sco-query`, MST người bán, mẫu số, ký hiệu và số hóa đơn. Hai nguồn API được giữ riêng; chưa có bằng chứng để gộp chúng.

- Loại bản ghi trùng trong một trang và giữa các trang của cùng phiên. Tổng hợp đã có được cập nhật đúng dòng theo định danh, kể cả sau sort.
- Chỉ bỏ request chi tiết khi toàn bộ JSON tổng hợp vẫn khớp trạng thái đã lưu và số dòng chi tiết thuộc đúng khóa còn đủ. Có tổng hợp không đồng nghĩa đã có chi tiết.
- Khi tổng hợp thay đổi hoặc chi tiết thiếu, tải lại chi tiết. Chỉ xóa khối chi tiết cũ sau khi ghi khối mới thành công; lỗi tải/ghi giữ khối cũ. Việc xóa được gom cuối pha, theo nhóm từ dưới lên.
- Chỉ bỏ tải XML khi ZIP đọc được, mọi XML/HTML trong ZIP có file giải nén tương ứng đúng kích thước, XML parse được và bằng chứng file khớp lần thành công đã lưu. Đổi thư mục, mất hoặc hỏng file làm tải lại. Kiểm tra này không phải checksum nội dung; HTML hoặc XML đã chỉnh nhưng vẫn hợp lệ và cùng kích thước có thể không được phát hiện.
- Thư mục ZIP và giải nén được khởi tạo từ ô chọn thư mục, kiểm tra tồn tại trước câu hỏi xóa dữ liệu.
- Dữ liệu tổng hợp/chi tiết cũ chưa có metadata chỉ được nhận lại khi MST người mua (mua vào) hoặc người bán (bán ra) khớp chính xác tài khoản đang chọn. Nếu chưa xác minh được, giữ nguyên dữ liệu và ghi thêm dòng thuộc tài khoản hiện tại. Không tự dọn các bản trùng lịch sử.

Metadata nằm ở cột ẩn `BM:BN` của tổng hợp, `AI:AJ` của chi tiết JSON; trạng thái nằm trong sheet very-hidden `TrangThaiTaiHD`, schema `Identity (v1)`. Không đổi cột hiển thị hoặc sheet hiện có. Nếu các cột này đã có nội dung khác hoặc sheet trạng thái có schema khác, dừng với lỗi thay vì ghi đè. Không dùng trạng thái JSON cho các sheet trích xuất XML riêng.

Trạng thái chứa khóa tài khoản và JSON tổng hợp hóa đơn, **không chứa token/password/cursor**. Đây vẫn là dữ liệu riêng tư. Không chia sẻ workbook đã tải dữ liệu. Build từ template sạch; sanitizer xóa metadata/trạng thái mới nhưng không phải công cụ xóa mọi dữ liệu hóa đơn trong workbook cá nhân.

## Ghi dữ liệu và hiển thị

- Các trường văn bản được vá trong writer tổng hợp, chi tiết JSON và XML dùng định dạng Text; giữ `=1+1`, dấu nháy đầu và số 0 đầu. Số tiền và ngày vẫn dùng kiểu số/ngày Excel.
- Số JSON ở các trường tiền/chi tiết được kiểm tra toàn chuỗi, đọc dấu chấm và exponent độc lập locale; chuỗi sai không được chuyển một phần thành số. Giữ trường thuế hỗn hợp dạng chuỗi hoặc số theo kiểu phản hồi.
- Mã trạng thái ngoài bảng hiển thị **Chưa xác định (mã)**; không làm rollback cả trang và không bị diễn giải thành trạng thái cần gọi API hóa đơn liên quan.
- Ghi tổng hợp dùng JSON đã parse; cập nhật nhiều dòng rời nhau vẫn rollback cả trang khi có lỗi. Khôi phục đúng công thức có sẵn và văn bản giống công thức, cùng cursor/STT.
- Cấu hình metadata thuế được đọc một lần mỗi hóa đơn chi tiết; số dòng sở hữu được đọc theo khối một lần đầu phiên. Báo cáo lỗi đọc thêm account theo khối, chỉ renumber khi thực sự xóa và ghi STT một lần bằng mảng. Chưa chuyển toàn bộ writer sang ghi mảng hoặc công bố tăng tốc writer trên GDT.

## Log, thời gian và kết quả

- Log có ngày và giờ, giữ phần cuối tối đa 12.000 ký tự và cắt theo ranh giới dòng. Detail/XML thêm HTTP, số lượt thử và thời gian từ lần gửi đầu, gồm retry/cooldown.
- Thời gian toàn phiên dùng đồng hồ có phần ngày; định dạng không quay về 00 giờ sau 24 giờ. Logic ngày ISO/UTC+7 hiện có được giữ và kiểm tra lại.
- `Retry-After` nhận số giây và HTTP-date dạng IMF-fixdate GMT. Ngày sai/chuỗi hỏng không gây overflow; khoảng chờ vượt miền hỗ trợ dừng phiên có thông báo, tránh gửi sớm. Chưa hỗ trợ hai định dạng HTTP-date lỗi thời.
- Còn lỗi của phiên/tài khoản/chiều hiện tại thì hiển thị số lỗi; cursor lặp cảnh báo danh sách có thể chưa đủ. Không hiện `Xong.` trước kết luận cuối. Lỗi scheduler XML truyền về handler chung và báo cáo lỗi.
- Progress vẫn đo tác vụ đã xử lý, không bảo đảm mọi hóa đơn đã tải đủ. Dừng/auth failure giữ dữ liệu đã ghi; resume chi tiết/XML thành công xóa lỗi cũ tương ứng.

## Kiểm thử

Các script chạy trên Windows với Excel COM, fixture giả hoặc HTTP localhost; không đăng nhập hay tải hóa đơn thật. Fixture runtime đóng workbook không lưu, kiểm tra hash artifact trước/sau; ca resume lưu/mở lại dùng bản sao tạm.

```powershell
.\build\Test-Release.ps1 -Version '6.7.6' -OutputPath '.\dist\release-candidate\TaiHoaDonDienTu_v6.7.6.xlsm'
.\build\Test-CoreFixesRuntime.ps1 -BuiltWorkbook '.\dist\release-candidate\TaiHoaDonDienTu_v6.7.6.xlsm'
```

- `Test-CoreFixesRuntime.ps1`: 44 nhóm, gồm mua/bán, text/số/trạng thái mới, metadata xung đột, dữ liệu cũ, tải mới sau xóa, trùng trong/giữa trang, resume, thay đổi dữ liệu, thiếu dòng, khác tài khoản/nguồn, thay chi tiết thành công/thất bại, rollback công thức, sort đổi dòng target, ZIP/XML/HTML thiếu/hỏng, thư mục XML, ngày/retry/log, đếm lỗi, scheduler bỏ request đã hoàn tất, schema lạ và lưu/mở lại.
- `Test-CompatibilityFixesRuntime.ps1`: 32 nhóm; thêm văn bản giống công thức và dấu nháy XML ở cả hai chiều vào 30 nhóm XML/auth/BKAV trước đó.
- Release gate chạy cả kiểm tra cấu trúc/sáu UserForm, dữ liệu public, chi tiết, related, 35 ca ngày ISO, review/extended, JSON song song và cập nhật online/offline giả lập.

Kết quả chạy cuối và hash workbook được ghi bên dưới sau khi release gate hoàn tất. Chưa kiểm tra Office 32-bit hoặc GDT thực tế. Không có workbook Excel gốc của tác giả để so sánh toàn bộ; kiểm tra tương thích dựa vào template/baseline và regression của repository.

## Phần để PR sau

- Retry trực tiếp các mục được chọn từ báo cáo lỗi, đặc biệt LIST/RELATED; phiên hiện tại vẫn truy vấn lại kỳ và tải lại related.
- Thống kê đầy đủ từng công đoạn: thành công, không có XML, chưa chạy/hủy và lỗi; hiện chỉ có progress, số lỗi phiên và số bỏ trùng/bỏ công đoạn.
- Benchmark writer theo khối trước khi thay phần ghi từng ô còn lại; kiểm kê các trường/phần ghi legacy ngoài writer được vá.
- Checksum file, xác minh nội dung chi tiết đã sửa thủ công và quản lý dung lượng trạng thái. Resume hiện kiểm tra chữ ký tổng hợp/số dòng và ZIP/XML theo điều kiện trên; trạng thái lớn quá giới hạn ô Excel sẽ không được lưu để skip.
