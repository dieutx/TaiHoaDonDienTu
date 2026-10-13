# Bản sửa XML, đăng nhập và link BKAV — 13/10/2026

Đã sửa các lỗi ưu tiên từ vòng rà soát repo. Đây là source đang phát triển, chưa phải release mới. README, VERSION, modVersion và update.json vẫn giữ phiên bản công khai 6.7.6. Kết quả bên dưới ghi lại vòng kiểm tra XML/auth/BKAV ban đầu; các sửa này được đưa vào PR cùng [bản sửa core tiếp theo](CORE_FIXES_2026-10-13.md), có thêm hai ca văn bản XML và chạy lại toàn bộ release gate.

## Hành vi sau sửa

- Parser XML nhận namespace mặc định/prefix trên cấu trúc hóa đơn hiện có và namespace XMLDSIG trên đường dẫn chữ ký hiện có. Mỗi bước XPath giữ nguyên vị trí trong cây và giới hạn namespace; trường hàng hóa thuộc namespace khác không ghi đè trường hóa đơn cùng tên.
- MST người bán/người mua và MST nhà cung cấp được ghi dạng Text. Ngày hóa đơn/ngày ký được ghi thành ngày thật của Excel, hiển thị `dd/mm/yyyy`; timestamp có múi giờ dùng helper UTC+7 sẵn có. Trường tùy chọn không có vẫn để trống.
- Giá trị số XML dùng dấu chấm được đọc độc lập thiết lập vùng Windows. Chuỗi không hợp lệ như `invalid`, `10,5` bị báo lỗi thay vì chuyển một phần thành 0 hoặc 10. Metadata thuế/tổng tiền trống dùng fallback cũ; giá trị sai được báo lỗi. Giữ cách tính thuế/tổng tiền hiện có.
- Khi một file lỗi giữa chừng, các dòng đang ghi của file đó được xóa, lỗi được ghi vào `BaoCao_LoiTaiHD`, và parser tiếp tục file khác. Nhập lại file đã sửa xóa lỗi cũ cùng định danh file. Handler dùng `Resume` để rời trạng thái xử lý lỗi trước khi tiếp tục, theo [tài liệu VBA của Microsoft](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/resume-statement).
- Form trích xuất khôi phục con trỏ và thông báo lỗi khi không đọc được thư mục/file hoặc bước chuẩn bị thất bại.
- Đăng nhập/CAPTCHA kiểm tra JSON root và các trường `token`, `key`, `content` trước khi cập nhật trạng thái phiên. JSON hỏng, trường thiếu/null/sai kiểu/rỗng có thông báo thân thiện. Không đưa response đăng nhập vào thông báo lỗi parse. Request giữ endpoint/header/payload cũ, đặt timeout resolve/connect 5 giây và send/receive 15 giây mỗi pha. Lỗi lưu workbook sau đăng nhập có thông báo riêng.
- Link BKAV ở sheet chi tiết dùng trường `id` nếu đúng định dạng GUID thay cho chuỗi `[DLHDon Id]`. Thiếu/sai GUID hiển thị thông báo; không đoán mã tra cứu. Link cấu hình riêng của người bán vẫn được ưu tiên như trước.

Giữ worksheet CodeName, tên component/control, bố cục cột và cặp FRM/FRX. Không sửa template, API architecture, cơ chế tải song song hay việc lưu tài khoản hiện có. Không triển khai tải PDF tổng quát, tự cài đặt hoặc chuyển cấu hình giữa workbook trong vòng này.

## Kiểm thử mới

`build/Test-CompatibilityFixesRuntime.ps1` có 30 nhóm và được thêm vào `build/Test-Release.ps1`:

- 18 nhóm XML cho cả mua/bán: không namespace, namespace mặc định/prefix, trường namespace khác, thuế sai, số sai ở dòng thứ hai, ngày sai, metadata thuế trống và file lỗi rồi file hợp lệ. Kiểm tra rollback, giữ MST đầu 0, ngày ký UTC+7, số thập phân, thuế/tổng tiền, KCT và xóa lỗi cũ khi nhập lại.
- 11 nhóm validation đăng nhập/CAPTCHA: hợp lệ, JSON hỏng, root array/scalar, trường thiếu/null/số/chuỗi trắng/object và thiếu nội dung CAPTCHA. Kiểm tra lỗi không chứa chuỗi dữ liệu nhạy cảm giả.
- 1 nhóm ghi link BKAV bằng macro chi tiết thực tế: GUID hợp lệ, GUID thiếu và chuỗi không hợp lệ.

Chạy trên workbook thử nghiệm bằng Excel COM; fixture chỉ sửa dữ liệu/VBA trong phiên bộ nhớ và đóng workbook không lưu. Các lần debug ban đầu từng bị kẹt tại dispatch COM của harness; bản test cuối dùng tham số có kiểu rõ ràng và tạo dòng lỗi giả trực tiếp thay vì dispatch macro báo lỗi nhiều tham số. Toàn bộ 30 nhóm chạy lại thành công trên bản build cuối với source production, không dùng instrumentation.

`Test-RelatedInvoiceRuntime.ps1` cũng được chỉnh cleanup: giải phóng COM con trước khi đóng workbook/Excel, bỏ chờ finalizer cưỡng bức và chỉ in PASS sau cleanup. Bài test đã chạy lại riêng, PASS và exit code 0. Không đổi assertion hoặc nghiệp vụ hóa đơn liên quan.

## Kết quả xác minh

Chạy các kiểm tra trên bản sao source tại `dist/compat-fixes-20261013/verification-root` để bảo toàn báo cáo có sẵn tại `tests`.

| Kiểm tra | Kết quả |
|---|---|
| Detail static, retry/UI static, parse PowerShell | PASS |
| Release gate với fixture exit 17 | PASS |
| Build và mở lại workbook; source VBA khớp bản sao source hiện tại | PASS |
| Cấu trúc baseline, 21 check, sáu UserForm | PASS, 0 check lỗi |
| Kiểm tra dữ liệu công khai | PASS |
| Chi tiết/tổng hợp mua bán, ngày, MST, link, sort/chiều cao | PASS |
| Hóa đơn liên quan | PASS; chạy riêng lại cleanup mới, exit 0 |
| Ngày ISO/UTC+7 | PASS, 35 ca và ngày thông tin liên quan |
| JSON/XML/ZIP/retry/pause/stop/auth/form | PASS, 6 nhóm review fixes |
| Review mở rộng | PASS, 11 nhóm |
| Regression XML/auth/BKAV mới | PASS, 30 nhóm |
| JSON song song | PASS, 12 nhóm localhost |
| Cập nhật: bản mới/cùng bản/JSON hỏng/offline | PASS, chạy riêng bước cuối, exit 0 |
| `git diff --check` | PASS |

Runner đã đi qua các bước trên tới test cập nhật rồi trả exit code 1 vì bản sao source chưa có thư mục `dist` mà test cập nhật sử dụng. Tạo thư mục đó và chạy lại riêng `Test-UpdateRuntime.ps1`: cả bốn ca PASS, tiến trình trả exit code 0. Vì vậy mọi bài kiểm tra đều đã đạt qua các lần chạy được ghi nhận; không ghi rằng một lần chạy duy nhất của runner đã kết thúc `RELEASE CHECKS PASSED`.

## Workbook và bằng chứng

- Workbook thử nghiệm: `dist/compat-fixes-20261013/final/TaiHoaDonDienTu_v6.7.6.xlsm`.
- SHA256: `761ABCA80795CA1E84921050723DFBEA9F7E32A6AE2C128E120382FD4E9A19F5`.
- Template không đổi, SHA256: `C965621E72B5E6BEACC7B489D623370DAB05CC6E641B86AA9469FF34B890A729`.
- Báo cáo mới: `dist/compat-fixes-20261013/final/compatibility-fixes-runtime-result.json`.
- Báo cáo runtime date/review/extended/parallel nằm cạnh workbook; cấu trúc/UserForm, chi tiết và hóa đơn liên quan nằm trong `verification-root/tests`.

Để kiểm tra lại từ repo hiện tại:

```powershell
.\build\Build-Excel.ps1 -Version '6.7.6' -OutputPath '.\dist\my-verification\TaiHoaDonDienTu_v6.7.6.xlsm'
.\build\Test-CompatibilityFixesRuntime.ps1 -BuiltWorkbook '.\dist\my-verification\TaiHoaDonDienTu_v6.7.6.xlsm'
```

## Giới hạn

Chỉ dùng fixture giả và HTTP localhost. Chưa thử đăng nhập/tải từ GDT thật, link BKAV thật hoặc toàn bộ ma trận Office 32-bit/64-bit. Validation token kiểm tra kiểu và nội dung không rỗng, không khẳng định token còn hiệu lực trên server. Namespace được hỗ trợ trên cấu trúc hiện có; chưa hỗ trợ mọi schema/chữ ký nhà cung cấp. Số dùng dấu phẩy bị báo lỗi để tránh suy đoán. Chưa có file gốc tác giả để chứng minh tương đương toàn bộ nghiệp vụ với bản gốc.
