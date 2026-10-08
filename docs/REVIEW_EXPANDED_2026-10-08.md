# Review mở rộng trước phát hành 6.7.6 — 08/10/2026

Tiếp nối [review ngày tháng, link tra cứu và XML/HTML](REVIEW_2026-10-08.md), mở rộng kiểm tra ghi worksheet, phân trang, retry, báo cáo lỗi, nhập ngày và tính toàn vẹn file tải lại. Giữ working tree đã có, không rewrite core, không đổi endpoint/authentication hay tên worksheet/component/control. Chưa commit, push, tag hoặc phát hành.

## Bằng chứng trước khi sửa

Chạy `build/Test-ExtendedReviewRuntime.ps1 -Audit` trên workbook giữ nguyên từ vòng review trước: `dist/review-20261008/TaiHoaDonDienTu_v6.7.6.xlsm`, SHA256 `65EB16A8BD7375B5C4D05D6FEC24A4CE2C27F4D575F4B9FC83EDC39390AAC987`. Workbook này đã sửa ngày UTC+7 nhưng chưa có các patch của vòng mở rộng. Cả 10 nhóm dưới đây đều tái hiện FAIL; bằng chứng máy đọc tại `dist/review-20261008/extended-review-before.json`.

| Lỗi tái hiện | Hành vi trước sửa | Bản sửa và regression |
|---|---|---|
| MST trong báo cáo lỗi | `0000000001` thành `1`, cập nhật/xóa không còn cùng định danh | Đặt cột định danh dạng Text trước khi ghi; kiểm tra thêm/cập nhật/xóa đúng một tác vụ |
| Ghi dở trang tổng hợp | Hóa đơn thứ hai sai ngày để lại dữ liệu trang, cursor từ 450 thành 451 và STT từ 1 thành 2 | Khôi phục nội dung trước trang, cursor và STT khi ghi lỗi; test giữ cả dữ liệu có sẵn trong vùng đích |
| Metadata thuế/tổng tiền | `dlieu` null hoặc thiếu ghi đè giá trị trực tiếp hợp lệ | Chỉ dùng metadata có giá trị; test giữ thuế 35, tổng tiền 235 và đối chiếu đúng |
| Cách ly XML | File không có hàng hóa ghi thông tin chung rồi lẫn sang file sau; đuôi `.XML` bị bỏ qua | Kiểm tra hàng hóa trước khi ghi, xóa vùng ghi của hóa đơn mới, nhận đuôi không phân biệt hoa/thường |
| Retry-After | Server yêu cầu 120 giây nhưng bị cắt xuống 60 | Chỉ áp trần backoff khi không có Retry-After dương; regression kiểm tra phép tính, không chờ thật 120 giây |
| Lỗi kết nối liên tiếp | Lỗi transport thứ hai thoát khỏi handler VBA | Dùng Resume để rời handler trước lần thử tiếp; localhost đóng kết nối hai lần để kiểm tra kết quả cuối |
| Retry giữa phân trang | Cursor hiện tại không được đưa vào tập đã thấy, ghi lặp hóa đơn khi response lặp state | Ghi nhận state của endpoint bắt đầu retry; kiểm tra chỉ ghi một trang |
| Escape tên file | `<A` và ký tự `ChrW(&H3CA)` tạo cùng định danh tên file | Escape mỗi ký tự bằng bốn chữ số hex; test hai tên khác nhau |
| Retry thông tin liên quan | HTTP 200 với JSON hỏng/API error/sai loại root xóa lỗi tải ban đầu | Kiểm tra parse và ghi thành công trước khi xóa; test ba kiểu lỗi rồi response hợp lệ để dọn lỗi cũ |
| Ngày nhập sai | `31/02/2026` tự đổi thành hôm nay; ô xóa trống cũng bị điền lại | Báo sai ngày, giữ nguyên ô nhập, không tạo kỳ tra cứu; test cả hai ô sai, ô trống và cặp ngày hợp lệ |

Các lỗi này được xác nhận trên bản sửa trung gian 6.7.6. Không suy rộng rằng cả 10 lỗi đều đã có trong release 6.7.5; lỗi lùi ngày của 6.7.5 có bằng chứng riêng ở báo cáo trước.

## Bổ sung từ review source

- `ApiGet` nhận token tùy chọn; request danh sách/chi tiết truyền token chụp đầu phiên. Regression xác nhận cả lời gọi cũ dùng token global và lời gọi mới giữ snapshot dù global thay đổi.
- ZIP mới được lưu trong staging và giải nén/kiểm tra trước khi thay ZIP đích. Fixture tải lại ZIP hỏng xác nhận ZIP hợp lệ đã có không mất; tiếp tục kiểm tra ZIP lồng thư mục, XML trùng tên gốc, HTML và hai người bán cùng số hóa đơn.
- Một lệnh PowerShell lỗi có thể bị che bởi lệnh thành công chạy sau trong workflow. Thêm `build/Test-Release.ps1`: mỗi check chạy trong tiến trình riêng, kiểm tra exit code và dừng ngay. Fixture `exit 17` xác nhận check sau không được chạy. Workflow build-release dùng runner này.
- Script kiểm tra workbook công khai bỏ chờ finalizer cưỡng bức sau khi đã đóng/release Excel COM và chỉ in PASS sau cleanup. Đây là phòng tránh kiểu kẹt cleanup đã ghi nhận trong vòng review trước, không phải một lỗi VBA mới.

## Quy trình kiểm thử bản cuối

```powershell
.\build\Test-Release.ps1 -Version '6.7.6' -OutputPath '.\dist\release-candidate\TaiHoaDonDienTu_v6.7.6.xlsm'
```

Runner kiểm tra static VBA/PowerShell và release gate, build từ template, kiểm tra dữ liệu công khai, chạy fixture chi tiết, kiểm tra cấu trúc và khởi tạo cả sáu UserForm, thông tin liên quan, 35 ca ISO date, ghi dữ liệu/XML/scheduler, review mở rộng và cập nhật phiên bản. Bộ kiểm tra chỉ báo thành công sau khi tất cả tiến trình con trả exit code 0.

Lần chạy cuối hoàn tất với `RELEASE CHECKS PASSED`, exit code 0. Bản được kiểm tra: `dist/release-candidate/TaiHoaDonDienTu_v6.7.6.xlsm`, SHA256 `AC2BCFCBB4E3616BEC305FEE5BB32838E05E0CF3883BDD5FA977CAAB877CCC89`.

| Kiểm tra | Kết quả bản cuối |
|---|---|
| Detail static, retry/UI static, parse PowerShell | PASS |
| Release gate với fixture exit 17 | PASS; không chạy check tiếp theo |
| Build, dữ liệu công khai và mở lại workbook | PASS |
| Cấu trúc worksheet/CodeName/component/control/FRM-FRX và sáu UserForm | PASS |
| Chi tiết mua/bán, ngày, MST, link, sort và chiều cao dòng | PASS |
| Relative/related runtime | PASS |
| ISO date và ngày thông tin liên quan | PASS: 35 ca ISO và fixture thông tin liên quan |
| JSON/thuế/XML/ZIP, scheduler, pause/stop/auth/retry | PASS; ZIP hỏng tải lại giữ ZIP hợp lệ |
| Review mở rộng | PASS cả 11 nhóm: 10 nhóm lỗi tái hiện trước sửa và token tương thích/snapshot |
| Update runtime | PASS: bản mới, cùng bản, JSON hỏng và offline |
| `git diff --check` | PASS |

Bằng chứng máy đọc: `tests/comparison-report.json`, `tests/detail-invoice-runtime-result.json`, `tests/retry-ui-test-results.json`, `dist/release-candidate/invoice-date-runtime-result.json`, `dist/release-candidate/review-fixes-runtime-result.json` và `dist/release-candidate/extended-review-runtime-result.json`. Fixture sửa/inject VBA chỉ dùng trong phiên kiểm thử, đóng workbook không lưu; workbook bàn giao giữ source production.

## Phạm vi đã xác minh

Dùng Excel COM trên máy hiện tại, JSON/XML/ZIP giả và HTTP localhost. Không đăng nhập hay tải hóa đơn thật, không thử live GDT/Hilo/XanhSM, không đo tốc độ hoặc tải chịu được của server thật. Không thay cấu hình XML scheduler có sẵn: 6 kết nối/300 ms; fixture scheduler dùng 4 kết nối để kiểm tra policy.

Chưa xác minh ma trận Office 32-bit/64-bit và mọi biến thể XML theo nhà cung cấp. Fixture namespace/schema/numeric XML cần tiếp tục mở rộng theo [ROADMAP](ROADMAP.md); không coi review này là bằng chứng mọi đầu vào đều hợp lệ. README và `update.json` giữ release công khai 6.7.5. Bản trong `release-candidate` dùng để thử trước khi phát hành; dữ liệu cũ đã mất thời điểm gốc cần tải lại để sửa ngày.
