# Sửa lỗi, chọn nguồn hóa đơn và căn bố cục form — 14/10/2026

Source đang phát triển, chưa phát hành phiên bản mới. Bằng chứng trước sửa nằm trong [review 14/10](REVIEW_2026-10-14.md).

## Chọn nguồn tải

Trường **Nguồn hóa đơn** nằm cạnh phần chọn đơn vị, có ba lựa chọn:

- **Tất cả**: giữ hành vi mặc định, truy vấn cả `query` và `sco-query`.
- **Hóa đơn thông thường**: chỉ truy vấn danh sách `query`.
- **Hóa đơn từ máy tính tiền**: chỉ truy vấn danh sách `sco-query`.

Cả mua vào và bán ra đều có ba lựa chọn này. Danh sách được chọn quyết định tập hóa đơn cho các pha related, chi tiết, XML và retry. Không đổi endpoint của các pha hiện có; fallback sau XML 500/504 vẫn dùng endpoint `query/detail` đã có. Nguồn được chụp đầu phiên và khóa khi tải, cùng các tùy chọn khác.

**No** tiếp tục tải bổ sung, chống trùng theo nguồn. Với **Yes** và một nguồn cụ thể, chỉ xóa các dòng có metadata xác nhận đúng tài khoản/chiều/nguồn; giữ nguồn khác, tài khoản khác và dòng legacy chưa có metadata. Dữ liệu mới được ghi tiếp sau vùng còn giữ. Với **Tất cả + Yes**, giữ hành vi xóa toàn bộ dữ liệu của chiều mua/bán đã chọn như trước.

## Các lỗi đã sửa

- Detail đã hoàn tất trong chính phiên được dùng lại khi fallback XML, kể cả chế độ tải mới; không append thêm một bản chi tiết.
- Lỗi parse tách theo loại request; thành công ở detail/related không xóa lỗi của relative còn thiếu. Resume detail chỉ dọn lỗi parse detail.
- Kiểm tra schema trạng thái và cột metadata trước khi xóa dữ liệu.
- Final retry giữ deadline Retry-After từ phản hồi cuối, kiểm tra trước mọi request LIST/DETAIL/RELATIVE/RELATED/XML. Khoảng chờ dùng đồng hồ có phần ngày và độ phân giải nhỏ hơn giây; vẫn nhận pause/stop/auth.
- Kiểm tra định danh seller/template/series/number trước khi ghi detail ở luồng thông thường và final retry, dùng chung bảo vệ với fallback XML.
- Trường mẫu/ký hiệu/số/ghi chú hóa đơn gốc được ghi dạng Text, giữ số 0 đầu và không chạy chuỗi giống công thức. Các cột văn bản trong báo cáo lỗi cũng được giữ dạng Text.

Không đổi tên component/control/sheet cũ hoặc cặp FRM/FRX. Bộ chọn nguồn và bố cục bổ sung được dựng bằng VBA lúc khởi tạo form. Tạm dừng và Dừng có cùng kích thước, nằm trên cùng hàng dưới hai nút thao tác chính và dùng cùng mép cột. Nhãn ghi chú được giữ và chuyển sang phần bên trái để không chồng nút; font của các control mới theo các control hiện có.

## Kiểm thử

`build/Test-ReviewSourceRuntime.ps1` kiểm tra hai chiều mua/bán: lỗi trùng fallback, giữ lỗi parse, từ chối định danh sai, giữ văn bản, không xóa khi schema lỗi, sáu tổ hợp nguồn/chiều, snapshot và khóa selector, xóa đúng scope, căn hàng/cột và font. `build/Test-FinalRetryAfterRuntime.ps1` đo hai request HTTP localhost: Retry-After 5 giây phải được giữ dù cooldown final trong fixture chỉ còn 1 giây. Hai bài test đã được thêm vào release gate.

Fixture dùng dữ liệu giả, inject mã phục vụ thử nghiệm trong bộ nhớ, đóng workbook không lưu và xác nhận hash không đổi. Kiểm tra bố cục có ảnh render từ form Excel thật. Không dùng tài khoản/hóa đơn thật hoặc gọi GDT. Chưa xác minh Office 32-bit.

## Kết quả kiểm thử 14/10/2026

Chạy trên bản archive của working tree hiện tại trong `dist/fixes-20261014/verification-root` để không ghi đè inventory/report của repository.

- `build/Test-Release.ps1 -Version 6.7.6`: **RELEASE CHECKS PASSED**, exit 0, đủ 16 bước: detail-invoice static, retry-ui static, release gate âm, build + structural/public workbook/detail runtime, build structural, related-invoice, invoice-date, review-fixes, extended review, compatibility fixes, core fixes, downloader parity, review/source, final retry-after, parallel JSON, update runtime. Log: `release-gate.log`.
- `build/Test-ReviewSourceRuntime.ps1`: `REVIEW FIXES RUNTIME TEST PASSED`, exit 0, hash workbook không đổi. Bảy ca lỗi review chạy cho cả mua vào và bán ra, mỗi ca đều `PASS`: tải mới chi tiết + XML 504 giữ 2 dòng chi tiết (không nhân đôi), lỗi parse relative còn 1 sau khi related hợp lệ, detail lệch định danh không ghi dòng và giữ 1 lỗi, ghi chú/số hóa đơn gốc và báo cáo lỗi giữ dạng Text, dòng legacy không bị xóa khi lọc nguồn, preflight schema từ chối thì dữ liệu tổng hợp còn nguyên 22 ô. Sáu tổ hợp nguồn × chiều (`all`/`query`/`sco-query` cho mua và bán) đều `PASS`; bốn ca xóa theo nguồn `PASS`. Log: `review-source.log`.
- `build/Test-FinalRetryAfterRuntime.ps1`: `FINAL RETRY-AFTER RUNTIME TEST PASSED`, exit 0. Server localhost trả `Retry-After: 5` trong khi cooldown final của fixture là 1 giây; khoảng cách thực giữa hai request là **5,074 giây** (trước sửa 0,749 giây). Log: `final-retry-after.log`.

### Bố cục form (ảnh render + đo toạ độ thật)

- `Inspect-Layout.ps1` đọc toạ độ control sau `UserForm_Initialize`: `cmdPauseResume` L=581,3 W=90 H=30 và `cmdStopDownload` L=677,3 W=90 H=30 cùng T=342, mép phải 767,3 trùng mép phải cột `cmdTaiHoaDon`/`cmdChiTietXML` (581,3–767,3); `cboInvoiceSource` L=317,3 T=78 W=246 trùng `cboTTHD` và cùng hàng `cboDonVi`; nhãn nguồn theo `Label22`; font các control mới là Verdana 9,8 như các control hiện có. Hàng nút 342–372 nằm dưới `cmdChiTietXML` (kết thúc 336) và cách `lblProgressMessage` (bắt đầu 378) 6 điểm; `Label21` kết thúc ở 569,3, không chồng nút. Log: `layout-dump.log`; dump UTF-8: `layout-dump.txt`, xác nhận nhãn trên artifact đúng tiếng Việt: “Nguồn hóa đơn”, “Tất cả”, “Hóa đơn thông thường”, “Hóa đơn từ máy tính tiền”, “Tạm dừng”, “Dừng”.
- `Inspect-Overlap.ps1` quét toàn bộ control của form: **0** control chồng lấn với `cboInvoiceSource`, `cmdPauseResume`, `cmdStopDownload`. Log: `overlap.log`.
- Ảnh render từ form Excel thật: `form-preview-final.png` (`Render-Form.ps1`, đang chọn nguồn “Hóa đơn từ máy tính tiền”).

### Artifact cuối

- `dist/fixes-20261014/verification-root/dist/TaiHoaDonDienTu_v6.7.6.xlsm`, 382.730 byte, `SHA256 = D1BCC73EDE0C065EB397346C5523476AB4CCB28E5E375675CDA53F88ABB4B9BC`.
- Môi trường: Windows, Windows PowerShell 5.1.22621.6133, Excel COM 16.0 (bản 64-bit). Excel 32-bit chưa được kiểm tra.
- Chưa phát hành phiên bản mới; `VERSION` vẫn là 6.7.6 và chưa commit/push.
- `dist/` chưa được Git theo dõi, nên muốn chia sẻ bằng chứng cần giữ riêng các script và log trong `dist/fixes-20261014/`.
