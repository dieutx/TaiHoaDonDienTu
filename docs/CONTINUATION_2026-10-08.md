# Xác minh lại bản ứng viên 6.7.6 — 08/10/2026

Tiếp tục từ [review mở rộng](REVIEW_EXPANDED_2026-10-08.md), kiểm tra lại workbook hiện có và tạo một bản build sạch từ source/template hiện tại. Không sửa source VBA, template, endpoint, authentication, tên component/control hoặc cặp FRM/FRX. Không commit, push, tag hay phát hành.

## Workbook hiện có khác bản trong báo cáo trước

`dist/release-candidate/TaiHoaDonDienTu_v6.7.6.xlsm` hiện có SHA256 `1CD65E27DF9CFB2DEB989BBFDBF862ADCC8C3304A9D583763D780AC92395BD67`, khác SHA256 ghi trong review trước. Kiểm tra lại bằng `build/Test-Build.ps1 -RunUserFormInstantiation` trả exit code 1: bảng `MENU|tblDanhSachMST` có vùng `A6:E9`, trong khi baseline là `A6:E8`.

Các check khác đều đạt, bao gồm source VBA khớp repository và khởi tạo sáu UserForm. Chưa xác định nguyên nhân workbook thay đổi; không coi kết quả PASS trước đó là bằng chứng cho file hiện tại. Giữ nguyên workbook này. Báo cáo FAIL được giữ tại `dist/verified-20261008/existing-candidate-comparison.json`.

## Bản build sạch đã kiểm thử

- Workbook: `dist/verified-20261008/TaiHoaDonDienTu_v6.7.6.xlsm`.
- SHA256 sau tất cả kiểm thử: `380FDCC644015CE923398E3150D7CBE763065464E6523BCE21150F080913B391`.
- Log: `dist/verified-20261008/verification.log`.
- Kết quả: `CLEAN CANDIDATE VERIFICATION PASSED`, tiến trình tổng trả exit code 0.

Build bằng `build/Build-Excel.ps1 -Version '6.7.6' -OutputPath '.\dist\verified-20261008\TaiHoaDonDienTu_v6.7.6.xlsm'`. Sau đó chạy các script trong tiến trình riêng, dừng ngay khi exit code khác 0: detail static, retry/UI static, release gate, cấu trúc cùng sáu UserForm, related runtime, invoice date runtime, review fixes runtime, extended review runtime và update runtime. Chạy lại retry/UI static cuối cùng để báo cáo sử dụng bằng chứng runtime mới.

| Kiểm tra | Kết quả |
|---|---|
| Build, source VBA, cấu trúc baseline và dữ liệu công khai | PASS |
| Ghi chi tiết mua/bán và sáu UserForm | PASS |
| Static VBA/PowerShell và release gate | PASS |
| Thông tin hóa đơn liên quan | PASS |
| ISO date và ngày thông tin liên quan | PASS; 35 ca ISO |
| JSON/thuế/XML/ZIP, scheduler, pause/stop/auth/retry | PASS |
| Review mở rộng | PASS; 11 nhóm |
| Cập nhật phiên bản | PASS; bản mới, cùng bản, JSON hỏng, offline |

Các JSON runtime nằm trong cùng thư mục workbook. Báo cáo cấu trúc, chi tiết và retry/UI tại `tests` đã được tạo lại cho bản build sạch; bản sao được giữ trong `dist/verified-20261008`. Fixture sửa VBA chỉ trong phiên kiểm thử hoặc bản sao dùng một lần; workbook bàn giao giữ source production.

Kiểm thử bằng Excel COM trên máy hiện tại, dữ liệu giả và HTTP localhost. Chưa thử đăng nhập/tải hóa đơn thật, tốc độ server thật hoặc ma trận Office 32-bit/64-bit. README, VERSION và update.json vẫn giữ phiên bản công khai 6.7.5.

## Suggested GitHub Issue: static check phụ thuộc báo cáo build cũ

`tests/Test-RetryUiStatic.ps1` đọc `tests/comparison-report.json` và trả FAIL khi báo cáo cấu trúc cũ có check lỗi. `build/Test-Release.ps1` chạy script này trước bước build, nên một báo cáo FAIL từ workbook đã thay đổi có thể chặn việc tạo ứng viên sạch. Đây là phụ thuộc nhận thấy khi đọc code; chưa tái hiện riêng bằng fixture trong vòng này.

Phạm vi đề xuất: tách kết quả kiểm tra source khỏi bằng chứng runtime cũ, hoặc chỉ đánh giá bằng chứng runtime sau build của lần chạy hiện tại. Rủi ro: không được bỏ qua lỗi source hoặc lỗi runtime thật. Kiểm thử dự kiến: báo cáo cũ FAIL không chặn build sạch; lỗi static hiện tại vẫn chặn build; lỗi cấu trúc/UserForm của workbook mới vẫn chặn phát hành. Chưa sửa pipeline trong vòng xác minh này.
