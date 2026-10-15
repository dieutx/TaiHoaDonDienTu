# HTTP 200 với nội dung HTML trong pha chi tiết

Sự cố ghi trong `BaoCao_LoiTaiHD`:

| Nguồn API | Công đoạn lỗi | Endpoint | HTTP status | Nội dung lỗi |
|---|---|---|---|---|
| sco-query | Parse dữ liệu | `/api/sco-query/invoices/detail?...` | 200 | `Error parsing JSON: <!DOCTYPE h ^ Expecting '{' or '['` |

## Nguyên nhân

Phản hồi chứa **trang web (HTML) thay cho JSON** nhưng vẫn ở trạng thái 200. Log hiện có chưa đủ để xác định nguyên nhân: có thể là phiên đăng nhập, gateway/WAF hoặc chuyển hướng đến trang web. Source không tắt redirect; chưa ghi nhận chuỗi chuyển hướng thực tế của request lỗi.

`IsSuccessfulGdtStatus` chỉ nhìn status, nên trước bản sửa đây được tính là tải thành công:

- `clsGdtXmlTask.Tick` đặt `Success = True`, `frmTaiHoaDon.WriteDetailRecord` chuyển body HTML cho `JsonConverter.ParseJSON`, lỗi parse được ghi vào `BaoCao_LoiTaiHD` kèm HTTP 200.
- Vì kết quả được coi là thành công, hoá đơn **không** vào `QueueFinalRetry` và không được thử lại: chi tiết bị mất trong phiên, chỉ còn một dòng báo cáo gây hiểu nhầm là dữ liệu hỏng.
- Hành vi này có từ bản gốc (commit `8f1ca0a`), không phải hồi quy của scheduler song song `605c103`. Scheduler mới chỉ làm tình huống này dễ gặp hơn vì pha chi tiết/thông tin liên quan gửi request song song (4 kết nối) thay vì một request mỗi lần kèm khoảng nghỉ.

## Thay đổi

- `modGdtRetry.IsGdtMarkupResponse`: nhận diện body bắt đầu bằng `<` (bỏ qua BOM/khoảng trắng) trên request JSON.
- `modGdtRetry.ExecuteGdtRequest`: 2xx + body markup là lỗi tạm thời — thử lại theo backoff hiện có, hết lượt thì `ShouldQueueFinalRetry` để pha cuối phiên thử lại; `InvalidJsonBody` giữ HTTP 200 thật trong báo cáo.
- `clsGdtXmlTask.Tick`/`RetryOrComplete`: cùng quy tắc cho task bất đồng bộ, kèm `throttle.RateLimited` để pha hiện tại giảm tải khi bị trả trang web. Task hết lượt đi vào `QueueFinalRetry` sẵn có của từng pha.
- `modGdtRetry.BuildGdtMarkupResponseMessage`: thay thông báo parse bằng `phản hồi không phải JSON mà là trang HTML/XML`, kèm `Content-Type` và đoạn đầu phản hồi (giới hạn 500 ký tự như `BuildHttpErrorMessage`).
- `clsGdtRequestResult.InvalidJsonBody` và log chi tiết của form hiển thị `Lỗi: phản hồi không phải JSON` thay cho `Lỗi HTTP 200`.

Giới hạn phạm vi có chủ đích: chỉ body dạng markup bị coi là lỗi tầng vận chuyển. JSON hỏng nhưng vẫn là JSON (`{bad json`), body rỗng và `null` giữ nguyên hành vi cũ vì `related` cho phép phản hồi rỗng hợp lệ. Endpoint ZIP/XML nhị phân không đổi.

## Kiểm thử

Fixture localhost `build/Test-ParallelJsonRuntime.ps1` thêm mode `markup`: hai lượt chi tiết đầu của hoá đơn số 2 trả `200` + `<!DOCTYPE html>...` với `Content-Type: text/html; charset=utf-8`, lượt retry cuối phiên nhận JSON thật.

- Negative control trên workbook build trước khi sửa (`dist/ui-pagination-20261015/TaiHoaDonDienTu_v6.7.7.xlsm`): `markup=FAIL: exhausted failure not queued exactly once; first report: Error parsing JSON: <!DOCTYPE h ^ Expecting '{' or '['` — tái hiện đúng dòng báo cáo của sự cố, chứng minh test bắt được lỗi.
- Sau khi sửa: `markup=PASS`, cả 13 mode PASS và `PARALLEL JSON RUNTIME TEST PASSED`.
- Các suite liên quan đường retry/throttle/ghi dữ liệu đều PASS: `Test-DownloaderParityRuntime`, `Test-FinalRetryAfterRuntime`, `Test-CoreFixesRuntime`, `Test-ReviewFixesRuntime`; `tests/Test-DetailInvoiceStatic.ps1` và `tests/Test-RetryUiStatic.ps1` PASS.

## Giới hạn

- Chưa xác minh trên GDT thật: chưa rõ tỉ lệ token hết hạn so với gateway trả trang web. Bản sửa không tự động đăng nhập lại; khi phiên đã bị chặn thì mọi request vẫn lỗi và người dùng cần đăng nhập lại.
- Chỉ phát hiện markup ở đầu body. Trang web không bắt đầu bằng `<`, hoặc JSON đúng cú pháp nhưng sai nội dung, vẫn đi theo đường parse cũ.
- Fallback detail sau XML 500/504 vẫn tuần tự và chịu cùng quy tắc retry này.
