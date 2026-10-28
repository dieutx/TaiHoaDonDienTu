# Chia kỳ ngày, CAPTCHA tra cứu và PDF từ XML — 28/10/2026

## Bằng chứng API

Đã phân tích HAR do maintainer cung cấp tại máy cục bộ. Không chép HAR, token, CAPTCHA hoặc dữ liệu hóa đơn thật vào repository.

- `query/invoices/sold` và `query/invoices/purchase`: HAR có kết quả HTTP 200 cho khoảng 30 ngày; live cũng trả 200 cho 5, 10 và 30 ngày.
- `sco-query/invoices/sold` và `sco-query/invoices/purchase`: HAR trả HTTP 400 với thông báo **“Khoảng thời gian tìm kiếm không được lớn hơn 10 ngày.”** cho khoảng 30 ngày.
- Request SCO trong HAR có thêm `ckey` và `cvalue`. Test live không gửi hai trường này trả HTTP 500 ở cả hai chiều. Khi lấy CAPTCHA mới từ `/api/captcha` và gửi đúng hai trường, live trả HTTP 200 cho 5/10 ngày, HTTP 400 và giới hạn 10 ngày cho 30 ngày.
- Vẫn dùng các endpoint hiện có và phân trang `state`; không dùng endpoint PDF giả định. Giới hạn của nguồn `query` chưa được suy rộng ngoài những khoảng đã kiểm tra.

## Cách chia kỳ

Nguồn `query` giữ kỳ theo tháng. Nguồn `sco-query` chia mỗi tháng thành các đoạn tối đa 10 ngày, tính cả ngày đầu/cuối. Mỗi kỳ chỉ chạy đúng một nguồn.

Nếu trang đầu của một kỳ trả HTTP 400 với thông báo giới hạn ngày theo mẫu đã xác minh, chương trình đọc số ngày và thay kỳ đó bằng các kỳ nhỏ hơn. Ví dụ giới hạn đổi thành 5 ngày thì chia lại thành các đoạn 5 ngày. Các đoạn liên tục, không chồng lặp, giữ đủ ngày cuối. Không chia lại sau khi đã ghi các trang của kỳ để tránh lặp dữ liệu.

Lỗi 400 khác, 401/403 hoặc 500 không bị suy đoán thành giới hạn ngày. Nếu không nhận diện được thông báo, lỗi vẫn đi vào luồng báo lỗi/retry hiện tại.

CAPTCHA tra cứu chỉ được thêm khi gọi danh sách SCO. Dùng bộ đọc SVG có sẵn; lấy CAPTCHA mới cho mỗi lần gọi trang hoặc retry cuối phiên. URL chương trình giữ để phân trang, báo lỗi và retry không có `ckey`/`cvalue`; nội dung lỗi được loại giá trị CAPTCHA nếu API phản hồi lại chúng. CAPTCHA đăng nhập và luồng tài khoản không thay đổi.

## Sử dụng PDF

1. Chọn **PDF** trên form tải và chọn thư mục lưu. Có thể chọn kèm tổng hợp, chi tiết, XML/HTML.
2. Chương trình tải/giải nén XML bằng scheduler hiện có, sau đó dựng PDF tuần tự. Nếu không có XML, lấy chi tiết API và ghép tổng tiền từ dòng tổng hợp đúng danh tính hóa đơn. Có đủ dữ liệu thì vẫn dựng PDF; không tạo XML thay thế.
3. PDF mang cùng tên cơ sở với XML: MST người bán, mẫu số, ký hiệu, số hóa đơn. PDF mới được dựng từ XML hiện tại ngay cả khi phần tải XML được resume.
4. Dùng **XML -> PDF** nếu đã có thư mục XML. Chương trình xử lý các file `.xml` ngay trong thư mục, lưu PDF bên cạnh, báo số thành công/lỗi. Có thể chạy macro `GdtExportXmlFolderPdf` trực tiếp.

Mẫu dựng theo bố cục HTML/CSS được cung cấp: thông tin người bán/người mua, bảng hàng hóa 9 cột, tổng hợp thuế suất, tổng thanh toán và thông tin thời điểm ký. Mẫu đã loại dữ liệu doanh nghiệp thật và được nhúng vào VBA khi build. Hỗ trợ XML `HDon/DLHDon` không namespace hoặc dùng namespace mặc định/prefix; bỏ qua trường lạ thuộc namespace khác. MST giữ số 0 đầu; số hóa đơn XML có đệm 0 vẫn đối chiếu được với số API.

Microsoft Edge hoặc Google Chrome chạy headless để in HTML ngoại tuyến thành PDF A4. Không gọi dịch vụ chuyển đổi bên ngoài. Nguồn kỹ thuật: [tài liệu Chrome Headless](https://developer.chrome.com/docs/automation-and-testing/headless-cli).

## Giới hạn và xử lý lỗi

- Cần Edge/Chrome đã cài ở vị trí chuẩn. Thiếu trình duyệt được phát hiện trước khi bắt đầu tải dữ liệu.
- PDF là **bản thể hiện từ XML**, không phải file PDF của nhà cung cấp. XML gốc vẫn là dữ liệu cần giữ để đối chiếu. Chương trình hiển thị thời điểm ký được đọc từ XML, chưa xác minh chữ ký số.
- Không có XML: dùng danh sách hàng hóa từ API detail, kết hợp tổng tiền từ API detail hoặc dòng tổng hợp khi detail để trống. Kiểm tra MST người bán, mẫu số, ký hiệu, số hóa đơn ở cả hai nguồn. Bắt buộc có ngày, người bán, hàng hóa, tiền hàng và các tổng tiền; không suy đoán số tiền thiếu. Dữ liệu chưa đủ hoặc sai danh tính được ghi ở bước `PDF`, giữ PDF cũ. XML có sẵn nhưng lỗi không bị tự bỏ qua để dùng dữ liệu khác.
- PDF từ API ghi rõ **không có XML gốc**, không hiển thị chữ ký giả. Lỗi/thiếu XML gốc vẫn được theo dõi riêng nếu tác vụ XML chưa hoàn tất; PDF dựng thành công không có nghĩa đã lấy được XML.
- Định dạng hiện hỗ trợ cấu trúc hóa đơn `HDon/DLHDon`; XML thông báo, biên lai hoặc schema khác được báo lỗi. Không suy diễn tổng tiền hoặc chữ ký khi trường XML vắng mặt.
- Dữ liệu XML được escape trước khi chèn HTML; cấm DTD, không tải tài nguyên từ mạng. Lỗi dựng PDF không làm XML bị tính là tải thất bại. PDF cũ chỉ được thay sau khi có PDF mới hợp lệ.
- Mỗi lần dựng có giới hạn chờ 60 giây, hỗ trợ nút Dừng trong vòng chờ. Có thể chọn tải lại PDF hoặc dùng nút chuyển thư mục để thử lại lỗi PDF.

## Kiểm thử

```powershell
.\build\Build-Excel.ps1 -Version '6.7.7' -OutputPath '.\dist\pdf-date-verify\TaiHoaDonDienTu_v6.7.7.xlsm'
.\build\Test-DatePdfRuntime.ps1 -BuiltWorkbook '.\dist\pdf-date-verify\TaiHoaDonDienTu_v6.7.7.xlsm'
.\build\Test-DatePdfRegression.ps1 -BuiltWorkbook '.\dist\pdf-date-verify\TaiHoaDonDienTu_v6.7.7.xlsm'
```

Test runtime dùng fixture giả, kiểm tra kỳ liên tục/không chồng lặp, nguồn đã chọn, giới hạn 10/5/3/1/17 ngày, lỗi 400 không liên quan, điều khiển PDF, namespace XML, escape HTML, trường ngoại lai, XML lỗi/DTD/thiếu hàng hóa, danh tính hóa đơn, giữ PDF cũ khi thất bại và dựng hóa đơn 61 dòng. PDF ngắn và dài được render để kiểm tra bố cục, lặp đầu bảng và tổng tiền ở trang cuối.

`Test-DatePdfLive.ps1` là bài test chỉ chạy khi chủ tài khoản cho phép; tham số tài khoản được truyền lúc chạy, chỉ in HTTP và số dòng. Bản workbook thử nghiệm đóng mà không lưu. Không ghi tài khoản vào script hoặc chạy live trong CI.

## Kết quả xác minh bản build cục bộ

Các kết quả ban đầu dưới đây được xác minh trên bản build cục bộ `dist/pdf-date-verify/TaiHoaDonDienTu_v6.7.7.xlsm`. Nhánh PDF dự phòng khi không có XML được bổ sung cho **6.7.8**, với kiểm thử riêng `build/Test-PdfFallbackRuntime.ps1` và bộ release gate; xem [ghi chú phát hành](RELEASE_v6.7.8.md).

| Kiểm tra | Kết quả |
|---|---|
| Static detail và retry/UI | PASS |
| Cấu trúc, source VBA, mở lại workbook và khởi tạo toàn bộ UserForm | PASS |
| Workbook không chứa tài khoản/token thử nghiệm | PASS |
| Detail, XML compatibility, core retry/resume | PASS |
| Lọc nguồn cố định theo phiên và xóa dữ liệu theo nguồn | PASS |
| Giao diện, phân trang, pause/resume/stop, tiến độ | PASS |
| JSON scheduler và retry | PASS |
| Date/PDF runtime, số hóa đơn đệm 0, đường dẫn Unicode và ký tự đặc biệt | PASS |
| PDF mẫu một trang và hóa đơn 61 dòng/bốn trang | Đã render và xem bố cục |
| Live bằng chính `ExecuteGdtListRequest` mới | Query 5/10/30 ngày: 200; SCO 5/10 ngày: 200; SCO 30 ngày: 400, giới hạn 10 |
| Live XML → giải nén → PDF | Thành công với một hóa đơn; dữ liệu tạm đã xóa |

Bài PDF runtime cuối được chạy riêng sau khi sửa harness truyền đường dẫn Unicode bằng `ChrW`. Báo cáo cuối tại `dist/pdf-date-verify/date-pdf-test/result.json`; các báo cáo test có sẵn trong `tests` được sao lưu và khôi phục, báo cáo mới nằm trong thư mục `regression` của bản build. Các kiểm tra live không lưu thông tin đăng nhập, CAPTCHA hoặc hóa đơn thật trong repository.
