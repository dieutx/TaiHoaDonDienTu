# Kiểm tra phiên bản tùy chọn

## Cấu trúc và mã nguồn

- `src/modules/modVersion.bas`: tên ứng dụng, `CURRENT_VERSION = "6.7.3"`, URL metadata của repo hiện tại.
- `src/modules/modUpdate.bas`: một HTTP GET có timeout, đọc JSON qua `JsonConverter` sẵn có, so sánh version và hiện Form. Mọi lỗi mạng/JSON đều bỏ qua để ứng dụng tiếp tục mở.
- `src/forms/frmUpdate.code.txt`: **toàn bộ code Form** để dán vào code window của `frmUpdate`. File dùng các escape Unicode để build không phụ thuộc code page của máy.
- `build/Build-Excel.ps1`: tạo `frmUpdate` với các control dưới đây và chèn `CheckForUpdate` vào `ThisWorkbook.Workbook_Open` trong workbook build. Không sửa core VBA hiện có.
- `update.json`: metadata dự kiến cho v6.7.3. Ngày và ghi chú cuối cùng phải được cập nhật khi Release v6.7.3 và asset đã tồn tại. Bản 6.7.3 không tự báo cập nhật khi metadata cùng version.

Nếu lắp thủ công bằng VBA Editor, import hai file `.bas`, tạo Form theo bảng, dán `frmUpdate.code.txt`, rồi thêm `CheckForUpdate` vào `Workbook_Open` hiện có. Template repo không có `Workbook_Open`; mã đầy đủ là:

```vb
Private Sub Workbook_Open()
    CheckForUpdate
End Sub
```

## Control của `frmUpdate`

Form: `(Name)=frmUpdate`, `Width=410`, `Height=375`, `StartUpPosition=1` (CenterOwner). Thanh tiêu đề dùng tên ứng dụng; `lblHeader` hiển thị tiêu đề tiếng Việt bằng Unicode.

| Loại | Name | Caption khi hiển thị | Left | Top | Width | Height |
|---|---|---|---:|---:|---:|---:|
| Label | lblHeader | Có phiên bản mới | 20 | 18 | 350 | 24 |
| Label | lblCurrentTitle | Phiên bản hiện tại: | 20 | 54 | 150 | 18 |
| Label | lblCurrent | 6.7.3 | 180 | 54 | 190 | 18 |
| Label | lblNewTitle | Phiên bản mới: | 20 | 81 | 150 | 18 |
| Label | lblNew | Từ metadata | 180 | 81 | 190 | 18 |
| Label | lblDateTitle | Ngày phát hành: | 20 | 108 | 150 | 18 |
| Label | lblDate | Từ metadata | 180 | 108 | 190 | 18 |
| Label | lblNotesTitle | Nội dung cập nhật: | 20 | 141 | 350 | 18 |
| Label | lblNotes | Danh sách từ JSON | 20 | 166 | 350 | 90 |
| Label | lblQuestion | Bạn có muốn tải phiên bản mới? | 20 | 268 | 350 | 20 |
| CommandButton | cmdDownload | TẢI BẢN MỚI | 20 | 300 | 160 | 30 |
| CommandButton | cmdContinue | TIẾP TỤC DÙNG BẢN HIỆN TẠI | 190 | 300 | 190 | 30 |

Đặt `WordWrap=True` cho `lblNotes`. Form không có checkbox hay dữ liệu lựa chọn được lưu. Nút tải chỉ mở URL Release trong trình duyệt; không tải, copy hoặc ghi đè file bằng VBA.

## Phát hành v6.7.3

1. Build bằng `build/Build-Excel.ps1 -Version '6.7.3' -OutputPath '.\dist\TaiHoaDonDienTu_v6.7.3.xlsm'` trên máy có Excel và quyền truy cập VBA project. Trong VBA Editor, xác nhận `ThisWorkbook.Workbook_Open`, `modVersion`, `modUpdate`, `frmUpdate`.
2. Trên GitHub repo `dieutx/TaiHoaDonDienTu`, tạo Release tag `v6.7.3`; upload đúng tên `TaiHoaDonDienTu_v6.7.3.xlsm`.
3. Khi Release đã tồn tại, điền ngày phát hành thật và bổ sung ghi chú của **mọi thay đổi đã merge** vào `update.json`; xác nhận URL asset hoạt động rồi push metadata lên nhánh `main`. Khi phát hành bản sau, tăng `CURRENT_VERSION` và metadata theo cùng quy trình.

File v6.7.2 đã phát hành chưa có mã kiểm tra phiên bản. Chỉ các workbook được build từ thay đổi này mới tự kiểm tra khi mở; không thể khiến file v6.7.2 đã tải tự thông báo nếu không phát hành lại file đó.

## Kiểm thử

Chạy `tests/Test-UpdateRuntime.ps1 -BuiltWorkbook '.\dist\TaiHoaDonDienTu_v6.7.3.xlsm'` trên máy có Excel. Script dùng bản sao tạm và HTTP server cục bộ, mở workbook với `Workbook_Open` bật, kiểm tra dữ liệu Form trước khi hiện cửa sổ, rồi xóa bản sao. Không gửi fixture lên GitHub.

- **Có bản mới:** trong bản sao workbook, tạm đặt `CURRENT_VERSION = "6.7.2"` và dùng JSON hợp lệ ghi 6.7.3. Mở file, kiểm tra Form và cả hai nút.
- **Không có bản mới:** với `CURRENT_VERSION = "6.7.3"` và JSON ghi 6.7.3, mở file không có Form.
- **Mất Internet:** ngắt mạng. Mở file, sau timeout ngắn ứng dụng vẫn dùng được.
- **JSON lỗi:** dùng JSON sai hoặc thiếu `version`. Mở file, không có Form, không có lỗi nghiêm trọng.
- **So sánh:** chạy `? CompareVersions("6.7.10", "6.7.9")` trong Immediate Window; kết quả là `1`.

Mã không đọc hay ghi Registry, AppData, folder ẩn, log hoặc định danh người dùng. Một HTTP GET vẫn khiến GitHub nhận địa chỉ IP ở tầng mạng như mọi request thông thường; VBA không thu thập hoặc lưu IP.
