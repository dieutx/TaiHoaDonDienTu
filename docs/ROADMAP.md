# Roadmap

Roadmap chỉ ghi nhận hướng cải thiện. Không item nào được coi là đã cam kết triển khai nếu chưa có Issue và phạm vi được maintainer chấp thuận.

## Pre-public security review

- [Ưu tiên cao] Sanitize giá trị giống JWT tại `template/App_Template.xlsm`, sheet `MENU`, ô `D7`; thu hồi/rotate phiên liên quan và build lại release.
- Rà soát định danh tại `MENU!E7` trước khi public.
- Xác nhận các MST/số điện thoại trong `LinkTraCuu` là dữ liệu tham chiếu công khai; thay bằng dữ liệu giả hoặc loại bỏ phần không cần thiết nếu không xác minh được.
- Thay các tham chiếu InvoiceGUID/PDF mẫu trong `src/modules/Module1.bas` bằng fixture giả nếu module này tiếp tục được phân phối.
- Quyết định LICENSE sau khi rà soát quyền đối với upstream.

## Stability

- Bổ sung fixture giả lập cho các nhánh retry 429/500/timeout.
- Bổ sung test cho response rỗng, `null` và thiếu trường.
- Rà soát helper PDF mẫu đang nằm trong `src/modules/Module1.bas` trước khi coi là tính năng chính thức.

## Compatibility

- Theo dõi thay đổi endpoint/header của hệ thống hóa đơn điện tử.
- Tạo fixture cho nhiều biến thể XML từ nhà cung cấp khác nhau.
- Ghi nhận ma trận tương thích Office 32-bit/64-bit và phiên bản Windows.

## User Experience

- Bổ sung screenshot đã sanitize cho luồng đăng nhập và tải hóa đơn.
- Cải thiện hướng dẫn xử lý lỗi ngay trên form mà không thay đổi API flow.
- Chuẩn hóa thông báo lỗi mạng, xác thực và parse response.

## Testing

- Tách fixture giả lập khỏi test script để dễ review.
- Bổ sung regression test cho date/time và locale khác nhau.
- Bổ sung kiểm tra tự động tên file release `TaiHoaDonDienTu_vX.X.X.xlsm`.

## Documentation

- Bổ sung video hoặc ảnh minh họa không chứa dữ liệu thật.
- Viết tài liệu schema các sheet đầu ra.
- Bổ sung FAQ cho người dùng kế toán.

## Developer Experience

- Viết hướng dẫn thiết lập self-hosted runner Excel an toàn.
- Chuẩn hóa báo cáo kiểm thử để dễ đọc trên Pull Request.
- Đánh giá tình trạng bản quyền/upstream và quyết định LICENSE phù hợp.
