# Bao cao kiem tra retry/UI

Ngay chay: 24/09/2026 09:50:47

| # | Tinh huong | Ket qua | Bang chung |
|---:|---|---|---|
| 1 | Thanh cong o lan dau | PASS_STATIC | Nhanh 2xx tra Success ngay. |
| 2 | 429 co Retry-After | PASS_STATIC | 429 retryable va uu tien header so. |
| 3 | 429 khong co header | PASS_STATIC | Backoff 2^n va jitter/cap. |
| 4 | 500 roi thanh cong | PASS_STATIC | 500 nam trong vong retry. |
| 5 | Timeout roi thanh cong | PASS_STATIC | Transport error/status 0 quay lai vong retry. |
| 6 | Het luot dua vao queue | PASS_STATIC | Exhausted retry duoc danh dau va queue khong trung. |
| 7 | Retry cuoi thanh cong | PASS_STATIC | List/detail/XML/relative/related deu co nhanh success. |
| 8 | Retry cuoi that bai | PASS_STATIC | That bai cuoi upsert Khong tai duoc. |
| 9 | 401/403 dung request | PASS_STATIC | Co auth toan phien chan request moi. |
| 10 | Upsert khong trung | PASS_STATIC | Tim khoa truoc khi ghi. |
| 11 | Bao cao mua/ban | PASS_STATIC | Direction dung Mua vao/Ban ra. |
| 12 | Context query/sco-query | PASS_STATIC | Queue giu nguon API va dinh danh. |
| 19 | Bang ky thoi gian tao lai luc tai | PASS_STATIC | taiHoaDon_Total tao lai bang ky thoi gian truoc vong lap; arrDate khong con phu thuoc AfterUpdate cua o ngay. |
| 20 | O ngay trong khong lam loi form | PASS_STATIC | txtDenNgay_AfterUpdate chi so sanh ngay khi ca hai o da co gia tri. |
| 21 | Loi chi tiet khong truy cap arrDate(k) | PASS_STATIC | Loi tai chi tiet ghi nhan ky hieu, so va ngay hoa don tu arrHDChiTiet(j), khong truy cap arrDate(k). |
| 22 | Tu dong mo rong buffer hoa don | PASS_STATIC | Co ham EnsureInvoiceBufferCapacity mo rong buffer dong, GetSleepDelayMs an toan va dat lai n luc tai. |
| 23 | Tien do theo so hoa don thuc te | PASS_STATIC | Danh sach chiem 0-10%; 88% tiep theo gom ca relative, related, chi tiet va XML. |
| 24 | Goi API hoa don lien quan | PASS_STATIC | Trang thai 2-5 goi relative va related; trang thai 6 chi goi related. |
| 25 | Header API lien quan | PASS_STATIC | Request dung token phien hien tai, action, end-point va khong luu cookie vao source. |
| 26 | Bao cao chuoi lien quan | PASS_STATIC | Ket qua relative thanh chuoi nhieu dong; related ghi rieng va chap nhan response rong. |
| 27 | Chuoi mau 18510 | PASS_RUNTIME | Mau 18510 tao 6 dong; loi API hien thi dung cot; cot cu da bo; FillDown chi tiet khong dung clipboard. |
| 28 | Dinh dang thong bao related | PASS_STATIC | Parse ca response query va sco-query thanh cau thong bao tieng Viet. |
| 33 | Loi API lien quan hien thi tai dong hoa don | PASS_STATIC | Relative ghi loi vao Chuoi hoa don lien quan; related ghi loi vao Thong tin lien quan sau lan thu cuoi. |
| 34 | Bo cot Co HD lien quan | PASS_STATIC | Thong tin lien quan chuyen sang cot 64; build xoa cot cu va khong con du lieu Co HD lien quan. |
| 35 | Khong chiem dung clipboard khi ghi chi tiet | PASS_STATIC | Du lieu chung duoc FillDown noi bo, khong Copy va khong xoa clipboard Windows. |
| 29 | Ngay ISO khong phu thuoc locale | PASS_STATIC | Tao ngay bang DateSerial/TimeSerial, khong ghep chuoi roi CDate theo locale. |
| 30 | Chi giu loi chua xu ly | PASS_STATIC | Tai thanh cong xoa dong loi cung khoa va danh lai STT; ap dung ca lan dau va retry. |
| 31 | Log chi tiet phan trang danh sach | PASS_STATIC | Log neu ky, nguon query/sco-query, trang, HTTP, so hoa don, tong luy ke, thoi gian va chan state lap. |
| 32 | Tam dung tiep tuc va dung an toan | PASS_STATIC | Hai nut runtime dieu khien co-operative; cac vong request va thoi gian cho deu kiem tra pause/stop. |
| 15 | Tieu de bao cao dung tieng Viet | PASS_STATIC | UniConvert duoc tai tao tu modMsgboxTV.bas; 17 tieu de khop retry-ui-expected-text.json |
| 16 | Literal Telex khong con sai | PASS_STATIC | Da decode 177 literal UniConvert; khong con chuoi sai da biet. |
| 17 | Script build parse duoc | PASS_STATIC | Moi script build parse duoc va khong co literal non-ASCII thieu BOM. |
| 18 | Khoi VBA can bang | PASS_STATIC | Khoi Sub/Function/If/With/Select/For/Do can bang trong 21 file nguon. |
| 13 | Mo lai khong repair | PASS_RUNTIME | Build da mo lai workbook trong Excel COM instance moi: PASS, 18 check, 0 check loi, file TaiHoaDonDienTu_v6.7.2.xlsm. |
| 14 | Compile VBAProject | PASS_RUNTIME | Excel khong co API compile; bang chung runtime: 5 UserForm instantiate thanh cong trong tien trinh rieng (frmCapNhatMK, frmDangNhap, frmDatePicker, frmTaiHoaDon, frmTrichXuatXML). |

Luu y: `PASS_STATIC` chi xac nhan nhanh code va cau hinh hien dien; `PASS_RUNTIME` lay tu tests/comparison-report.json do build/Build-Excel.ps1 tao ra. Muc 13 va 14 la bang chung Excel; muc 17 va 18 la bang chung tinh khong can Excel.
