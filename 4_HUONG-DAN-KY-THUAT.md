# TAPACO NHÂN SỰ – TIỀN LƯƠNG · HƯỚNG DẪN CÀI ĐẶT & VẬN HÀNH

Phiên bản 3.26 · 11/2026 — 5 trường hợp tính lương (HĐLĐ + BH, khoán, BH nơi khác, nghỉ hưu, thử việc), trạng thái tính thuế theo thời điểm, Điều 168.3 BLLĐ, quyết toán TNCN năm và gợi ý giảm thuế; quy định lương, BHXH, thuế TNCN theo thời kỳ (trước / sau 01/7/2026), bảng chấm công thêm giờ, so sánh thuế 2026; thêm hồ sơ, điều động lưu nền; công văn OCR rồi AI phân tích; sửa hồ sơ, quá trình công tác lưu nền không phải chờ; tin nhắn hỏi nhanh từ bộ nhớ đệm; bảng chấm công ngày nhóm theo bộ phận, tự lưu chạy nền không phải chờ; báo cáo bảng lương lọc theo công ty, tách nhóm công ty → bộ phận; khai báo CBCNV theo công ty thành viên (TAPACO, Thịnh Vượng Toàn Cầu, An Thuận Phát…), hạch toán BHXH – lương – thuế theo công ty, Tổng quan thống kê theo công ty và theo phòng ban – công trường; tiền làm thêm giờ miễn thuế TNCN theo Luật 109/2025 & NĐ 253/2026, không tính đóng BH; AI đọc ảnh CCCD, đề xuất điền lý lịch, TCHC soát rồi lưu; ảnh CCCD mặt trước / mặt sau ngay trong lý lịch; sửa lỗi nút Ảnh / Tệp trong menu (…) của Tin nhắn không mở hộp chọn tệp; yêu cầu báo cáo có mục "Bảng số lượng" chọn sẵn gợi ý nhân lực, máy thi công, công việc; bảng phân quyền tích chọn được (Giám đốc, Quản trị); phần quản lý của Phân tích nhóm (AI) và Chấm công GPS chỉ dành cho Giám đốc, TCHC, Kế toán; AI đọc bản quét công văn và điền hộ vào sổ, Giám đốc duyệt chia sẻ văn bản, trao đổi xử lý ngay dưới mỗi văn bản; sổ công văn đến – đi (đăng ký vào sổ, bút phê giao xử lý có hạn, báo cáo kết quả, lưu hồ sơ) và đề xuất nội bộ (trình tự duyệt cấu hình riêng theo từng loại, người theo dõi – người thực hiện, đính kèm ảnh và tệp, toàn bộ lưu trên Drive Công ty); trợ lý quy định công ty trong Tin nhắn; đánh giá tổng hợp CBCNV (chấm công + vị trí + nhóm công việc + AI nhận xét, chốt, công bố, phản hồi); kiểm tra vị trí khi nhắn tin (ngoài phạm vi chấm công → sheet Kiểm tra công, báo TCHC, Ban Giám đốc, Kế toán); kiểm tra chấm công GPS bằng AI (cảnh báo đỏ/vàng, nhắc bộ phận, giải trình); chấm công bằng điện thoại (GPS + khuôn mặt, bản đồ, điểm riêng, công tác); thông báo đẩy như Zalo; lưu trữ và phân tích nhóm chat bằng AI; tin nhắn kiểu Zalo; bảng tin; lương, nghỉ phép, ý kiến CBCNV. Bản dành cho người dùng: xem "Hướng dẫn cài đặt" và "Hướng dẫn sử dụng" (Word / PDF) trong gói.

## 1. Thành phần gói triển khai

| File | Vai trò |
|---|---|
| `index.html` | Ứng dụng (PWA một file): giao diện cho mọi vai trò, chạy trên điện thoại và máy tính |
| `Code.gs` | Máy chủ Google Apps Script: đăng nhập, phân quyền, tính lương, ghi Google Sheets |
| `manifest.webmanifest`, `sw.js`, `icon-192.png`, `icon-512.png` | Cài app lên màn hình điện thoại, chạy toàn màn hình |
| `HUONG-DAN.md` | Tài liệu này |
| `PROMPT-TAO-LAI-APP-NHAN-SU-TAPACO.md` | Master prompt để dựng lại toàn bộ hệ thống |

Mở `index.html` khi chưa cấu hình máy chủ sẽ vào **chế độ dùng thử** với dữ liệu mẫu (lưu trên trình duyệt). Tài khoản mẫu: `quantm` (Giám đốc), `landt` (TCHC), `hant` (KT tiền lương), `huongpt` (KT công nợ), `maibt` (Trưởng phòng Kỹ thuật), `thanglv` (CHT công trình Cầu Hòa Lạc), `anhvd`, `baopq`, `tunghv` (Nhân viên), `admin` — mật khẩu `123456`. Kỳ 08/2026 đã duyệt, kỳ 09/2026 đang ở bước Chấm công để đi thử cả quy trình.

## 2. Cài máy chủ Google Sheets + Apps Script (khoảng 10 phút)

1. Tạo một Google Sheet mới bằng tài khoản Google của Công ty, đặt tên ví dụ **TAPACO_NhanSu_TienLuong**.
2. Menu **Tiện ích mở rộng → Apps Script**. Xóa nội dung mặc định, dán toàn bộ `Code.gs`, bấm **Lưu**.
3. **Cài đặt dự án** (biểu tượng bánh răng): Múi giờ = `(GMT+07:00) Asia/Ho_Chi_Minh`.
4. Chọn hàm **setup** ở thanh công cụ → **Chạy** → cấp quyền khi được hỏi.
   - Hàm tạo 14 sheet: `Users, NhanVien, QuaTrinh, LenLuong, KhenThuong, NghiPhep, TaiLieu, ChamCong, BangLuong, KyLuong, YKien, PheDuyet, CauHinh, NhatKy`.
   - Tạo thư mục Google Drive **"TAPACO – Tài liệu hồ sơ nhân sự (KHÔNG CHIA SẺ)"** để chứa file đính kèm. Khi được hỏi, cấp quyền truy cập Google Drive.
   - Nâng cấp từ bản cũ: chạy hàm `setup` một lần nữa (không xóa dữ liệu) để tạo sheet TaiLieu, tạo thư mục và cấp quyền Drive.
   - Tạo tài khoản quản trị **admin / Tapaco@2026** (bắt buộc đổi mật khẩu khi đăng nhập lần đầu).
5. **Triển khai → Tùy chọn triển khai mới → Ứng dụng web**:
   - Thực thi với tư cách: **Tôi**
   - Người có quyền truy cập: **Bất kỳ ai**
   - Bấm Triển khai, sao chép **URL ứng dụng web** (dạng `https://script.google.com/macros/s/…/exec`).

> "Bất kỳ ai" chỉ cho phép gọi API. Mọi thao tác đều phải đăng nhập, phân quyền kiểm tra tại máy chủ. **Không chia sẻ file Google Sheet** cho CBCNV: Sheet chứa toàn bộ dữ liệu lương.

## 3. Đưa ứng dụng lên mạng

Cách khuyến nghị (giống các app TAPACO trước): đưa 6 file `index.html, manifest.webmanifest, sw.js, icon-192.png, icon-512.png, webapp.txt` lên **GitHub Pages** hoặc **Netlify** (bắt buộc HTTPS).

**Gắn URL máy chủ – tệp `webapp.txt` (v3.16, khuyến nghị):**
1. Apps Script → **Triển khai → Quản lý bản triển khai** → sao chép **URL ứng dụng web** (dạng `https://script.google.com/macros/s/…/exec`; tài khoản Google Workspace có dạng `…/a/macros/<tên miền>/s/…/exec`).
2. Mở `webapp.txt` (Notepad), dán URL vào một dòng riêng, lưu. Dòng bắt đầu bằng `#` là ghi chú.
3. Đưa `webapp.txt` lên cùng thư mục với `index.html`.
- Mở app lần đầu: app tự đọc `webapp.txt` và kết nối máy chủ, người dùng không phải dán URL.
- **Đổi máy chủ** (tạo bản triển khai mới có URL khác, chuyển sang Google Sheet khác…): chỉ sửa `webapp.txt` rồi đưa lên lại. Mỗi máy tự nhận URL mới khi mở app (hoặc khi quay lại app sau 15 phút), báo "Máy chủ đã đổi…" và yêu cầu đăng nhập lại.
- `webapp.txt` chưa có URL hợp lệ hoặc không đọc được (mất mạng): app giữ URL đang dùng.
- Lưu ý: "Triển khai → Quản lý bản triển khai → Sửa → **Phiên bản mới**" giữ nguyên URL, không cần sửa `webapp.txt`. Chỉ khi chọn "Triển khai mới" (URL mới) mới cần sửa.

**Mở `index.html` trực tiếp từ ổ đĩa / thư mục Google Drive trên máy tính (file://)** (v3.17.1): trình duyệt (Chrome, Edge) **không cho trang web tự đọc tệp bên cạnh**, nên app không đọc được `webapp.txt` – màn hình đăng nhập hiện khung vàng "Đang mở trực tiếp từ ổ đĩa". Hai cách:
- Bấm **Chọn tệp webapp.txt để kết nối máy chủ** → chọn tệp webapp.txt trong thư mục 1_UngDung. Chỉ cần làm 1 lần trên mỗi máy / trình duyệt.
- Hoặc chạy công cụ **0_CongCu/GanURL_MayChu.bat** (bấm đúp): công cụ đọc `1_UngDung/webapp.txt` và ghi URL vào dòng `DEFAULT_API` của `1_UngDung/index.html` (giữ bản cũ `index.html.bak`). Sau đó mở `index.html` ở máy nào cũng tự kết nối. Đổi URL trong webapp.txt thì chạy lại công cụ.
- Mở từ ổ đĩa chỉ để thử: không cài được lên màn hình chính, không có thông báo đẩy, camera / định vị (chấm công GPS) có thể bị chặn. Dùng thật thì đưa lên GitHub Pages / Netlify (https).

Cách cũ vẫn dùng được: sửa dòng `const DEFAULT_API = '';` trong `index.html`; hoặc trên từng máy: màn hình đăng nhập → **Cấu hình máy chủ** → dán URL (có `webapp.txt` thì tệp này được ưu tiên ở lần mở app sau).

Cài lên điện thoại: mở link bằng Chrome (Android) hoặc Safari (iPhone) → **Thêm vào màn hình chính**.

## 4. Thiết lập ban đầu (Quản trị)

1. Đăng nhập `admin`, đổi mật khẩu.
2. **Cài đặt**: kiểm tra tên Công ty, lương cơ sở, lương tối thiểu vùng, tỷ lệ BH, biểu thuế, công chuẩn, chu kỳ nâng bậc, danh mục phụ cấp, chức vụ. Phòng ban và công trình khai ở màn hình **Phòng ban & công trình** (mục 4h). Nhập **thang bảng lương** của Công ty (nếu có).
3. **Tài khoản**: sau khi có hồ sơ CBCNV, cấp tài khoản và gán chức danh cho Giám đốc, TCHC, Kế toán tiền lương, Kế toán công nợ, các Trưởng phòng và Chỉ huy trưởng (mục 4i).
   **Cài đặt → Danh mục vi phạm**: nhập các vi phạm và mức khấu trừ đúng theo Nội quy lao động / Quy chế lương thưởng đã ban hành.
4. **Hồ sơ CBCNV → Nhập Excel**: tải file mẫu, điền danh sách hiện có (mã NV, họ tên, phòng ban, chức vụ, ngày vào, loại HĐ, lương chính, **ngày hưởng bậc hiện tại**, phụ cấp, số người phụ thuộc, ngân hàng…), nhập vào.
   - "Ngày hưởng bậc hiện tại" là mốc tính hạn nâng bậc. Để trống thì hệ thống lấy ngày vào làm, khiến người làm lâu năm hiện là "quá hạn".
5. **Tài khoản → Cấp hàng loạt cho CBCNV**: tạo tài khoản chức danh Nhân viên cho mọi người chưa có.
   - Tên đăng nhập = tên + chữ đầu họ đệm + ngày sinh dạng ddmmyy (Lê Văn Thắng, sinh 20/11/1990 → `thanglv201190`; trùng thì thêm số).
   - Mật khẩu ban đầu = tên + năm sinh (→ `thang1990`). Hồ sơ thiếu ngày sinh thì dùng mật khẩu mặc định trong Cài đặt.
   - Sau khi tạo, app hiện danh sách tên đăng nhập và mật khẩu ban đầu (xuất được Excel) để gửi riêng cho từng người.
   - Mỗi người phải đổi mật khẩu riêng ở lần đăng nhập đầu, sau đó tự đổi được bất cứ lúc nào (menu tài khoản → Đổi mật khẩu).
6. **Mã nhân viên** tự sinh theo dạng NV0001, NV0002… khi thêm hồ sơ (để trống ô Mã) hoặc nhập Excel có ô Mã NV để trống.
7. **Đặt lại mật khẩu** cho người khác (quên mật khẩu): chỉ **Quản trị** hoặc **Giám đốc**.
   - Vào Tài khoản → bấm người cần đặt lại → **Đặt lại mật khẩu**.
   - Chọn về mặc định (tên + năm sinh) hoặc nhập mật khẩu mới; người dùng phải đổi lại khi đăng nhập.
   - Giám đốc không đổi được mật khẩu tài khoản Quản trị. TCHC tạo tài khoản được nhưng không đổi mật khẩu của người khác.
   - Lưu ý: mật khẩu mặc định đoán được từ hồ sơ, nên bắt buộc đổi ở lần đăng nhập đầu.

## 4a2. Chấm dứt hợp đồng: tự khóa tài khoản, giới hạn chấm công (v3.18)

**Khai báo:** Quá trình công tác → **+ Thêm** → loại **Nghỉ việc** hoặc **Thanh lý hợp đồng** (mới) → **Từ ngày = ngày chấm dứt HĐ**: *kể từ ngày này không còn làm việc*. Ghi số quyết định, người ký. (Hoặc sửa hồ sơ: ô *Ngày chấm dứt HĐ (nghỉ việc)*.)

| Thời điểm | Hồ sơ | Tài khoản đăng nhập | Bảng chấm công |
|---|---|---|---|
| Ngày chấm dứt đã qua / hôm nay | chuyển **Đã nghỉ việc** | **khóa ngay**, phiên đang mở bị chặn ở thao tác kế tiếp; đăng nhập báo "Tài khoản đã bị khóa (nghỉ việc từ dd/mm/yyyy)" | kỳ có ngày chấm dứt: **vẫn có tên** (nhãn đỏ "Chấm dứt HĐ từ dd/mm"), chỉ chấm được **các ngày trước** ngày chấm dứt; từ kỳ sau: không còn tên |
| Khai trước (ngày chấm dứt sắp tới) | vẫn Đang làm việc, hiện "sắp tới" | vẫn dùng được, **đến ngày tự khóa** | như trên khi đến kỳ đó |

- Lưu bảng chấm công: máy chủ tự bỏ các ô từ ngày chấm dứt (và trước ngày vào làm). Nút "Điền đủ công" ở tab Tổng hợp chỉ điền số ngày làm việc trong thời hạn HĐ.
- **Nhập nhầm:** xóa bản ghi Nghỉ việc / Thanh lý → hồ sơ và tài khoản tự mở lại. **Vào làm lại:** thêm "Tiếp nhận" sau ngày chấm dứt → mở lại.
- Trang **Tài khoản** hiện "Đã khóa" + lý do "Tự động – nghỉ việc từ …" (hoặc "Khóa từ dd/mm/yyyy" nếu khai trước). Quản trị vẫn mở tay được (ví dụ cho người đã nghỉ xem phiếu lương cuối) – mở tay thì không bị khóa lại; khóa tay của Quản trị không bị tự mở.
- Tài khoản **Quản trị** không bao giờ tự khóa (tránh mất quyền quản trị hệ thống).
- Hồ sơ đã "Đã nghỉ việc" từ trước bản 3.18 mà tài khoản còn mở: lần đầu Quản trị / TCHC mở app sau nâng cấp, các tài khoản này được khóa (trừ tài khoản Quản trị đã mở tay).
- Lưu ý: từ v3.18 "ngày nghỉ việc" là **ngày chấm dứt** (ngày đầu tiên không còn làm việc). Hồ sơ cũ đã ghi ngày nghỉ là *ngày làm việc cuối cùng* thì sửa thành ngày hôm sau.

## 4a3. Định dạng ngày dd/mm/yyyy; một thư mục dữ liệu trên Drive (v3.19)

**Ngày:** mọi ô ngày nhập và hiển thị theo **dd/mm/yyyy** trên mọi máy (trước đây ô ngày theo ngôn ngữ của trình duyệt, máy cài tiếng Anh hiện mm/dd/yyyy).
- Gõ số liền: `05112026` → tự thành `05/11/2026`; gõ gọn `1/3/26` → `01/03/2026`. Bấm 📅 để chọn trên lịch.
- Ngày không có thật (31/02…) → ô báo đỏ, không nhận.
- Chữ ngày dạng `yyyy-mm-dd` hiện trên màn hình (giờ tải lên, thời gian ghi…) tự đổi sang `dd/mm/yyyy`. Dữ liệu trong Google Sheet vẫn lưu `yyyy-mm-dd` (để sắp xếp, tính toán đúng) – không phải sửa sheet.

**Thư mục dữ liệu trên Google Drive:** mọi tệp của app nằm trong **một thư mục mẹ** đặt **cùng thư mục với Google Sheet** (thư mục chứa app):

```
<thư mục chứa Google Sheet>/
  TAPACO Nhân sự – Dữ liệu app (KHÔNG CHIA SẺ)/
    ├─ Tài liệu hồ sơ nhân sự (KHÔNG CHIA SẺ)/   hồ sơ CBCNV, ảnh CCCD, ảnh chấm công GPS, tệp tin nhắn, bảng tin, công văn, đề xuất…
    └─ Lưu trữ nhóm chat (KHÔNG CHIA SẺ)/        lưu trữ nhóm chat, báo cáo phân tích AI (Docs, PDF, Excel)
```

- Bản ≤ 3.18 có 2 thư mục mẹ riêng ("TAPACO – Tài liệu hồ sơ nhân sự" ở *Drive của tôi* và "TAPACO – Lưu trữ nhóm chat" cạnh Sheet). Sau nâng cấp, lần đầu app dùng đến tệp, hai thư mục này được **chuyển nguyên vẹn** vào thư mục mẹ (giữ tên cũ, giữ mã tệp – mọi liên kết, tệp đính kèm cũ vẫn mở được). Muốn gom ngay: Apps Script → chọn hàm **gomDuLieuDrive** → **Chạy** → xem nhật ký thực thi có đường dẫn thư mục mẹ.
- Chuyển thư mục chứa Google Sheet sang chỗ khác sau này: thư mục dữ liệu không tự đi theo – kéo thư mục "TAPACO Nhân sự – Dữ liệu app" sang cùng chỗ (app nhận theo mã thư mục, không theo đường dẫn).

## 4a4. Công ty thành viên: khai báo nhân sự theo công ty, hạch toán BHXH – lương – thuế riêng (v3.20)

**Danh mục:** *Phòng ban & công trình → tab Công ty thành viên* (Quản trị, TCHC). Mặc định 3 công ty: **TPC – TAPACO**, **TC – Thịnh Vượng Toàn Cầu**, **ATP – An Thuận Phát**. Mã số thuế, địa chỉ, tên đầy đủ (TNHH / Cổ phần…) **để trống – Anh/Chị bấm "Sửa" điền đúng theo giấy phép**. Thêm công ty mới bằng "+ Công ty"; công ty không còn hoạt động đặt trạng thái "Ngừng" (không xóa được công ty còn hồ sơ hoặc đã có bảng lương).

**Khai báo cho CBCNV:**
- Hồ sơ CBCNV có ô **Công ty thành viên** (công ty ký HĐLĐ). Hồ sơ cũ chưa khai báo được tính vào công ty đầu danh sách (TAPACO), hiện chữ "(mặc định)".
- Gán nhanh nhiều người: tab Công ty thành viên → **Gán công ty cho CBCNV** → chọn công ty, tích người (có "chọn cả nhóm") → Gán.
- Nhập Excel: cột **Công ty** ghi mã (ATP) hoặc tên (An Thuận Phát, không dấu cũng được). Xuất Excel có cột này.
- Danh sách CBCNV có cột Công ty và bộ lọc "Tất cả công ty".

**Hạch toán:** lúc tính lương, công ty của từng người được **chốt vào bảng lương** (cột `congTy` sheet BangLuong). Chuyển công ty sau đó không làm đổi các kỳ đã tính; kỳ tiếp theo hạch toán theo công ty mới. Kỳ tính trước bản 3.20 lấy theo hồ sơ hiện tại.
- *Báo cáo lương*: ô chọn công ty ở góc trên – mọi báo cáo (bảng lương tháng, năm, bảng tính chi tiết, chi phí lương & BH) và file Excel xuất riêng cho công ty đó, tiêu đề in tên công ty. Để "Tất cả công ty" thì tab Chi phí lương & BH có thêm bảng theo công ty; file Excel có sheet **TheoCongTy** (SUMIFS) và cột Công ty ở sheet DuLieu.

**Báo cáo bảng lương – ô lọc công ty (v3.21)**, góc trên màn hình Báo cáo lương:
- *Tất cả công ty – tách theo công ty, bộ phận* (mặc định): nhóm từng công ty (TAPACO → Thịnh Vượng Toàn Cầu → An Thuận Phát theo thứ tự danh mục), trong mỗi công ty nhóm theo phòng ban / công trường; có dòng cộng bộ phận, **cộng công ty**, tổng cộng các công ty. File Excel giữ đúng cách nhóm (công thức SUBTOTAL lồng nhau, không cộng trùng); danh sách chuyển khoản thêm cột Công ty.
- *Toàn bộ các công ty – gộp theo bộ phận*: như trước 3.20, chỉ nhóm theo bộ phận.
- *Từng công ty*: chỉ người của công ty đó, tiêu đề in tên công ty – dùng để ký, nộp riêng từng pháp nhân.

**Tổng quan (Ban Giám đốc, TCHC, Kế toán):**
- Thẻ từng công ty: số đang làm việc, đã nghỉ, quỹ lương kỳ chọn.
- Bảng **theo công ty**: đang làm việc (số thử việc), đã nghỉ việc (số nghỉ trong năm), tiếp nhận trong năm, khen thưởng, kỷ luật (quyết định đã duyệt trong năm), quỹ lương, BHXH-YT-TN người lao động, BHXH-YT-TN + KPCĐ doanh nghiệp, thuế TNCN, tổng chi phí; dòng tổng cộng.
- Bảng **theo phòng ban, công trường** – gộp các công ty (cột "Theo công ty" cho biết mỗi bộ phận có bao nhiêu người của công ty nào), kèm quỹ lương, BH, thuế TNCN từng bộ phận. Ô chọn bên phải lọc riêng 1 công ty.
- Ô chọn kỳ: từng tháng hoặc **cả năm**. Bấm thẻ / dòng công ty hoặc dòng bộ phận để mở danh sách CBCNV tương ứng.

## 4b. Tài liệu đính kèm hồ sơ

### Hồ sơ đính kèm ngay trong lý lịch (v3.17)

Ở **Hồ sơ CBCNV → tab Thông tin** (và **Hồ sơ của tôi**), dưới khối ảnh CCCD có mục **Hồ sơ đính kèm** gồm 3 nhóm:

| Nhóm | Loại tài liệu | Thư mục trên Google Drive Công ty |
|---|---|---|
| Hợp đồng, quyết định | Hợp đồng lao động, Phụ lục hợp đồng, Quyết định | `<mã – họ tên>/1. Hợp đồng – Quyết định` |
| Lý lịch | Sơ yếu lý lịch | `<mã – họ tên>/2. Lý lịch` |
| Hồ sơ cá nhân | các loại còn lại: CCCD, MST, sổ BHXH, bằng cấp – chứng chỉ, giấy khám sức khỏe, người phụ thuộc, khác… | `<mã – họ tên>/3. Hồ sơ cá nhân` |

- Bấm **+ Tải lên** ở nhóm cần thêm → chọn loại, số hiệu, ngày ký / cấp, hạn → chọn **một hoặc nhiều tệp** (PDF, Word .doc/.docx, Excel .xls/.xlsx, ảnh JPG/PNG/HEIC; tối đa 10 MB/tệp – sửa ở Cài đặt; ảnh lớn tự thu nhỏ). Nhiều tệp thì mỗi tệp thành một dòng, cùng loại / ngày / số hiệu, tên lấy theo tên tệp.
- Tệp lưu **riêng tư** trong thư mục "TAPACO – Tài liệu hồ sơ nhân sự (KHÔNG CHIA SẺ)" trên Google Drive của tài khoản chạy Apps Script, chia theo CBCNV và nhóm. Xem / tải về qua app theo phân quyền, mỗi lần xem đều ghi nhật ký. Người có quyền Tài liệu bấm **Mở thư mục trên Drive** (chỉ tài khoản quản lý Drive của Công ty mở được).
- **TCHC** tải mọi loại, có hiệu lực ngay. **CBCNV** tự tải các loại trong *Loại CBCNV được tự tải lên* (lý lịch, hồ sơ cá nhân) – chờ TCHC xác nhận; hợp đồng, quyết định do TCHC cập nhật.
- Tab **Tài liệu** vẫn có đủ danh sách, giấy tờ bắt buộc, hạn hiệu lực, xác nhận, sửa, xóa như trước. Tệp tải trước bản 3.17 nằm ngay trong thư mục của CBCNV (không chia nhóm).

### Ảnh căn cước công dân trong lý lịch (v3.10)

Ở **Hồ sơ CBCNV → tab Thông tin** (và **Hồ sơ của tôi**), ngay dưới *Thông tin cá nhân* có khối **Ảnh căn cước công dân** gồm hai ô **Mặt trước** và **Mặt sau**:
- Bấm **📷 Tải ảnh lên** (trên điện thoại chọn chụp ảnh hoặc lấy từ thư viện). Ảnh được thu nhỏ (≤ 1600 px, JPG) rồi lưu thành tài liệu loại **Căn cước công dân**, tên "CCCD – mặt trước / mặt sau", tự ghi kèm **số CCCD** và **ngày cấp** đang có trong lý lịch. Ảnh hiện ngay trong ô; bấm vào ảnh để xem lớn, tải về.
- **TCHC (người có quyền sửa hồ sơ)** tải cho mọi CBCNV, ảnh có hiệu lực ngay; bấm **Thay ảnh** thì ảnh cũ cùng mặt bị thay.
- **CBCNV** tự tải cho hồ sơ của mình (khi loại "Căn cước công dân" nằm trong danh sách *Loại CBCNV được tự tải lên*); ảnh hiện nhãn **Chờ TCHC xác nhận**. Tự tải lại khi ảnh cũ còn đang chờ thì ảnh cũ được thay; ảnh đã được TCHC xác nhận thì giữ nguyên để TCHC xét.
- **AI đọc ảnh, đề xuất điền lý lịch (v3.11)** — người có quyền sửa hồ sơ (TCHC) bấm nút này khi đã có ít nhất một mặt ảnh. AI đọc cả hai mặt và mở **hộp soát** gồm ảnh thẻ bên trái để đối chiếu, bên phải là bảng 4 cột: *Mục – Đang có trong lý lịch – AI đọc được (sửa được) – Cập nhật*. Các mục: số CCCD, ngày sinh, giới tính, ngày cấp, nơi cấp, có giá trị đến, quê quán, địa chỉ (nơi thường trú); họ tên chỉ để đối chiếu.
  - AI **chỉ ghi chữ in trên thẻ mà đọc rõ**; chỗ mờ, bị che hoặc không có thì để trống, không đoán (ví dụ chỉ có ảnh mặt trước thì ngày cấp, nơi cấp để trống). Số không đủ 12 chữ số, ngày không hợp lệ hoặc ở tương lai thì không đề xuất.
  - Ô **tích sẵn** chỉ khi lý lịch đang trống. Ô AI đọc **khác** lý lịch tô đỏ nhạt, ghi "Khác lý lịch – soát kỹ" và **không tích sẵn**. Sửa được chữ AI đọc trước khi lưu; gõ vào ô trống thì tự tích.
  - **Cảnh báo tự động** (không cần AI): số CCCD 12 số có chữ số thứ 4 cho biết giới tính và thế kỷ sinh (0/1: nam/nữ sinh 1900–1999; 2/3: nam/nữ sinh 2000–2099), chữ số 5–6 là 2 số cuối năm sinh → app đối chiếu với ngày sinh, giới tính đọc được; ngày cấp trước ngày sinh; thẻ đã hết hạn; họ tên trên thẻ khác hồ sơ (có thể nhầm ảnh người khác).
  - Bấm **Lưu vào lý lịch**: chỉ ô đã tích được ghi; hạn thẻ ghi vào tài liệu ảnh CCCD để app cảnh báo khi sắp hết hạn. Nhật ký ghi ai cập nhật những ô nào, **không ghi số liệu cá nhân**. CBCNV không có nút này.
  - **Lưu ý dữ liệu cá nhân:** ảnh thẻ được gửi tới nhà cung cấp AI đã cài (Gemini hoặc Claude) để đọc. Công ty nên nêu việc này trong quy định / thông báo xử lý dữ liệu cá nhân với CBCNV. Bản dùng thử dùng AI mô phỏng, không đọc ảnh thật – các ô là dữ liệu MẪU.
- Ai xem được: người có quyền **Xem tài liệu đính kèm của mọi CBCNV** (Bảng phân quyền) và chính người đó. Người khác thấy "Đã có ảnh – bạn không có quyền xem". Ảnh nằm riêng tư trên Drive Công ty, có trong tab **Tài liệu**, mỗi lần xem đều ghi nhật ký; đăng xuất là xóa ảnh khỏi bộ nhớ của trình duyệt.

Mỗi hồ sơ CBCNV có tab **Tài liệu**. Loại tài liệu mặc định: hợp đồng lao động, phụ lục, căn cước công dân, mã số thuế, quyết định, sổ BHXH, bằng cấp – chứng chỉ, sơ yếu lý lịch, giấy khám sức khỏe, hồ sơ người phụ thuộc, khác. Danh sách này sửa được trong Cài đặt.

**Tải lên**
- Chọn file hoặc kéo thả vào ô tải lên. Nhận PDF, ảnh chụp (JPG, PNG, HEIC), Word, Excel, tối đa 10 MB mỗi file.
- Trên điện thoại có thể chụp ảnh trực tiếp. Ảnh lớn được tự thu nhỏ trước khi gửi.
- Nhập thêm: số hiệu (số HĐ, số QĐ, số CCCD…), ngày ký hoặc ngày cấp, ngày hết giá trị.
- **Gắn với** một quá trình công tác, quyết định lương hoặc khen thưởng – kỷ luật. Tài liệu gắn vào hiện biểu tượng 📎 ngay trên dòng thời gian quá trình công tác.
- TCHC có thể đánh dấu **ẩn với CBCNV** cho từng tài liệu, ví dụ biên bản nội bộ.

**CBCNV tự tải lên**
- CBCNV vào **Hồ sơ của tôi → Tài liệu** để tự tải các loại được phép (CCCD, mã số thuế, bằng cấp, giấy khám sức khỏe, hồ sơ người phụ thuộc…).
- Tài liệu CBCNV tự tải ở trạng thái **chờ TCHC xác nhận**. Số trên menu "Hồ sơ CBCNV" báo số tài liệu đang chờ.
- Hợp đồng và quyết định chỉ TCHC tải lên.

**Rà soát giấy tờ** (Hồ sơ CBCNV → tab Rà soát giấy tờ) cho biết:
- người còn thiếu giấy tờ bắt buộc (mặc định: hợp đồng lao động và CCCD; đổi được trong Cài đặt);
- tài liệu đang chờ xác nhận;
- giấy tờ hết hạn hoặc sắp hết hạn trong 30 ngày. Danh sách sắp hết hạn cũng hiện ở Tổng quan.

**Bảo mật**
- File lưu trong thư mục Drive riêng của tài khoản Google chạy Apps Script, mỗi CBCNV một thư mục con. File **không có link chia sẻ**. **Không chia sẻ thư mục này cho ai**: muốn xem tài liệu phải đăng nhập app.
- Ai được xem:
  - Quản trị, Giám đốc, TCHC, Kế toán lương: xem tài liệu của mọi người. Kế toán công nợ không xem được.
  - CBCNV: chỉ xem tài liệu của chính mình, trừ tài liệu bị ẩn.
- Mỗi lần mở xem tài liệu được ghi vào Nhật ký: ai xem, tài liệu nào, lúc nào.
- Xóa tài liệu trong app thì file trên Drive được chuyển vào thùng rác của Drive, còn khôi phục được trong 30 ngày.

> Bản dùng thử lưu file ngay trong trình duyệt nên chỉ chứa được vài chục MB. Các tài liệu mẫu đều ghi chữ "MẪU", không phải giấy tờ thật.

## 4c. Cài đặt quy trình & người phê duyệt (Quản trị, Giám đốc)

Menu **Hệ thống → Quy trình & phê duyệt** gồm 4 mục. Mỗi bước / cấp có thể giao cho:
- **một vai trò**: mọi tài khoản có vai trò đó, vd Tổ chức – Hành chính;
- **một người cụ thể**: vd Kế toán trưởng, Phó Giám đốc;
- **Trưởng phòng / Chỉ huy trưởng của bộ phận người đề nghị**: chỉ dùng cho duyệt nghỉ và ý kiến; app tìm theo tài khoản có chức danh TP/CHT gán cho bộ phận đó (mục 4i).

Một bước giao cho nhiều người thì ai trong số đó xử lý cũng được.

1. **Quy trình lập & phê duyệt lương**: đổi tên bước, đổi người xử lý, thêm bước *Kiểm tra / soát xét* (vd Kế toán trưởng soát xét trước khi trình Giám đốc), bỏ bước Công nợ hoặc Vi phạm nếu không dùng, sắp xếp lại thứ tự.
   - Bắt buộc: bước 1 là Chấm công, có bước Tính lương, bước cuối là Phê duyệt cuối. Giám đốc phê duyệt ngay được ở mọi bước.
   - Kỳ lương mới mở chạy theo quy trình đã lưu tại lúc mở kỳ. Đánh dấu "Áp dụng ngay cho các kỳ đang xử lý" để đổi cả kỳ đang làm dở; kỳ nào đang đứng ở bước vừa bị xóa thì giữ quy trình cũ.
   - Bước có nhập liệu lương (chấm công, tính lương, công nợ, vi phạm) chỉ giao cho tài khoản quản lý, không giao cho tài khoản CBCNV.
2. **Quy trình duyệt đơn nghỉ**: các cấp duyệt nối tiếp nhau. Mỗi cấp có điều kiện "áp dụng khi số ngày nghỉ > N" và có thể chỉ áp dụng cho một số loại nghỉ, vd nghỉ không lương luôn qua Giám đốc.
   - Cấp mà người duyệt chính là người nộp đơn được bỏ qua.
   - Giám đốc và Quản trị duyệt vượt cấp được.
   - Đơn đã nộp giữ quy trình tại lúc nộp.
3. **Tiếp nhận ý kiến về lương**: danh sách chủ đề CBCNV chọn khi gửi ý kiến, và người xử lý của từng chủ đề.
4. **Trưởng bộ phận (dự phòng)**: từ bản 2.3, người duyệt của bộ phận lấy từ tài khoản Trưởng phòng / CHT. Bảng này chỉ dùng khi bộ phận chưa có TP/CHT; chưa khai nữa thì chuyển TCHC.

Bản dùng thử: đơn nghỉ đi qua *Trưởng phòng / CHT → TCHC*. `thanglv` là CHT công trình Cầu Hòa Lạc, `maibt` là Trưởng phòng Kỹ thuật; `anhvd` có một đơn đang chờ CHT duyệt.

## 4g. Dùng trên điện thoại

- **Thanh điều hướng dưới cùng**, chạm bằng ngón cái:
  - CBCNV: Bảng tin · Tin nhắn · Phiếu lương · Nghỉ phép · Thêm.
  - Quản lý không gắn hồ sơ: Bảng tin · Tin nhắn · Tổng quan · Lương · Thêm.
  - Mỗi mục có số tin mới hoặc việc cần xử lý. "Thêm" mở menu đầy đủ.
- **Tin nhắn toàn màn hình**: khi mở một cuộc trò chuyện, thanh trên và thanh dưới được ẩn, khung soạn tin nằm sát đáy, ô nhập không làm điện thoại tự phóng to.
  - Nút **Quay lại** của điện thoại đóng cuộc trò chuyện chứ không thoát app.
  - Chạm vào tin để trả lời, thả cảm xúc, thu hồi.
- Cài lên màn hình chính: mở app bằng trình duyệt điện thoại → "Thêm vào màn hình chính" (Android: menu ⋮; iPhone: nút Chia sẻ). App mở như ứng dụng riêng, không có thanh địa chỉ.

**Tin nổi bật khi mở app**
- Khi đăng bài, tích **"Tin nổi bật: hiện nổi khi CBCNV mở app"**, có thể đặt **đến hết ngày** (bỏ trống thì hiện đến khi bỏ tích).
- Mỗi CBCNV mở app sẽ thấy cửa sổ nổi với ảnh, tiêu đề, tóm tắt và nút **Xem chi tiết** (với văn bản cần xác nhận là **Đọc và xác nhận**). Có nhiều tin nổi bật thì bấm "Tin tiếp"; "Để sau" thì lần mở sau hiện lại.
- Mỗi tin chỉ hiện một lần với mỗi người; sửa bài thì hiện lại.
- Đầu Bảng tin có **dải tin nổi bật và tin ghim** vuốt ngang, chạm để mở bài.

## 4f. Nhắn tin nội bộ

Menu **Công ty → Tin nhắn**. Mọi tài khoản đang hoạt động đều nhắn được cho nhau, cách dùng tương tự Zalo / Telegram.

**Trò chuyện riêng**
- Gõ tên đồng nghiệp vào ô tìm kiếm rồi bấm vào tên.
- Danh bạ lấy từ các tài khoản đang hoạt động, kèm chức vụ và phòng ban.

**Nhóm**
- Ai cũng tạo được nhóm: bấm **Tạo nhóm**, đặt tên, chọn thành viên. Có thể chọn cả phòng ban một lần.
- Người tạo là **trưởng nhóm**.
- Mọi thành viên đều thêm được người vào nhóm.
- Việc chỉ trưởng nhóm làm được:
  - đổi tên, mô tả nhóm;
  - mời thành viên ra khỏi nhóm;
  - đặt hoặc bỏ trưởng nhóm.
- Thành viên tự rời nhóm được. Nếu không còn trưởng nhóm, người vào nhóm sớm nhất được nhận quyền trưởng nhóm.
- Người tạo nhóm hoặc Quản trị **giải tán** được nhóm: lịch sử vẫn xem được, nhưng không gửi tin mới.
- Người đã rời hoặc bị mời ra không xem được nhóm nữa.

**Tin nhắn**
- Gõ chữ, chèn biểu tượng cảm xúc. **Enter** để gửi, **Shift+Enter** để xuống dòng.
- Gửi ảnh (tự thu nhỏ) và tệp PDF, Word, Excel, PowerPoint, ZIP, TXT, mỗi tệp ≤ giới hạn MB trong Cài đặt. Có thể dán ảnh hoặc kéo thả tệp vào khung chat.
- **Trả lời** trích dẫn một tin; bày tỏ cảm xúc (👍 ❤️ 😆 😮 😢 🙏); **thu hồi** tin của mình (ảnh, tệp đính kèm bị xóa).
- Trên điện thoại: chạm vào tin để hiện các thao tác trên.
- Theo dõi trạng thái:
  - "Đã gửi / Đã xem" khi chat riêng; "Đã xem: …" trong nhóm;
  - "đang soạn tin…";
  - chấm xanh khi đồng nghiệp đang mở app.
- Mỗi cuộc trò chuyện có thể **ghim** lên đầu hoặc **tắt thông báo**. Xem ảnh và tệp đã gửi trong "Thông tin".
- Thông báo tin mới:
  - số tin chưa đọc hiện trên menu và tiêu đề tab;
  - ô thông báo nhỏ khi đang ở màn hình khác;
  - thông báo của máy tính / điện thoại khi bấm **Bật thông báo trên máy** (lúc app đang mở ở nền).

**Kho tệp & lưu trữ trên Google Drive**
- Mọi ảnh, tài liệu gửi trong chat được lưu trên Google Drive của Công ty, mỗi cuộc trò chuyện một thư mục con trong "_TIN NHẮN NỘI BỘ":
  - nhóm: `Tên nhóm [mã]`;
  - chat riêng: `Riêng – Tên A & Tên B [mã]`.
- Tên tệp trên Drive có dạng `ngày_tên-tệp-gốc`. Đổi tên nhóm thì thư mục đổi tên theo.
- Nút **Kho tệp** (đầu danh sách tin nhắn, hoặc "Mở Kho tệp của cuộc trò chuyện" trong Thông tin nhóm) cho xem toàn bộ ảnh, tài liệu đã gửi trong các cuộc trò chuyện mình đang tham gia:
  - lọc theo cuộc trò chuyện, loại (ảnh / tài liệu), người gửi, tháng, tên tệp;
  - nhóm theo tháng, có tổng số tệp và dung lượng;
  - **Xem, Tải về, Đến tin nhắn** (mở cuộc trò chuyện và nhảy tới đúng tin đã gửi tệp).
- Thu hồi tin thì tệp bị chuyển vào thùng rác Drive và biến mất khỏi kho.
- Quản trị thấy thêm:
  - nút **Mở thư mục trên Google Drive** của từng cuộc trò chuyện;
  - nút **Xếp tệp cũ vào thư mục theo nhóm** (cho tệp gửi từ bản 2.0, mỗi lần 150 tệp).

**Cách hoạt động và giới hạn**
- Máy chủ Apps Script không giữ kết nối liên tục, nên app hỏi tin mới định kỳ:
  - khoảng 3 giây khi đang mở màn hình Tin nhắn;
  - 15 giây khi ở màn hình khác;
  - 45 giây khi app chạy nền.
- Tin nhắn vì vậy đến chậm vài giây. Cách này phù hợp quy mô nội bộ (vài trăm người).
- Từ bản 3.1, khi đã tắt app điện thoại vẫn nhận **thông báo đẩy** (Firebase Cloud Messaging) nếu Quản trị đã cấu hình và người dùng đã bật trên máy (mục 4s). Nội dung tin vẫn tải về theo chu kỳ hỏi tin ở trên khi mở app.
- Dữ liệu lưu ở đâu:
  - tin nhắn ở sheet **TinNhan**: chỉ ghi thêm, mỗi tin một dòng, số dòng là mã đồng bộ. **Không sắp xếp, xóa hay chèn dòng trong sheet này**;
  - danh sách nhóm và thành viên ở sheet **HoiThoai**, **ThanhVien**;
  - tệp ở thư mục Drive "_TIN NHẮN NỘI BỘ" (không chia sẻ), danh mục tệp ở sheet **TepChat**.
- Tin nhắn nằm trên hệ thống của Công ty: người quản lý file Google Sheet đọc được dữ liệu gốc. Cần thông báo điều này cho CBCNV và không gửi mật khẩu, số tài khoản, thông tin cá nhân nhạy cảm qua tin nhắn.

## 4e. Bảng tin công ty

Menu **Công ty → Bảng tin công ty** là trang đầu tiên của CBCNV sau khi đăng nhập. Bảng tin dùng để đăng thông báo, chính sách – quy định, tin nội bộ, truyền thông – hoạt động, vinh danh – khen thưởng.

**Chuyên mục**: Thông báo · Chính sách – Quy định · Tin nội bộ · Truyền thông – Hoạt động · Vinh danh – Khen thưởng.
**Bố cục theo màn hình (3.5.1)**: khung bảng tin giãn theo bề ngang màn hình (chặn ở 1800px trên màn siêu rộng, căn giữa); cột phải co giãn `clamp(300px, 21vw, 380px)`. Từ 1560px, danh sách bài viết (`.bt-feed`) chia **2 cột** để dòng chữ không quá dài; dưới mức đó và trên điện thoại vẫn 1 cột. Trước 3.5.1 khung bị chặn cứng 1180px nên màn 1920 bỏ trống khoảng 500px bên phải.



**Người được đăng bài**
- Quản trị, Giám đốc, Tổ chức – Hành chính đăng được bài.
- Sửa, xóa bài: người đăng, Giám đốc hoặc Quản trị.
- Ghim, bỏ ghim: người có quyền đăng bài.

**Soạn bài** (nút "+ Đăng bài")
- Nội dung: tiêu đề, nội dung (viết `**chữ đậm**`; đường link tự nhận).
- Ảnh: tự thu nhỏ ≤ 1600 px trước khi tải.
- Tệp đính kèm: PDF, Word, Excel.
- Văn bản chính sách: thêm số văn bản và ngày hiệu lực.
- Người nhận: toàn công ty hoặc chọn phòng ban / đơn vị. Người ngoài đối tượng không thấy bài và không mở được ảnh, tệp.
- Tùy chọn: **Ghim lên đầu**; **Yêu cầu CBCNV xác nhận đã đọc** (tự bật cho chuyên mục Chính sách – Quy định).
- Có thể **Lưu nháp**; bản nháp chỉ người có quyền đăng bài thấy.

**Phía CBCNV**
- Xem bài, phóng to ảnh, mở PDF ngay trong app.
- Bấm "Hữu ích", bình luận.
- Với văn bản yêu cầu xác nhận, bấm **Tôi đã đọc và hiểu**; app ghi lại thời điểm xác nhận.
- Menu có số đếm bài mới và văn bản chưa xác nhận. Cột phải có các mục: "Cần bạn xác nhận", "Đang ghim", "Chính sách – quy định".

**Theo dõi xác nhận**
- Người đăng bấm "Đã xác nhận x/y" trên bài để xem ai đã hoặc chưa xác nhận, và xuất Excel làm bằng chứng phổ biến văn bản.
- Danh sách tính trên CBCNV đang làm việc, có tài khoản, thuộc đối tượng nhận.

**Lưu trữ**: ảnh và tệp nằm trong thư mục Drive "_BẢNG TIN CÔNG TY" (bên trong thư mục tài liệu nhân sự, không chia sẻ). CBCNV chỉ mở được qua app.

## 4d. Khai báo mức hưởng theo thời gian

Mỗi CBCNV có **diễn biến mức hưởng**, mỗi lần thay đổi ghi kèm **ngày hiệu lực**:
- lương thời gian (thỏa thuận);
- lương cơ bản (đóng BH);
- từng khoản phụ cấp: nhà ở, xăng xe, điện thoại, trách nhiệm, chuyên cần…;
- số người phụ thuộc;
- mức ăn ca riêng (đ/ngày; để trống thì theo nơi làm việc VP/DA).

**Cách khai**
- Hồ sơ CBCNV → tab **Lương & mức hưởng** → **+ Khai báo thay đổi mức hưởng**, hoặc menu **Lương & nâng bậc** → tab **Khai báo mức hưởng**.
- Chỉ nhập khoản thay đổi; ô trống giữ nguyên mức đang hưởng; nhập 0 để thôi hưởng khoản đó. Ô nhập hiện sẵn "đang hưởng: …" theo ngày hiệu lực đã chọn.
- **Nhập hàng loạt bằng Excel**: Tải mẫu Excel (sheet KhaiBao có sẵn mã NV, sheet MucDangHuong để đối chiếu) → điền ngày hiệu lực và các mức thay đổi, mỗi lần thay đổi một dòng → Nhập Excel. Dùng cách này để đưa lịch sử lương 01–12/2026 từ file Excel cũ vào app.

**Phê duyệt**
- Thay đổi lương thời gian, lương cơ bản do TCHC khai → **chờ Giám đốc duyệt**. Giám đốc / Quản trị khai thì có hiệu lực ngay.
- Phụ cấp, số người phụ thuộc, ăn ca có hiệu lực ngay khi TCHC lưu (có ghi nhật ký người khai).
- Quyết định lên lương đã duyệt (Đề xuất lương) cũng tự đưa vào diễn biến.
- Chỉ Quản trị sửa hoặc xóa được khai báo đã có hiệu lực. Khi xóa một mốc, các mốc sau và hồ sơ tự tính lại mức cũ.

**Khi tính lương**
- Mỗi kỳ dùng mức có hiệu lực vào **ngày cuối tháng**: kỳ 04/2026 dùng mức khai từ 01/03, kỳ 06/2026 dùng thêm mức khai từ 01/05…
- Số người phụ thuộc theo kỳ quyết định giảm trừ gia cảnh của kỳ đó.
- Lưu khai báo (hoặc duyệt, xóa) thì các kỳ lương đang xử lý từ ngày hiệu lực trở đi được **tính lại tự động**. Kỳ đã phê duyệt không đổi.
- Bảng **Diễn biến mức hưởng** trong hồ sơ cho thấy mức từng giai đoạn; ô tô màu là khoản thay đổi tại mốc đó. CBCNV xem được diễn biến của mình ở "Hồ sơ của tôi".
- Đã có diễn biến thì màn hình "Sửa hồ sơ" khóa các ô lương, phụ cấp, NPT, ăn ca để không mất lịch sử.

## 4h. Phòng ban, công trình & điều động

Menu **Quản lý → Phòng ban & công trình** (TCHC, Quản trị sửa; Giám đốc, Kế toán xem).

- Danh mục mặc định: Ban Giám đốc, Phòng Tổ chức – Hành chính, Phòng Tài chính – Kế toán, Phòng Kỹ thuật, Phòng Kinh tế – Kế hoạch. Thêm bằng **+ Phòng ban** / **+ Công trình** (tên, mã, địa điểm, ngày bắt đầu – kết thúc).
- Mỗi công trình đang thi công là một **Ban chỉ huy công trình** (vd "BCH công trình Cầu Hòa Lạc"). Mỗi thẻ hiện số người, Trưởng phòng / CHT và danh sách CBCNV; bấm tên để mở hồ sơ.
- **Điều động CBCNV**: chọn bộ phận đến, từ ngày, số quyết định, chức vụ mới (nếu đổi), rồi tích chọn nhiều người. Mỗi người được ghi một dòng "Điều động" vào Quá trình công tác, dòng trước được khép lại. Ngày bắt đầu trong tương lai thì hồ sơ tự chuyển bộ phận khi đến ngày. **Lịch sử điều động** của từng bộ phận xem trên thẻ.
- Trưởng phòng / CHT được điều động thì quyền duyệt đi theo họ sang bộ phận mới. Cần giữ quyền ở bộ phận cũ thì sửa lại tại Tài khoản.
- Đổi tên bộ phận thì hồ sơ, quá trình, tài khoản và đối tượng bài viết Bảng tin được cập nhật theo. Công trình xong: điều động hết người đi rồi đặt trạng thái **Đã kết thúc**. Không kết thúc hay xóa được bộ phận còn người.
- Hồ sơ CBCNV chọn bộ phận từ danh mục. Chuyển bộ phận theo giai đoạn nên dùng Điều động để giữ lịch sử.

## 4i. Tài khoản & chức danh

- Chỉ **1 tài khoản Quản trị** (`admin`), dùng để quản trị hệ thống. Mọi người khác đăng nhập bằng tài khoản cá nhân theo họ tên, gắn với hồ sơ CBCNV.
- Mỗi tài khoản có **một chức danh**: Giám đốc, Tổ chức – Hành chính, Kế toán tiền lương, Kế toán công nợ, Trưởng phòng, Chỉ huy trưởng công trình, Nhân viên. Trưởng phòng / CHT chọn thêm **Bộ phận quản lý**.
- Quản trị gán mọi chức danh. TCHC chỉ tạo, sửa tài khoản chức danh Nhân viên.
- Màn hình Tài khoản hiện ai đang giữ từng chức danh, quyền của từng chức danh và bảng phân quyền.

| Chức danh | Việc trong quy trình |
|---|---|
| Kế toán công nợ | Chấm công; kiểm tra công nợ, tạm ứng |
| Kế toán tiền lương | Lập bảng lương |
| Tổ chức – Hành chính | Duyệt chấm công, nghỉ, nâng lương (lập/xác nhận đề xuất), bảng lương; bổ sung khấu trừ theo quy chế (mất mát, bồi thường…); xử lý đề nghị của CBCNV |
| Trưởng phòng / Chỉ huy trưởng | Duyệt đơn nghỉ của CBCNV thuộc bộ phận mình; xem nhân sự bộ phận (menu **Bộ phận của tôi**). Từ bản 3.7.1 **không** vào Chấm công GPS (quản lý) và Phân tích nhóm (AI) |
| Giám đốc | Toàn quyền phê duyệt. Giám đốc duyệt thì bỏ qua các bước còn lại |

## 4j. Lịch nghỉ lễ, Tết

Menu **Công ty → Lịch nghỉ lễ**: mọi người xem; TCHC và Quản trị sửa.

- **Nạp lịch theo luật năm …**: app tự tính các ngày nghỉ hưởng nguyên lương:
  - Tết Dương lịch 1/1.
  - Tết Âm lịch 5 ngày, gồm ngày cuối năm Âm lịch và mùng 1 đến mùng 4, tính theo âm lịch.
  - Giỗ Tổ 10/3 Âm lịch.
  - 30/4 và 1/5.
  - Quốc khánh 2/9 và 1 ngày liền kề.
  - Ngày Văn hóa Việt Nam 24/11, áp dụng từ năm 2026.
  - Ngày lễ trùng Chủ nhật thì app tự thêm ngày **nghỉ bù** vào ngày làm việc kế tiếp.
- Hộp chọn hiện danh sách để bỏ chọn trước khi thêm. Ngày cụ thể của Tết và ngày liền kề Quốc khánh theo thông báo hằng năm của Chính phủ; đối chiếu và sửa lại nếu khác.
- **+ Thêm ngày nghỉ**: chọn từ ngày – đến ngày, tên, loại. Loại gồm *Lễ, Tết*, *Nghỉ bù* hoặc *Công ty cho nghỉ* (vd nghỉ mát). Bấm một ngày trên lịch để sửa hoặc xóa.
- Các ngày này không bị trừ vào đơn nghỉ phép và được điền sẵn ký hiệu **L** trong bảng chấm công theo ngày.

## 4k. Chấm công theo ngày

Menu **Chấm công & Tính lương** → tab **Chấm công theo ngày**. Kế toán công nợ nhập ở bước Chấm công; người khác chỉ xem.

| Ký hiệu | Nghĩa | Quy đổi |
|---|---|---|
| X | Làm việc cả ngày | 1 công |
| ½ | Làm nửa ngày | 0,5 công |
| CT | Đi công tác | 1 công |
| KP | Nghỉ không phép | 1 ngày nghỉ không lương, gợi ý ghi vi phạm |
| L | Nghỉ lễ, Tết | 1 ngày lễ hưởng lương |
| P / R / Ô / KL | Nghỉ phép năm / việc riêng có lương / ốm / không lương | 1 ngày nghỉ tương ứng |
| P½, R½, Ô½, KL½ | Nghỉ nửa ngày theo đơn | 0,5 nghỉ + 0,5 công |

**Cách nhập**
1. Bấm **Điền mặc định**: ngày thường điền X, ngày lễ điền L, Chủ nhật để trống.
2. Chọn ký hiệu trên thanh công cụ rồi bấm vào ô, hoặc kéo chuột qua nhiều ô. Bấm số ngày ở đầu cột để tô cả cột. Bấm tên một người để tô các ô trống của người đó.
3. Bấm **Lưu bảng chấm công ngày**.

**Ngày nghỉ theo đơn**
- Ngày có **đơn nghỉ đã duyệt** tự hiện P / R / Ô / KL. Các ô này bị **khóa, không chấm công**. Muốn đổi thì sửa hoặc hủy đơn.
- Đơn được duyệt hoặc hủy sau khi đã lưu bảng chấm công: bảng tự cập nhật. Ngày của đơn bị hủy chuyển thành X.
- Ngày có đơn đang chờ duyệt có chấm vàng ở góc ô.

**Tổng hợp sang tab "Tổng hợp công & công nợ"**
- Khi lưu, các cột Ngày công, Nghỉ HL, Nghỉ ốm, Nghỉ KL, Ngày lễ, Giờ làm lễ được tính lại. Các cột này bị khóa ở tab Tổng hợp. Giờ tăng ca ngày thường vẫn nhập ở tab Tổng hợp.
- Làm vào ngày lễ: vẫn hưởng lương ngày lễ, cộng thêm số giờ làm lễ (giờ làm/ngày trong Cài đặt).
- Làm Chủ nhật: mặc định tính công thường. Muốn tính giờ làm ngày nghỉ tuần (200%) thì đổi tại Cài đặt → "Làm Chủ nhật".

**Xuất Excel** ra bảng chấm công tháng để in và ký.

### Nhóm theo bộ phận, tự lưu chạy nền (v3.22)

- Bảng chấm công ngày **nhóm theo bộ phận công tác**: dòng tiêu đề từng phòng ban (trước) và công trường (sau, theo thứ tự danh mục), ghi số người. Bấm tên bộ phận để tô các ô trống của cả bộ phận bằng ký hiệu đang chọn. Dưới tên người hiện chức vụ và công ty.
- Ô lọc **bộ phận** và **công ty** phía trên bảng. "Điền mặc định" chỉ điền cho những người đang hiện khi có lọc.
- **Tự lưu**: ngừng tô khoảng 1,5 giây là app tự gửi lên máy chủ, chạy nền – cứ chấm tiếp, không phải chờ, bảng không bị vẽ lại (giữ vị trí cuộn). Dòng trạng thái: "Đang lưu…" / "x người chưa lưu" / "✓ Đã lưu lên máy chủ lúc hh:mm:ss". Nút **Lưu ngay** lưu tức thì.
- Ô tô trong lúc đang gửi được lưu ở lượt kế tiếp, không mất. Mất mạng: thay đổi giữ trên máy, báo "Chưa lưu được", tự thử lại mỗi 15 giây; đóng trang khi còn thay đổi chưa lưu thì trình duyệt hỏi lại.
- Máy chủ: lưu chấm công ở bước Chấm công không còn đọc sheet BangLuong (giảm ~60% dữ liệu đọc mỗi lần lưu, phản hồi nhỏ hơn).

## 4l. Tạm ứng và vi phạm – khấu trừ

**Tab Tạm ứng** (Kế toán công nợ, ở bước Công nợ)
- **+ Thêm tạm ứng**: chọn người, loại, ngày, số tiền, ghi chú hoặc số phiếu chi.
- Loại mặc định: *Ứng lương*, *Ứng tiền ăn*, *Ứng xăng xe – đi lại*, *Ứng khác*. Sửa danh sách loại trong Cài đặt.
- Tổng các khoản của một người là số tạm ứng trừ vào lương kỳ đó. Ô "Tạm ứng" ở tab Tổng hợp bị khóa theo bảng kê.
- Người đã nhập số tổng từ trước: khi thêm khoản đầu tiên, số cũ được giữ thành khoản "Ứng khác".
- Phiếu lương in từng khoản tạm ứng.
- **Nhập tạm ứng bằng Excel**
  1. Bấm **Tải file mẫu Excel**. File gồm:
     - Sheet *TamUng*: danh sách CBCNV của kỳ, mỗi loại tạm ứng một cột, cột Tổng là công thức.
     - Sheet *ChiTiet*: kê từng lần ứng (Mã NV, loại, ngày, số tiền).
     - Sheet *HuongDan*: cách điền.
  2. Kế toán điền rồi bấm **Nhập từ Excel**. App tự nhận diện:
     - Người: theo Mã NV; thiếu mã thì theo họ tên, nếu trùng duy nhất.
     - Loại tạm ứng: theo tên cột hoặc ô Loại, kể cả viết tắt như "xăng xe", "tiền cơm".
     - Số tiền: nhận 1500000, 1.500.000, 1,5tr, 500k.
     - Ngày: dd/mm/yyyy.
  3. Màn hình xem trước hiện tổng theo loại, các dòng không nhận được (sai mã, số tiền không đọc được) và cảnh báo. Chọn **Cộng thêm vào bảng kê hiện có** hoặc **Thay bảng kê của những người có trong file**, rồi bấm Lưu.
  4. Cũng nhận file tự lập, miễn có dòng tiêu đề chứa "Mã NV" hoặc "Họ và tên", kèm các cột loại tạm ứng hoặc cột "Loại" + "Số tiền".

**Tab Vi phạm – khấu trừ** (TCHC, ở bước TCHC duyệt công)
- Mỗi vi phạm thuộc một nhóm: **Nghỉ không phép**, **Vi phạm hợp đồng**, **Mất mát, hư hỏng tài sản**, **Vi phạm nội quy – an toàn**, **Vi phạm khác**. Đầu tab có thẻ tổng tiền theo từng nhóm.
- Ngày đánh KP trên bảng chấm công ngày hiện thành gợi ý "Ghi vi phạm", điền sẵn số ngày.
- Danh mục mức khấu trừ (Cài đặt → Danh mục vi phạm) có thêm cột Nhóm.
- **Lưu ý pháp lý:** Bộ luật Lao động không cho phạt tiền thay kỷ luật (Điều 127). App chỉ khấu trừ theo 2 hình thức hợp lệ: *giảm thưởng / phụ cấp theo quy chế*, hoặc *bồi thường thiệt hại* (mất mát, hư hỏng tài sản), tối đa 30% lương thực trả (Điều 102). Mức cụ thể phải có trong Nội quy / Quy chế đã ban hành.

## 4m. Phiếu lương chi tiết

CBCNV xem phiếu lương ở **Phiếu lương của tôi** (sau khi kỳ được phê duyệt). Phiếu gồm 5 phần:

| Phần | Nội dung |
|---|---|
| Thông tin | Lương chính, lương đóng BH, công hưởng lương / công chuẩn, số người phụ thuộc; ngày công, nghỉ, giờ tăng ca |
| I. Thu nhập | Từng khoản theo công thức lương, nhóm Lương / Phụ cấp, tiền ăn / Làm thêm / Thưởng; ghi rõ phần miễn thuế; giảm trừ theo quy chế (kèm vi phạm) |
| II. Trích theo lương | BHXH, BHYT, BHTN, đoàn phí |
| III. Thuế TNCN | Thu nhập không chịu thuế, chịu thuế, giảm trừ bản thân, người phụ thuộc, BH được trừ, thu nhập tính thuế, thuế |
| IV. Phải thu, khấu trừ | Từng khoản tạm ứng (loại, ngày), khấu trừ công nợ, bồi thường (kèm vi phạm) |
| V. Còn nhận | = I − II − III − IV, kèm bằng chữ; dòng ghi phần công ty đóng thêm (BH, KPCĐ) |

## 4n. Báo cáo lương và xuất Excel có công thức

Menu **Quản lý → Báo cáo lương** (Giám đốc, TCHC, Kế toán, Quản trị). Mỗi tab có nút **Xuất Excel**. Kỳ chưa phê duyệt ghi "BẢN NHÁP". Báo cáo năm mặc định chỉ lấy kỳ đã duyệt (bỏ dấu "Chỉ kỳ đã duyệt" để lấy cả nháp).

| Báo cáo | Trên app | File Excel |
|---|---|---|
| Bảng lương tháng | Theo bộ phận, cộng nhóm, bấm dòng xem phiếu | Sheet bảng thanh toán lương: từng khoản thu nhập, khấu trừ; Tổng thu nhập, Tổng khấu trừ, Còn nhận là công thức; cộng theo bộ phận bằng SUBTOTAL; chỗ ký nhận và ký duyệt. Sheet ChuyenKhoan (ngân hàng, số TK, số tiền) |
| Bảng lương năm | 12 tháng, biểu đồ quỹ lương + BH DN | Sheet tổng hợp 12 tháng (Còn nhận, Tổng chi phí là công thức); sheet thu nhập, thu nhập chịu thuế, thuế TNCN, còn nhận của từng người theo tháng (SUM hàng, cột) |
| Lương cá nhân | Chọn người + năm | Bảng 12 tháng của 1 người: từng khoản, BH, TN chịu thuế, giảm trừ, thuế, khấu trừ, còn nhận; dòng cộng năm (dùng đối chiếu quyết toán thuế) |
| Bảng tính lương chi tiết | Bấm một người để xem từng bước tính | **Toàn bộ phép tính là công thức Excel**: dữ liệu đầu vào → giá trị trung gian → từng khoản thu nhập (dịch từ màn hình Công thức lương) → phần miễn thuế → giảm trừ quy chế → lương đóng BH → BHXH/BHYT/BHTN (có trần) → đoàn phí → giảm trừ gia cảnh → thuế lũy tiến → bồi thường (tối đa 30%) → còn nhận → BH DN, KPCĐ, tổng chi phí. Tham số (tỷ lệ BH, trần, giảm trừ, biểu thuế) ở sheet ThamSo: sửa tham số là cả bảng tự tính lại. Cột **Chênh lệch** so với số trên app phải bằng 0 |
| Chi phí lương & bảo hiểm | Theo năm hoặc tháng, theo bộ phận, tỷ trọng | Sheet DuLieu (từng người – từng tháng) và 2 sheet tổng hợp theo bộ phận, theo tháng bằng SUMIFS / COUNTIFS: quỹ lương, BH NLĐ, BHXH/BHYT/BHTN DN, KPCĐ, tổng chi phí, tỷ trọng |

- File Excel mở được bằng Excel, Google Sheets, LibreOffice. Đã kiểm chứng: tính lại toàn bộ công thức bằng LibreOffice cho ra đúng số của app (hơn 600 công thức, 2 mẫu công thức lương).
- Bảng tính chi tiết dùng công thức lương đã lưu theo kỳ, còn tham số BH, thuế lấy theo Cài đặt hiện tại. Nếu sau khi duyệt kỳ có đổi tham số, cột Chênh lệch sẽ khác 0 để biết.
- Thư viện Excel tải từ mạng (cdn.jsdelivr.net). Máy không có mạng thì nút xuất báo lỗi.

## 4o. Tin nhắn kiểu Zalo: nhóm, bình chọn, nhắc hẹn, báo cáo định kỳ, thông báo

**Giao diện**
- **Danh sách trò chuyện:**
  - Thanh tìm kiếm màu xanh, nút **+** (tạo nhóm, nhắn tin mới, kho tệp, cài đặt thông báo).
  - Lọc Tất cả / Chưa đọc / Nhóm.
  - Mỗi dòng có ảnh đại diện (nhóm ghép ảnh thành viên), tin cuối, thời gian ("3 giờ"), số tin chưa đọc màu đỏ.
- **Khung chat:**
  - Tin của mình màu xanh nhạt, của người khác màu trắng, có tên người gửi và giờ trong bong bóng.
  - Mốc thời gian dạng "15:56 Hôm nay".
  - Nút ♥ thả tim; chuột phải hoặc giữ để chọn biểu tượng khác.
  - Chạm vào tên nhóm ở đầu khung để xem thông tin.
  - Thanh nhập: biểu tượng cảm xúc, ô "Tin nhắn", nút **…** (Ảnh, Tệp, Bình chọn, Nhắc hẹn, Yêu cầu báo cáo, Kho tệp), nút ảnh. Gõ chữ thì hiện nút gửi.
  - Dải trên cùng của nhóm nhắc báo cáo hôm nay mình chưa nộp, hoặc nhắc hẹn sắp tới.

**Trưởng nhóm / phó nhóm**

| Việc | Trưởng nhóm | Phó nhóm | Thành viên |
|---|:-:|:-:|:-:|
| Đổi tên, mô tả nhóm | ✓ | ✓ | |
| Phân quyền trưởng / phó nhóm | ✓ | | |
| Mời thành viên ra (phó không mời được trưởng, phó khác) | ✓ | ✓ | |
| Tạo yêu cầu báo cáo định kỳ | ✓ | ✓ | |
| Thêm người, tạo bình chọn, nhắc hẹn | ✓ | ✓ | ✓ |
| Giải tán nhóm | người tạo nhóm | | |

**Bình chọn:** câu hỏi, các phương án, cho chọn nhiều hay không, cho thành viên thêm phương án, hạn bình chọn. Bấm vào phương án để chọn hoặc đổi lựa chọn. Thẻ hiện số phiếu và người đã chọn. Người tạo, trưởng hoặc phó nhóm khóa được bình chọn.

**Nhắc hẹn:** nội dung, ngày giờ, lặp lại (không lặp, hằng ngày, hằng tuần, hằng tháng). Đến giờ, app gửi thẻ "Đến giờ nhắc hẹn" vào nhóm và báo trên điện thoại các thành viên.

**Yêu cầu báo cáo định kỳ** (trưởng hoặc phó nhóm tạo):
1. Chọn mẫu có sẵn rồi sửa, thêm, bớt mục. Mẫu có sẵn:
   - **Báo cáo công trường (nhân lực, máy, công việc)** — mẫu mặc định từ bản 3.9, dùng các mục *Bảng số lượng*
   - Báo cáo công trường hằng ngày
   - Báo cáo công việc trong ngày
   - Báo cáo vật tư công trường
   - Báo cáo an toàn – vệ sinh lao động
   - Mẫu trống
   Mỗi mục có kiểu nhập (**bảng số lượng**, chữ ngắn, đoạn văn, số, chọn trong danh sách, Có/Không, ngày) và đánh dấu bắt buộc hay không.

   **Bảng số lượng (chọn mục gợi ý)** — từ bản 3.9. Người lập yêu cầu chọn sẵn các dòng; người báo cáo chỉ gõ số vào từng dòng:
   - Bộ gợi ý có 3 nhóm. **Nhân lực** (người): NC ván khuôn, NC thép, NC bê tông, NC xây, NC trát – ốp lát, NC lắp cống, NC lắp dựng kết cấu, Thợ điện – nước, Thợ hàn, Thợ vận hành máy, Lái xe, Công nhật, Lao động phổ thông, Kỹ thuật – giám sát, Thợ khác. **Máy thi công** (cái): Máy đào, ủi, lu, san, Ô tô tự đổ, Cần cẩu, Máy trộn bê tông, Xe bơm bê tông, Máy đầm, Máy cắt uốn thép, Máy phát điện, Máy hàn, Máy bơm nước, Máy nén khí, Máy khoan, Máy toàn đạc, Máy thủy bình. **Công việc thực hiện** (đơn vị riêng từng việc): Đào đất, Đắp đất, Đắp cát (m³), Lu lèn nền (m²), Cấp phối đá dăm (m³), Cốt thép (tấn), Ván khuôn (m²), Bê tông (m³), Xây (m³), Trát, Ốp lát, Sơn bả (m²), Lắp đặt cống, ống cấp thoát nước, kéo dây điện (m), Thảm bê tông nhựa (m²), Vận chuyển vật liệu (chuyến).
   - Trong mẫu mặc định, mỗi mục đã điền sẵn cả nhóm gợi ý tương ứng. Bấm × để bỏ dòng không dùng, **Bỏ hết** để chọn lại từ đầu; ô *+ Chọn 1 mục gợi ý* thêm từng dòng, **+ Cả nhóm** thêm cả nhóm; ô *Dòng khác* + *Đơn vị* → **Thêm** để thêm dòng riêng của công trường (vd "NC sơn kẻ đường – người").
   - Tích **Cho người báo cáo tự thêm dòng khác** nếu muốn người điền tự thêm dòng chưa có trong danh sách.
   - Khi điền: mỗi dòng một ô số (gõ được kiểu Việt Nam `1.250`, `42,5`), dòng không có thì để trống. Với nhân lực và máy, app **tự cộng Tổng** ngay khi gõ; khối lượng các công việc khác nhau thì không cộng gộp.
   - Kết quả lưu thành từng dòng "NC thép: 4 người … Tổng: 13 người", hiện đúng như vậy khi xem, trong bảng tổng hợp, Excel và tệp lưu trữ. **Phân tích nhóm (AI)** đọc thẳng số liệu này: tổng nhân lực kèm chi tiết từng loại thợ, danh sách máy, khối lượng từng công việc, và cộng dồn từng dòng theo cả kỳ.
2. Đặt **giờ phát mẫu** lên nhóm (vd 16:00), **hạn nộp** (vd 17:30), các ngày trong tuần, và **người phải báo cáo** (để trống = cả nhóm).
3. Hằng ngày đến giờ, thẻ "Báo cáo ngày …" hiện trên nhóm, có thanh tiến độ "2/5 người đã hoàn thành" và ảnh đại diện từng người.
4. Người được giao bấm **Điền báo cáo**, điền mẫu rồi bấm **Hoàn thành báo cáo**. Nhóm thấy tin "✅ Đã hoàn thành …". Sửa lại được trong ngày.
5. Đến hạn nộp, app nhắc tên những người chưa nộp (1 lần/ngày).
6. **Xem tổng hợp**: bảng từng người × từng mục theo ngày, xuất được Excel.
7. Trong Thông tin nhóm → tab Báo cáo, tạm dừng, mở lại hoặc xóa được yêu cầu báo cáo.

> Trên máy chủ thật, nhắc hẹn và báo cáo được phát bởi trình kích hoạt của Apps Script. Sau khi nâng cấp, mở Apps Script và **chạy 1 lần hàm `caiDatNhacTuDong`** (chấp nhận quyền). Trình kích hoạt chạy 5 phút/lần, nên giờ nhắc có thể lệch tối đa 5 phút.

**Thông báo trên điện thoại**
- Nút **+** → **Cài đặt thông báo** → **Cho phép thông báo**. Bật/tắt từng mục: thông báo tin mới, hiện nội dung, âm báo, rung. Có nút gửi thông báo thử.
- **Khi đang mở app:** tin mới hiện banner nổi phía trên màn hình (vuốt lên để ẩn, bấm để mở) kèm tiếng "ting".
- **Khi app chạy nền:** thông báo hiện trên thanh trạng thái và màn hình khóa. Bấm vào là mở đúng cuộc trò chuyện. Số tin chưa đọc hiện trên biểu tượng app nếu máy hỗ trợ.
- Tắt thông báo riêng từng nhóm trong Thông tin nhóm.
- **Giới hạn:** app phải đang mở hoặc chạy nền (chưa vuốt tắt hẳn) mới báo được. Muốn nhận thông báo cả khi đã tắt app như Zalo thì cần dịch vụ đẩy thông báo (Firebase Cloud Messaging, miễn phí), sẽ làm ở bước sau.
- Nên **cài app ra màn hình chính**:
  - Android: Chrome → ⋮ → Thêm vào màn hình chính.
  - iPhone (iOS 16.4 trở lên): Safari → Chia sẻ → Thêm vào MH chính, mở app từ biểu tượng rồi cho phép thông báo.

## 4p. Lưu trữ nhóm chat và phân tích bằng AI cho Ban Giám đốc

**Tự lưu file kết quả sau khi phân tích (3.5.3)**: chạy xong lượt phân tích, hệ thống tự dựng **Excel kết quả** (các sheet: nhật ký thi công, khối lượng, vật tư – thiết bị, chứng từ, sự kiện, thầu phụ – NCC, cá nhân) và lưu vào đúng thư mục **“4. Phân tích AI”** của nhóm, cạnh Google Docs, PDF và tệp JSON đã có. Tên tệp trùng với tệp JSON của đợt (`PhanTich_<từ>_<đến>_<lần lập>.xlsx`); chạy lại phân tích cùng kỳ thì **ghi đè**, không sinh tệp rác. Mã tệp được ghi vào JSON của đợt nên mở lại báo cáo cũ vẫn thấy. Dòng “Đã lưu trên Drive: …” ngay đầu báo cáo liệt kê những gì đã lưu. Excel do trình duyệt dựng (thư viện chạy ở trình duyệt) rồi gửi lên máy chủ; mất mạng hoặc thiếu thư viện thì chỉ báo ở dòng tiến trình, **không làm hỏng lượt phân tích** và vẫn tải được bằng nút Excel.


**Số liệu máy bóc sẵn từ báo cáo và lưới an toàn (3.5.2)**: trong nhật ký gửi AI, ngay dưới mỗi lần nộp báo cáo có thêm dòng `[số liệu máy đọc được, dùng làm chuẩn]` — máy tự tách ô văn bản thành hạng mục + khối lượng + đơn vị (ví dụ “Đổ bê tông trụ T3: 42 m³; lắp cốt thép dầm D3 60%” → 2 hạng mục). Máy **chỉ nhận một con số khi đi kèm đơn vị hợp lệ**, nên ký hiệu cấu kiện như “đài móng T4”, “dầm D3” không bị hiểu nhầm thành khối lượng. Sau khi AI trả kết quả, hàm `aiBuBC` đối chiếu lại: ngày nào báo cáo **có** khối lượng, nhân lực, thiết bị hoặc vật tư mà AI trả về **rỗng** thì hệ thống điền lại từ số liệu máy đã bóc (ghi chú “lấy từ báo cáo định kỳ”) và quy việc về đúng tài khoản người nộp. Nhờ vậy khối lượng do CBCNV báo cáo không bị mất kể cả khi AI bỏ sót.

**Vật tư tách cột nhập – xuất (3.5.2)**: bảng tổng hợp vật tư (trên màn hình, báo cáo Word/PDF và Excel) nay mỗi vật tư **một dòng** với các cột **Nhập · Xuất · Đề nghị cấp · Nhập − Xuất**, thay cho một cột “Hoạt động” gộp chung làm cùng một vật tư bị tách thành nhiều dòng. Bảng chi tiết từng lần cũng đưa số lượng vào đúng cột. Cột “Nhập − Xuất” chỉ hiện khi biết **cả hai chiều**, và là chênh lệch của những gì được nêu trong nhóm — **không phải tồn kho**; vẫn phải đối chiếu phiếu nhập, phiếu xuất kho.


**Báo cáo ngày trong phân tích AI (3.5.1)**: nhật ký mỗi ngày gửi cho AI đặt khối **“## Báo cáo định kỳ (số liệu do thành viên nộp – nguồn ưu tiên)” TRƯỚC** phần “## Diễn biến trong ngày”, và mỗi dòng người nộp ghi kèm tài khoản: `**Họ tên** (tài_khoản, nộp lúc hh:mm)`. Khi nhật ký một ngày vượt 45.000 ký tự, hàm `ltRutGon` chỉ lược bớt **phần trò chuyện** và giữ nguyên khối báo cáo — trước 3.5.1 nhật ký bị cắt thẳng từ cuối nên khối báo cáo (nằm cuối) **mất toàn bộ**, kéo theo mất số nhân lực, khối lượng và công việc do thành viên báo cáo.

Prompt trích xuất (`AI_P_DOT`) nêu rõ phải đọc hết khối báo cáo và ánh xạ tên trường của mẫu báo cáo sang các mục nhật ký thi công: *Nhân lực có mặt / Số người trên công trường / Ban chỉ huy có mặt* → `nhanLuc`; *Khối lượng thực hiện trong ngày / Việc đã hoàn thành / Việc đang thực hiện / Công việc chính* và mọi khối lượng có đơn vị (m³, m², md, tấn) → `thiCong`; *Máy, thiết bị hoạt động* → `thietBi`; *Vật tư nhập về / sử dụng / đề nghị cấp* → `vatTuNhap`, `vatTuXuat`, `yeuCauVatTu`; *Vướng mắc, đề xuất / Khó khăn, cần hỗ trợ* → `kienNghi` (kèm tài khoản người nộp) và `tonTai`; *Kế hoạch ngày mai* → `keHoachMai`; *Thời tiết* → `thoiTiet`; *Sự cố an toàn / Kiểm tra bảo hộ* → `anToan`, `suCo`. Số liệu báo cáo được ưu tiên hơn tin nhắn; lệch nhau thì lấy theo báo cáo và ghi chênh lệch vào `ghiChu`. Việc do ai báo cáo thì quy về đúng tài khoản người đó trong `caNhan.dongGop`.


**Lưu trữ thành tệp trên Google Drive**

Toàn bộ dữ liệu từng **nhóm chat** được lưu thành tệp, trong thư mục **“TAPACO – Lưu trữ nhóm chat (KHÔNG CHIA SẺ)”** đặt **cạnh file Google Sheet** dữ liệu. Mỗi nhóm có một thư mục riêng:

| Thư mục / tệp | Nội dung |
|---|---|
| `0. Thông tin nhóm.md` | Tên, mô tả, thành viên và vai trò, các bình chọn, nhắc hẹn, yêu cầu báo cáo |
| `1. Tệp & ảnh` | Ảnh, tệp gốc gửi trong nhóm (tệp mới tự vào đây; thư mục tệp cũ của nhóm được chuyển vào khi lưu trữ lần đầu) |
| `2. Nhật ký từng ngày` | Mỗi ngày một tệp `YYYY-MM-DD.md`: toàn văn tin nhắn theo giờ, ảnh, tệp kèm mã tệp, bình chọn, nhắc hẹn, báo cáo định kỳ đã nộp (đúng hạn / trễ hạn, ai chưa nộp) |
| `3. Báo cáo định kỳ` | Mỗi yêu cầu báo cáo một tệp `.csv`, mở bằng Excel: mỗi dòng một lần nộp |
| `4. Phân tích AI` | Báo cáo phân tích dạng **Google Docs** và tệp dữ liệu `.json` |

- **Tự động mỗi đêm:** hệ thống lưu các nhóm có tin mới. Giám đốc, TCHC, Kế toán, Quản trị bấm **Lưu trữ ngay** được bất cứ lúc nào (trưởng, phó nhóm không còn quyền này từ bản 3.7.1).
- **Ngày cũ có thay đổi** (nộp bù báo cáo, thu hồi tin) được ghi lại cho đúng.
- **Tin đã thu hồi** không lưu nội dung.
- **Tin nhắn riêng 1–1 không đưa vào lưu trữ.**
- Thư mục là **riêng tư**, chỉ chủ Google Drive (tài khoản chạy Apps Script) mở được. Nếu thư mục chứa file Google Sheet đang được chia sẻ, trang Phân tích hiện cảnh báo cho Quản trị để tắt chia sẻ.

**Phân tích bằng AI** (menu **Công ty → Phân tích nhóm (AI)**; Giám đốc, TCHC, Kế toán còn vào được từ Thông tin nhóm → **Báo cáo AI**)

1. Chọn **nhóm** hoặc **trò chuyện riêng** (bộ lọc ở đầu danh sách, chỉ Giám đốc và Quản trị thấy trò chuyện riêng), chọn khoảng ngày (tối đa 12 tháng), có thể cho AI xem ảnh chứng từ.
2. Bấm **Chạy phân tích AI**. App chạy lần lượt:
   - lưu trữ;
   - AI lập **nhật ký thi công** từng đợt (Gemini 3 ngày/đợt, Claude 2 ngày/đợt; mỗi lần gọi AI dưới 60 giây);
   - AI lập **bức tranh tổng thể** và đánh giá **18 khía cạnh**;
   - AI đánh giá **nhân sự, nhà thầu phụ, nhà cung cấp** và đề xuất **giải pháp**;
   - lập **Báo cáo** và **Nhật ký thi công** thành Google Docs kèm PDF trong thư mục “4. Phân tích AI”.

   Phân tích 1 tháng mất khoảng 5–10 phút. Bước nào lỗi thì bấm **Chạy lại từ đợt lỗi**.
3. Các tab của báo cáo:
   - **Tổng thể:** mức rủi ro cho công ty (Thấp / Trung bình / Cao); bức tranh tổng thể; ma trận 18 khía cạnh tô màu theo mức; biểu đồ hoạt động; điểm mạnh, tồn tại, kiến nghị Giám đốc chỉ đạo; sự cố, sự kiện.
   - **Đánh giá 18 khía cạnh:**
     - nhân sự & tinh thần làm việc; thời tiết & ảnh hưởng thi công; kiểm soát chất lượng; khối lượng & sản lượng; tiến độ; nhà thầu phụ;
     - nhà cung cấp; an toàn lao động; an ninh trật tự; sự cố; năng lực kiểm soát của Ban chỉ huy công trường; phối hợp phòng ban – công trường;
     - quản lý tài chính (tạm ứng, thanh toán); quản lý kho vật tư; giao nhận vật tư; quản lý thiết bị; quản lý định mức thi công; rủi ro cho công ty.

     Mỗi khía cạnh có mức (Tốt / Đạt / Cần chú ý / Nghiêm trọng / Chưa đủ dữ liệu), nhận xét, dẫn chứng theo ngày và giải pháp. Khía cạnh nào nhóm không trao đổi thì ghi "Chưa đủ dữ liệu" và gợi ý bổ sung mục vào mẫu báo cáo, không đoán.
   - **Giải pháp:** bảng vấn đề – giải pháp – bộ phận phụ trách – thời hạn – ưu tiên; kiến nghị Giám đốc chỉ đạo.
   - **Nhân sự:** mỗi người có:
     - **ngày làm việc** (kể từ khi vào nhóm, trừ Chủ nhật, lễ);
     - **ngày có tương tác**;
     - **ngày không tương tác** (không tin nhắn, tệp, báo cáo, bình chọn và không nghỉ có đơn; rê chuột để xem các ngày);
     - **ngày nghỉ có đơn** (lấy từ đơn nghỉ đã duyệt trong app);
     - báo cáo nộp / phải nộp / đúng hạn;
     - **điểm tham gia** 0–100: 50% tỷ lệ ngày có tương tác, 30% tỷ lệ báo cáo đúng hạn, 20% mức đóng góp.

     Các số trên do máy tính; xếp loại, nhận xét, đề xuất (khen, nhắc nhở…), đóng góp và **kiến nghị của từng cá nhân** do AI tổng hợp.
   - **Nhật ký thi công:** mỗi ngày gồm:
     - thời tiết (sáng/chiều, ảnh hưởng); nhân lực (tổng, **số người Ban chỉ huy**, công nhân, thầu phụ, người nghỉ có đơn); thiết bị và tình trạng;
     - công tác thi công: khối lượng thực hiện **so với kế hoạch**; vật tư nhập / xuất; yêu cầu vật tư; nghiệm thu; kiến nghị, đề xuất;
     - an toàn, bảo hộ; an ninh; sự cố; tồn tại, vướng mắc; **kế hoạch ngày mai**.
   - **Khối lượng:** số liệu báo cáo định kỳ (máy cộng), khối lượng theo hạng mục và tỷ lệ đạt so kế hoạch.
   - **Vật tư, thiết bị:** tổng hợp vật tư theo nhập / xuất / đề nghị, chi tiết có trích dẫn, thiết bị hỏng / sửa / thay thế.
   - **Thầu phụ, NCC:** từng đơn vị, đánh giá (Tốt / Đạt / Cần chú ý / Kém), các lần được nêu, đề xuất.
   - **Chứng từ:** ảnh, tệp kèm mô tả.
4. **Xuất file:**
   - **Báo cáo → Word / PDF.**
   - **Nhật ký thi công → Word / PDF:** mỗi ngày một trang, có ô ký Người lập và Chỉ huy trưởng.
   - **Excel:** 9 trang tính.
   - **Google Docs:** chỉ Quản trị thấy nút này.

   File Word tạo ngay trên máy (cần mạng để tải thư viện lần đầu). PDF lấy từ bản Google Docs đã lưu trên Drive; nếu chưa có thì app mở hộp thoại in, chọn **Lưu dưới dạng PDF**.
5. **Xuất nhật ký từng ngày (Excel):** toàn văn tin nhắn theo ngày cùng các báo cáo định kỳ. Không cần AI.

> Để nhật ký thi công đầy đủ, nên thêm vào mẫu báo cáo định kỳ các mục: Nhân lực, **Ban chỉ huy có mặt**, Thiết bị, Khối lượng thực hiện và **kế hoạch**, Vật tư nhập/xuất, Nghiệm thu, An toàn, Vướng mắc, **Kế hoạch ngày mai**. AI chỉ ghi những gì nhóm đã trao đổi.

**Quyền (từ bản 3.7.1):** chỉ **Giám đốc, TCHC, Kế toán tiền lương, Kế toán công nợ** (và Quản trị để cài đặt, hỗ trợ) vào được **Phân tích nhóm (AI)** — lưu trữ, phân tích, xem lại báo cáo của **mọi nhóm**, kể cả nhóm mình không tham gia. **Trò chuyện riêng** (thư mục “_Trò chuyện riêng”) chỉ Giám đốc và Quản trị phân tích. Trưởng, phó nhóm, Trưởng phòng, Chỉ huy trưởng không còn vào phần này (vẫn quản lý thành viên, bình chọn, yêu cầu báo cáo của nhóm như cũ). Thành viên thường không xem được đánh giá cá nhân.

**Thông báo cho CBCNV:** trong Thông tin nhóm / trò chuyện có dòng ghi rõ tin nhắn là kênh công việc, được lưu trữ và Ban Giám đốc có thể dùng AI tổng hợp, đánh giá. Nên đưa nội dung này vào quy chế nội bộ.

**Cài đặt AI** (Quản trị → Phân tích nhóm (AI) → **Cài đặt AI**)
- Chọn **Google Gemini** (khóa tại aistudio.google.com/apikey; model mặc định `gemini-flash-latest` – tên rút gọn luôn trỏ bản Flash mới nhất của Google. Từ v3.13.3, nếu Google ngừng model đang dùng (lỗi "no longer available" / 404), app tự dò danh sách model của khóa API, ưu tiên model Google gợi ý, gọi lại và lưu model mới vào Cài đặt AI) hoặc **Anthropic Claude** (khóa tại console.anthropic.com; mặc định `claude-haiku-4-5-20251001`, có thể chọn `claude-sonnet-5-5` với khoảng ngày ngắn). Tên model sửa được khi nhà cung cấp ra model mới.
- Dán khóa API → **Lưu** → **Thử kết nối**. Khóa lưu trong **Thuộc tính tập lệnh** của Apps Script, không ghi vào Google Sheet, không gửi về trình duyệt.
- Mỗi lần phân tích, nội dung nhóm trong khoảng ngày được gửi tới nhà cung cấp AI. Với Gemini, nên bật thanh toán (gói trả phí) để dữ liệu không bị dùng cải thiện sản phẩm. Kiểm tra điều khoản của nhà cung cấp trước khi dùng.

> **Nguyên tắc số liệu:** số đếm, tỷ lệ, tổng báo cáo định kỳ do hệ thống tính. AI được yêu cầu chỉ dùng dữ liệu có trong nhóm, mỗi số vật tư / khối lượng phải kèm trích dẫn; tài khoản, ngày, mã tệp AI trả về được kiểm tra lại trước khi lưu. Dù vậy, nhận xét của AI cần đối chiếu chứng từ gốc trước khi dùng để khen thưởng, kỷ luật.

## 4q. Chấm công bằng điện thoại: định vị GPS + chụp khuôn mặt (v3.0)

Cách làm tham khảo AMIS Chấm công (MISA), Tanca và Base Checkin: mỗi công trình hoặc văn phòng có **tọa độ và bán kính cho phép**. Mỗi lượt chấm công cần **vị trí đúng vùng** và **ảnh khuôn mặt**, có so khớp với mẫu đã đăng ký.

**1. Khai báo điểm chấm công** (menu Quản lý → **Chấm công GPS → Điểm chấm công**)
- Ai được khai báo: TCHC, Giám đốc và Quản trị (từ bản 3.7.1 Trưởng phòng / Chỉ huy trưởng không khai báo điểm nữa).
- Mỗi bộ phận có thể có nhiều điểm (cổng chính, bãi đúc, lán trại…). Mỗi điểm gồm:
  - tên điểm;
  - vĩ độ, kinh độ;
  - **bán kính 20–5.000 m** (công trường nên đặt 150–300 m, văn phòng 80–150 m);
  - giờ vào / giờ ra riêng (bỏ trống thì dùng giờ chung).
- Ba cách lấy tọa độ:
  - **Lấy vị trí hiện tại**: đứng tại công trường và bấm;
  - **Dán link Google Maps**;
  - gõ số trực tiếp.
- Mục **Điểm dùng chung** (VD trụ sở công ty) áp dụng cho mọi người.

**2. Đăng ký khuôn mặt** (CBCNV làm một lần)
- Vào **Chấm công** ở thanh dưới, bấm **Đăng ký khuôn mặt**.
- Nhìn thẳng vào camera, đủ sáng, và chụp 3 lần. Ứng dụng lưu mẫu đặc trưng khuôn mặt cùng 1 ảnh chân dung.
- Điện thoại dùng để đăng ký trở thành **thiết bị chính** của người đó.
- Muốn đăng ký lại (đổi máy, ảnh không đạt) thì báo TCHC vào tab **Khuôn mặt → Xóa mẫu**.

**3. Chấm công vào / ra** — ứng dụng đi qua 3 bước:
1. **Kiểm tra định vị.**
   - Nếu GPS đang tắt hoặc quyền vị trí bị chặn, ứng dụng **không cho chấm công** và hiện hướng dẫn bật định vị cho Android / iPhone. Bật xong bấm **Tôi đã bật định vị – thử lại**.
   - Nếu GPS bật, máy tự **hiệu chuẩn**: đo liên tục đến khi có ít nhất 3 lần đo ổn định đạt độ chính xác (mặc định ±50 m), tối đa 25 giây, rồi lấy trung bình.
2. **Kiểm tra vị trí.** Tính khoảng cách tới điểm gần nhất được phép.
   - Trong bán kính thì mở camera.
   - Ngoài bán kính thì báo **Sai vị trí** (cách bao nhiêu mét), **không ghi công**. Lượt thử vẫn được lưu để quản lý biết.
3. **Chụp ảnh khuôn mặt.**
   - Camera trước; nút Chụp chỉ bật khi nhận diện được khuôn mặt.
   - Ảnh có in sẵn họ tên, giờ, tọa độ.
   - Bấm **Xác nhận chấm công** để gửi. **Giờ ghi nhận là giờ máy chủ**, không lấy giờ điện thoại.

**4. Kết quả mỗi lượt**
| Trạng thái | Khi nào | Tính công |
|---|---|---|
| **Hợp lệ** | Đúng vùng + khuôn mặt khớp mẫu + đúng thiết bị đã đăng ký | Có |
| **Chờ duyệt** | Đúng vùng nhưng: khuôn mặt không khớp, máy không so khớp được, dùng thiết bị khác, hoặc định vị bất thường | Chỉ khi quản lý duyệt |
| **Sai vị trí** | Ngoài bán kính | Không (không duyệt được) |
| **Từ chối** | Quản lý từ chối (bắt buộc ghi lý do) | Không |

Lượt vào muộn quá số phút cho phép (mặc định 10 phút) được đánh dấu **trễ**. Lượt ra trước giờ được đánh dấu **về sớm**.

**5. Quản lý** (Chấm công GPS)
- Ai xem được: **chỉ Giám đốc, TCHC, Kế toán** (và Quản trị) — toàn công ty. Từ bản 3.7.1 Trưởng phòng / Chỉ huy trưởng không còn mục **Chấm công GPS** trong menu Quản lý; họ vẫn chấm công của chính mình ở mục **Chấm công**.
- **Trong ngày**: ai đã vào, chưa vào, đi trễ, chờ duyệt.
- **Nhật ký & duyệt**:
  - xem ảnh từng lượt, khoảng cách, độ chính xác, độ lệch khuôn mặt;
  - tích chọn rồi **Duyệt hợp lệ** / **Từ chối**;
  - không ai tự duyệt lượt của chính mình.
- **Tổng hợp tháng**: mỗi người được tính đủ công, nửa công, số lần trễ và phút trễ. Có nút **Xuất Excel**.
- **Cài đặt**:
  - giờ vào / ra chung;
  - phút cho phép trễ;
  - số giờ tối thiểu để tính đủ công (mặc định 6 giờ);
  - độ chính xác yêu cầu;
  - ngưỡng so khớp khuôn mặt (0,3–0,7; nhỏ hơn là chặt hơn).

**6. Đưa vào bảng chấm công lương**
- Kế toán công nợ (bước Chấm công của kỳ lương) mở **Chấm công theo ngày**, bấm **Lấy công từ chấm công GPS**, kiểm tra rồi bấm **Lưu**.
- Cách điền:
  - Ngày có đủ lượt vào + ra hợp lệ và đủ giờ được điền **X**.
  - Thiếu lượt ra hoặc chưa đủ giờ được điền **½**.
  - Ô đã có ký hiệu (nghỉ phép, ốm, công tác…) **được giữ nguyên**; chỉ ô trống được điền.

**7. Lưu trữ:** ảnh chấm công lưu trong thư mục lưu trữ cạnh Google Sheet: `_CHẤM CÔNG GPS/<năm-tháng>`. Thư mục này **không chia sẻ**.

> **Giới hạn cần biết**
> - App chạy trên trình duyệt (PWA), nên **không phát hiện được ứng dụng giả lập GPS (Fake GPS)**. AMIS cũng chỉ chặn được trên app cài đặt. Các lớp bù lại:
>   - ngưỡng độ chính xác;
>   - yêu cầu nhiều lần đo ổn định;
>   - gắn cờ tọa độ bất thường (Chờ duyệt);
>   - khóa thiết bị;
>   - so khớp khuôn mặt;
>   - giờ máy chủ;
>   - **ảnh lưu làm bằng chứng** để quản lý soát.
> - Chống giả mạo bằng ảnh chụp chỉ ở mức cơ bản (nhận diện liên tục nhiều khung hình). Quản lý nên xem ảnh các lượt Chờ duyệt.
> - Bộ nhận diện khuôn mặt (~6 MB) tải một lần khi chấm công lần đầu, sau đó dùng lại. Máy không tải được vẫn chụp được ảnh, nhưng lượt đó sẽ **Chờ duyệt**.
> - Bắt buộc mở app qua **https** (GitHub Pages, Netlify… đều đạt), nếu không trình duyệt chặn GPS và camera.
> - Trong bản dùng thử, có thể chọn mô phỏng: Định vị tắt / Đứng tại công trường / Cách 2 km.
> - Ảnh khuôn mặt là dữ liệu cá nhân. Nên thông báo mục đích sử dụng trong quy chế chấm công và lấy xác nhận của người lao động.

## 4r. Bản đồ chấm công, điểm riêng, công tác (v3.1)

**Bản đồ.** Bản đồ dùng OpenStreetMap, miễn phí, không cần khóa. Nếu Công ty có nhà cung cấp bản đồ riêng, nhập địa chỉ lớp bản đồ ở **Chấm công GPS → Cài đặt**.
- **Điểm chấm công → Sửa:** khai báo điểm ngay trên bản đồ.
  - Bấm vào ô của một điểm để chọn điểm đó.
  - Đặt tâm bằng một trong ba cách: bấm lên bản đồ, kéo ghim cam, hoặc gõ địa chỉ vào ô **Tìm**.
  - Vòng tròn hiện đúng bán kính đã khai báo.
  - Vẫn dùng được nút **Lấy vị trí hiện tại** và **Dán link Google Maps**.
- **Tab Bản đồ** (quản lý): xem mọi người đã chấm công ở đâu trong ngày.
  - Mỗi lượt là một chấm màu: xanh là hợp lệ, cam là chờ duyệt, đỏ là sai vị trí.
  - Bản đồ cũng hiện vòng các điểm chấm công và điểm công tác.
  - Bấm tên một người để bản đồ nhảy tới vị trí của người đó.
- **Màn hình chấm công:** bản đồ hiện vị trí của bạn (kèm sai số) so với các điểm được phép. Nếu điện thoại không tải được bản đồ, app hiện sơ đồ vòng tròn như bản 3.0.

**Điểm riêng của cá nhân.** Dùng cho người làm việc cố định ở nơi khác bộ phận (kho, mỏ, trạm trộn, văn phòng đại diện).
- Khai báo ở **Điểm chấm công → + Điểm riêng cho nhân viên**.
- Người đó chấm công được ở điểm riêng và các điểm của bộ phận; công vẫn tính về bộ phận đang công tác.
- Ai khai báo thay được: TCHC, Kế toán, Giám đốc (Quản trị).

**Công tác** (đổi chỗ làm việc vẫn chấm công được)
- **CBCNV đăng ký** tại **Chấm công → Công tác → + Đăng ký công tác**:
  - nhập từ ngày – đến ngày, nơi công tác, mục đích, phương tiện;
  - chọn một trong hai chế độ:
    - **Theo điểm cố định:** chọn nơi làm việc trên bản đồ (bán kính mặc định 300 m). Đi nhiều nơi thì thêm nhiều điểm, mỗi điểm cho mọi ngày hoặc một ngày cụ thể.
    - **Linh hoạt:** dành cho người đổi chỗ hằng ngày (khảo sát, đi nhiều mỏ, nhiều công trường). Trong các ngày công tác, chấm công tại nơi đang đứng; mỗi lượt vẫn có định vị và ảnh khuôn mặt, quản lý xem lại trên bản đồ.
- **Khai báo tại chỗ:** khi chấm công báo **Sai vị trí**, có nút **Đi công tác? Khai báo tại vị trí này**. Đăng ký được tạo với vị trí đang đứng rồi chấm công tiếp; lượt chấm công **chờ duyệt** cùng đăng ký.
- **Duyệt** (tab **Công tác**): **TCHC và Kế toán** nhận thông báo và duyệt (Giám đốc duyệt được); không qua trưởng bộ phận / CHT.
  - **Duyệt:** các lượt chấm công đang chờ vì công tác tự thành **Hợp lệ**.
  - **Từ chối** (bắt buộc ghi lý do): các lượt đó bị từ chối.
  - Không tự duyệt công tác của mình.
- **Quản lý cử đi:** bấm **+ Khai báo công tác cho nhân viên**. Đăng ký được duyệt ngay và báo cho người được cử.
- Ngày công tác **không tính đi trễ / về sớm**.
- Khi bấm **Lấy công từ chấm công GPS**:
  - ngày công tác có chấm công hợp lệ được điền ký hiệu **CT**;
  - ngày công tác không chấm công được liệt kê để Kế toán xem xét.
- Hủy:
  - CBCNV tự hủy được đăng ký chưa diễn ra;
  - chuyến đã bắt đầu thì quản lý hủy.

## 4s. Thông báo đẩy trên điện thoại – giống Zalo (v3.1)

Điện thoại **đổ chuông, rung, hiện thông báo trên thanh trạng thái và màn hình khóa**, kể cả khi đã tắt app. Bấm vào thông báo thì mở đúng màn hình liên quan.

| Loại | Khi nào | Ai nhận |
|---|---|---|
| Tin nhắn | Tin mới trong nhóm / trò chuyện riêng, nhắc hẹn, yêu cầu báo cáo | Thành viên (trừ người gửi và người đã **tắt thông báo** nhóm đó) |
| Bảng tin | Bài mới đăng, văn bản cần xác nhận đã đọc; bình luận bài của bạn | Đối tượng của bài; tác giả bài |
| Ý kiến – kiến nghị | Ý kiến mới về phiếu lương; phản hồi | Người xử lý; người gửi ý kiến |
| Việc cần duyệt | Đơn nghỉ đến lượt bạn duyệt, kết quả duyệt đơn; kỳ lương chuyển đến bước của bạn; phiếu lương đã duyệt; đăng ký công tác chờ duyệt | Người xử lý bước; người làm đơn; toàn bộ CBCNV có phiếu lương |
| Chấm công – công tác | Công tác được duyệt / từ chối, được cử đi công tác | Người đi công tác |

**Quản trị cài đặt một lần** (Cài đặt → **Thông báo đẩy → Cấu hình**). Dùng Firebase Cloud Messaging của Google, miễn phí:
1. Vào console.firebase.google.com bằng tài khoản Google của Công ty → **Tạo dự án**.
2. Trong dự án, bấm biểu tượng **</> (Web)** → đặt tên → **Đăng ký**. Sao chép đoạn `firebaseConfig = {…}` dán vào ô 1.
3. **Cài đặt dự án → Cloud Messaging → Web Push certificates → Generate key pair**. Dán khóa vào ô 2.
4. **Cài đặt dự án → Tài khoản dịch vụ → Tạo khóa riêng tư mới**. Tải tệp JSON về, chọn tệp ở ô 3.
   - Khóa chỉ lưu trong **Thuộc tính tập lệnh** của Apps Script: không ghi vào Google Sheet, không gửi về trình duyệt.
   - Tệp JSON trên máy nên xóa sau khi cài; không gửi qua tin nhắn.
5. Tích **Bật thông báo đẩy** → **Lưu**.

**Mỗi CBCNV bật trên điện thoại của mình**
- App tự mời bật sau khi đăng nhập. Hoặc vào menu tên mình → **Cài đặt thông báo** → **Bật thông báo đẩy** → cho phép thông báo.
- Bấm **Gửi thử từ máy chủ** để kiểm tra.
- Có thể tắt từng loại thông báo, và ẩn nội dung trên màn hình khóa.
- **Android:** dùng Chrome; nên **Thêm vào màn hình chính**. Âm chuông chỉnh ở Cài đặt → Ứng dụng → Chrome/TAPACO → Thông báo.
- **iPhone (iOS 16.4 trở lên):** bắt buộc mở Safari → Chia sẻ → **Thêm vào MH chính**, mở TAPACO từ biểu tượng rồi mới bật được thông báo. Safari thường không nhận thông báo đẩy.
- **Máy dùng chung:** người đăng nhập sau nhận thông báo. **Đăng xuất** thì máy thôi nhận thông báo của người đó.

> **Lưu ý vận hành**
> - Mỗi thông báo tới 1 thiết bị là 1 lần gọi máy chủ. Apps Script của tài khoản Gmail thường giới hạn khoảng 20.000 lần/ngày, Google Workspace khoảng 100.000. Nhóm đông người nhắn liên tục sẽ tốn nhiều lượt.
> - Tin liên tiếp của cùng một nhóm trong 20 giây vẫn cập nhật nội dung thông báo nhưng **chỉ đổ chuông 1 lần**.
> - Thiết bị đã gỡ app hoặc hết hạn được tự xóa khỏi sheet **ThietBiPush**.
> - Nội dung xem trước (tiêu đề, một đoạn tin) đi qua máy chủ của Google để tới điện thoại. Ai không muốn có thể tắt **Hiện nội dung tin nhắn**.
> - Khi đang mở app, tin mới hiện banner nổi kèm tiếng "ting" thay cho thông báo hệ thống.

## 4u. Kiểm tra chấm công GPS bằng AI (v3.2)

Chỉ **Giám đốc, TCHC, Kế toán tiền lương, Kế toán công nợ, Quản trị** dùng được. Vào **Chấm công GPS → tab "🔍 Kiểm tra AI"**.

**1. Chạy kiểm tra**
- Chọn khoảng ngày (tối đa 3 tháng) và phạm vi (toàn công ty hoặc một bộ phận) → **Chạy kiểm tra**.
- Bộ quy tắc của hệ thống rà từng lượt chấm công và gắn cờ **ĐỎ** (sai phạm cần kiểm tra) hoặc **VÀNG** (cần lưu ý):

| Nhóm | Cảnh báo | Mức |
|---|---|---|
| Sai phạm vi | Chỉ có lượt ngoài vùng, không có lượt hợp lệ trong ngày | Đỏ |
| | Nhiều lượt thử ngoài vùng (≥ 3) | Vàng |
| | Lượt hợp lệ sát biên (≥ 90% bán kính) | Vàng |
| | Sai số định vị lớn (> 80% ngưỡng) | Vàng |
| | Sai số ≤ 2 m – bất thường, nghi ứng dụng giả lập vị trí | Đỏ |
| | Hai lượt liên tiếp cách nhau > 2 km với tốc độ > 150 km/h | Đỏ |
| Sai người | Khuôn mặt không khớp mẫu (kể cả lượt đã được duyệt hợp lệ) | Đỏ |
| | Một điện thoại chấm công cho ≥ 2 người | Đỏ |
| | Tín hiệu định vị bất thường | Đỏ |
| | Chấm trên thiết bị khác máy đã đăng ký; khuôn mặt sát ngưỡng; tính công dù chưa so khớp | Vàng |
| Sai thời gian | Trễ / về sớm > 60 phút | Đỏ (≤ 60 phút: vàng) |
| | Thiếu chấm ra / thiếu chấm vào | Đỏ |
| | Chấm công trong ngày có đơn nghỉ cả ngày đã duyệt | Đỏ |
| | Làm chưa đủ giờ; chấm Chủ nhật / lễ; ≥ 5 lượt/ngày; chờ duyệt quá 2 ngày | Vàng |
| | Ngày làm việc không chấm công, không có đơn nghỉ / công tác (người đã đăng ký khuôn mặt) | Vàng |

- Chạy lại nhiều lần **không tạo cảnh báo trùng**. Cảnh báo đã kết luận (Đã xử lý / Bỏ qua) không báo lại. Cảnh báo cũ mà lần quét mới không còn sai sót (vd lượt chờ đã được duyệt, đã bổ sung chấm công) **tự đóng**.
- Lượt chấm công có cảnh báo đang mở được **đánh dấu đỏ / vàng** ở các tab Trong ngày, Nhật ký & duyệt.

**2. AI phân tích & nhận xét**: bấm **AI phân tích**. AI (Gemini / Claude đã cài ở "Phân tích nhóm (AI)") chỉ nhận **số liệu tổng hợp và mô tả cảnh báo**, không nhận ảnh. AI trả về: tổng quan, mức rủi ro, nhận xét từng bộ phận, cá nhân cần chú ý, nguyên nhân có thể, đề xuất. Tên bộ phận / mã NV mà AI nêu được đối chiếu với số liệu (loại bỏ nếu không có). Không cài AI thì các cảnh báo vẫn dùng đầy đủ.

**3. Nhắc kiểm tra lại**: tích chọn cảnh báo (hoặc bấm **Nhắc kiểm tra lại** ở dòng bộ phận) → đặt **hạn kiểm tra**, lời nhắn → **Gửi nhắc**. **TCHC và Kế toán** nhận thông báo đẩy "⚠ Kiểm tra lại chấm công" để cùng kiểm tra; có thể chọn báo cả CBCNV liên quan. Từ bản 3.7.1 **không gửi cho Trưởng phòng / Chỉ huy trưởng**.

**4. TCHC / Kế toán kết luận**: bấm từng lượt xem ảnh, vị trí, khuôn mặt (liên hệ người lao động nếu cần) → tích chọn → **✍ Ghi kết quả & đóng**: chọn kết quả (*Đúng – không sai sót* / *Đã điều chỉnh, bổ sung* / *Sai sót – đề nghị xử lý theo quy chế*) và nội dung kiểm tra; cảnh báo chuyển **Đã xử lý**, lưu người kết luận và thời gian.

**5. Các thao tác khác**: **Đã xử lý** (đóng nhanh không ghi kết quả), **Mở lại** (ghi lý do), **Bỏ qua** (ghi lý do). **Xuất Excel** danh sách cảnh báo và bảng theo bộ phận.

Dữ liệu: sheet **KiemTraCC** (mỗi đợt kiểm tra, thống kê, nhận xét AI) và **CanhBaoCC** (từng cảnh báo, trạng thái, giải trình, xử lý).

## 4v. Kiểm tra vị trí khi nhắn tin (v3.3)

**Mục đích:** phát hiện CBCNV rời phạm vi chấm công trong giờ làm việc mà không đăng ký công tác.

**Cách hoạt động**
- Khi CBCNV gửi tin nhắn trong app (nhắn riêng hoặc nhóm), trong giờ làm việc, tối đa 1 lần mỗi `chuKyPhut` (mặc định 15 phút), app lấy vị trí (`getCurrentPosition`, độ chính xác cao, chấp nhận vị trí cũ ≤ 60 giây) và gọi `vitri.bao` sau khi tin đã gửi (không làm chậm tin nhắn).
- Máy chủ so với `gpsDiemCua(nv, hôm nay)`: điểm bộ phận, điểm dùng chung, điểm riêng, điểm công tác trong ngày. Lấy điểm gần nhất theo (khoảng cách − bán kính).
- Trong phạm vi (≤ bán kính + max(dung sai 50 m, sai số định vị)): **không lưu**.
- Ngoài phạm vi: ghi một dòng vào sheet **KiemTraCong** và báo đẩy cho tài khoản chức danh `tchc`, `giamdoc`, `ktluong`, `ktcongno`. Cùng người: báo lại sau `baoLaiPhut` (60 phút) hoặc khi di chuyển ≥ 500 m; lần đo giữa đợt vẫn ghi, cột `baoCao` để trống.
- Bỏ qua (`bo`): tắt; chức danh trong `boQuaVaiTro` (mặc định admin, giamdoc); ngoài giờ vào–ra, Chủ nhật, ngày lễ (khi `chiGioLam`); ngày nghỉ có đơn đã duyệt (cả ngày); công tác linh hoạt; sai số > `doChinhXacToiDa` (100 m); vị trí cũ hơn 30 phút; trong chu kỳ (CacheService); chưa có điểm chấm công.
- Minh bạch: lần đầu nhắn tin, mỗi tài khoản thấy hộp thông báo và bấm **Đã hiểu** (`vitri.xacnhan`, ghi NhatKy). Vị trí không hiện trong cuộc trò chuyện. App web chỉ đo khi người dùng đang mở app và gửi tin.

**Sheet KiemTraCong** (tạo bởi `upgradeSchema`): `id, ngayGhi, gioGhi, hoTen, maNV, viTriChamCong, viTriNgoaiPhamVi, khoangCach, gioDo` (8 cột theo yêu cầu: ngày ghi, giờ ghi, tên, mã NV, vị trí chấm công, vị trí xác định ngoài phạm vi, khoảng cách đo được, giờ xác định tại thời điểm đo), sau đó `boPhan, username, diemCC, latCC, lngCC, banKinh, lat, lng, doChinhXac, vuotBanKinh, banDo, hoiThoai, baoCao, ghiChu`. `gioDo` = giờ máy chủ − tuổi của vị trí do điện thoại báo.

**Giao diện:** Chấm công GPS → tab **📍 Ngoài phạm vi** (GĐ, TCHC, KT lương, KT công nợ, QT): lọc ngày / bộ phận, bản đồ, danh sách, **Xuất Excel** (cùng cột), **Mở sheet Kiểm tra công** (`env.sheetUrl`). Cài đặt: tab Cài đặt → "Kiểm tra vị trí khi nhắn tin" (`vitri.caidat`, TCHC / GĐ / QT).

**Pháp lý:** vị trí là dữ liệu cá nhân (Nghị định 13/2023/NĐ-CP). Bổ sung vào Quy chế chấm công, thông báo mục đích, phạm vi (chỉ giờ làm việc, chỉ lưu khi ngoài phạm vi), thời gian lưu, người được xem; lấy xác nhận của CBCNV trước khi bật.

## 4w. Đánh giá tổng hợp CBCNV bằng AI (v3.4)

**Ai dùng:** Giám đốc, TCHC, Quản trị (menu Quản lý → "Đánh giá CBCNV (AI)"). Nhân viên chỉ xem đánh giá **đã công bố** của mình trong "Hồ sơ của tôi" và phản hồi.

**Luồng:** `dgt.chay {tu, den, boPhan}` (≤ 186 ngày, đến ngày tương lai cắt về hôm nay) → hệ thống tính chỉ số + điểm cho từng người, ghi `DanhGiaTH` (đợt) và `DanhGiaCN` (từng người) → `dgt.ai {dotId, ids ≤ 6}` (chỉ đọc) + `dgt.luuai` lưu nhận xét; `dgt.ai {tong:true}` nhận xét chung → `dgt.chot {id, xepLoai, ketLuan, congBo}` → nhân viên `dgt.cuatoi`, `dgt.phanhoi` (báo TCHC, GĐ) → `dgt.baocao` (Google Docs + PDF trong `_ĐÁNH GIÁ CBCNV (KHÔNG CHIA SẺ)`), xuất Excel ở trình duyệt; `dgt.xoa` (không xóa được khi còn người đã công bố).

**Phạm vi người:** CBCNV có tài khoản hoặc có dữ liệu chấm công GPS, đang làm việc trong kỳ; bỏ tài khoản Giám đốc, Quản trị.

**Chỉ số (hệ thống tính, lưu trong `chiSo`):**
- Chấm công: tính từ `tinhTu` = max(đầu kỳ, ngày vào làm, ngày đầu tiên có chấm công GPS) đến hôm qua; ngày làm việc trừ Chủ nhật / lễ; ngày có công (lượt hợp lệ), nghỉ có đơn đã duyệt, công tác đã duyệt, ngày không có công, đi trễ (lần, phút), về sớm, thiếu chấm ra, sai vị trí, chờ duyệt, từ chối, khuôn mặt không khớp.
- Tuân thủ: cảnh báo `CanhBaoCC` trong kỳ (bỏ cảnh báo kết luận "Đúng – không sai sót" hoặc "Bỏ qua"), đỏ / vàng theo nhóm, chưa giải trình, sai sót đã kết luận; `KiemTraCong`: lượt, số ngày, xa nhất vượt bán kính.
- Nhóm công việc: mọi nhóm (`ltThongKe`): số tin, ảnh / tệp, ngày tương tác (hợp nhất các nhóm), báo cáo định kỳ phải nộp / đã nộp / đúng hạn; nhận xét cá nhân của tệp `PhanTich_<tu>_<den>_*.json` mới nhất trùng kỳ trong thư mục "4. Phân tích AI" của nhóm. **Không đọc nội dung tin nhắn.**

**Điểm:** chuyên cần = 100 × ngày có công / (ngày LV − nghỉ có đơn) − 2 × trễ − min(15, phút trễ / 20) − 2 × về sớm − 3 × thiếu chấm ra; tuân thủ = 100 − 8 × đỏ − 3 × vàng − 5 × sai vị trí − 4 × ngày ngoài phạm vi − 10 × sai sót − 2 × chưa giải trình; tương tác = tỷ lệ ngày tương tác (có báo cáo định kỳ: ½ + ½ tỷ lệ đúng hạn). Tổng = 40 / 35 / 25 % trên các tiêu chí có dữ liệu, cần ≥ 2 tiêu chí. Xếp loại theo điểm: ≥ 85 Tốt, ≥ 70 Khá, ≥ 55 Đạt, còn lại Cần cải thiện.

**AI:** hệ thống nhắc AI chỉ dùng số liệu, nêu dẫn chứng, không suy đoán đời tư, sai lệch vị trí là "cần làm rõ", số tin nhắn chỉ là mức tương tác. Máy chủ lọc mã NV / bộ phận AI bịa, xếp loại ngoài danh mục thay bằng xếp loại theo điểm, mỗi mục tối đa 5 ý. Không có khóa AI: điểm, chỉ số, chốt, công bố vẫn dùng được.

## 4x. Trợ lý quy định công ty trong Tin nhắn (v3.5)

**Ai dùng:** mọi CBCNV có tài khoản — vào **Tin nhắn**, mục **"Trợ lý quy định công ty"** ghim ở đầu danh sách. Quản lý nội dung: Giám đốc, TCHC, Quản trị (menu Công ty → "Quy định công ty (trợ lý)").

**Trợ lý trả lời bằng gì:** chỉ bằng hai nguồn do phòng TCHC đưa vào, không bằng kiến thức riêng của AI.
1. **Mục quy định** (sheet `BotQuyDinh`) — TCHC tự nhập: nhóm, tiêu đề (viết đúng như câu CBCNV hay hỏi), nội dung trả lời, căn cứ (số hiệu – tên văn bản), ngày hiệu lực, trạng thái còn/hết hiệu lực.
2. **Văn bản công ty** (sheet `BotVanBan`) — quy chế, quyết định đã tải lên Drive rồi đăng ký bằng đường dẫn tệp. Máy trích chữ và lưu đệm vào `_QUY ĐỊNH CÔNG TY (KHÔNG CHIA SẺ)/_trich-xuat/`. Google Tài liệu và tệp văn bản đọc trực tiếp; PDF, Word được chuyển sang Google Tài liệu tạm để lấy chữ rồi xóa bản tạm ngay.

**Luồng:** `bot.ds` (danh mục, câu hỏi gợi ý) → `bot.hoi {hoi}` (chỉ đọc, không giữ khóa ghi vì gọi AI chạy lâu) → trình duyệt ghi nhật ký bằng `bot.ghi` → CBCNV đánh giá bằng `bot.danhgia {id, huuIch}`. Quản lý nội dung: `bot.luu` / `bot.xoa` (mục quy định), `bot.vb.luu` / `bot.vb.xoa` (văn bản), `bot.nhatky` (nhật ký + báo cáo nội dung còn thiếu).

**Chọn nguồn đưa cho AI:** so khớp câu hỏi với từng mục quy định theo **âm tiết trọn vẹn** và **cụm 2 âm tiết** (không so chuỗi con — so chuỗi con làm "trợ" khớp trong "trong", "công" khớp trong "công ty" và trợ lý lấy nhầm căn cứ). Lấy tối đa 25 mục quy định và 3 văn bản liên quan nhất; văn bản dài chỉ lấy các đoạn liên quan, tổng nguồn giới hạn khoảng 52.000 ký tự.

**Bốn chốt chống bịa số liệu** (quan trọng nhất — CBCNV sẽ làm theo câu trả lời):
1. Prompt hệ thống cấm nêu bất kỳ con số, mức tiền, số ngày, tỷ lệ, thời hạn, số hiệu văn bản hay tên biểu mẫu nào không có trong phần NGUỒN; cấm suy ra từ luật chung hay thông lệ ngành; cấm dùng kiến thức ngoài NGUỒN.
2. Máy chủ **đối chiếu căn cứ AI nêu với danh sách nguồn đã cấp**; tên văn bản AI tự đặt bị loại bỏ.
3. Không còn căn cứ nào ⇒ coi như chưa có quy định.
4. Khi chưa có quy định, máy chủ **bỏ hẳn phần chữ AI sinh ra** và thay bằng câu trả lời chuẩn ("Công ty chưa có quy định về nội dung này… đề nghị liên hệ phòng TCHC"). Lý do: chữ không có căn cứ thì không kiểm chứng được, rất dễ lọt con số bịa ra ngoài.

**Chưa cài khóa AI:** trợ lý vẫn dùng được ở mức tra cứu — so câu hỏi với **tiêu đề** mục quy định, yêu cầu phủ ≥ 60 % âm tiết chủ đề và ≥ 2 âm tiết khớp, rồi trích **nguyên văn** nội dung TCHC đã nhập (không sinh chữ mới), ghi rõ "máy chủ chưa bật AI". Không đạt ngưỡng thì trả lời là chưa tìm thấy quy định phù hợp.

**Nhật ký và báo cáo nội dung còn thiếu:** mọi câu hỏi được ghi vào `BotHoi` (người hỏi, câu hỏi, kết quả, căn cứ, đánh giá hữu ích). Màn hình quản lý gom các câu **chưa trả lời được** theo độ giống nhau của tập âm tiết (Jaccard ≥ 0,6) thành danh sách "nội dung CBCNV hỏi mà chưa có quy định" kèm số lần hỏi — đây chính là việc cần làm của TCHC. Giữ 2.500 câu gần nhất.

**Việc TCHC cần làm trước khi cho CBCNV dùng:**
1. Vào "Quy định công ty (trợ lý)" → thẻ **Mục quy định** → nhập các câu hỏi thường gặp, **lấy đúng nội dung từ quy chế, quyết định công ty đã ban hành**. Ghi số hiệu văn bản vào ô Căn cứ.
2. Tải quy chế gốc lên Drive → thẻ **Văn bản công ty** → đăng ký bằng đường dẫn tệp. Tài khoản chạy app phải mở được tệp.
3. Hằng tuần xem thẻ **Câu hỏi của CBCNV** → bổ sung các nội dung còn thiếu.
4. Khi công ty sửa quy chế: chuyển mục cũ sang **Hết hiệu lực** (giữ lại để tra cứu) và nhập mục mới, thay vì sửa đè.

**Lưu ý:** trợ lý không trả lời việc riêng của từng người (lương cá nhân, số ngày phép còn lại, khiếu nại) — hướng người hỏi sang TCHC hoặc trưởng bộ phận. Câu trả lời mang tính tra cứu; văn bản công ty và ý kiến TCHC là căn cứ chính thức.

## 4y. Công văn đến – đi và Đề xuất nội bộ (v3.6)

Hai mục làm việc với văn bản, thay cho sổ công văn giấy và tờ trình giấy. Mọi tệp đều nằm trên **Google Drive của Công ty** — cùng tài khoản chứa file Google Sheet dữ liệu, trong thư mục riêng tư do app quản lý, **không chia sẻ cho CBCNV**; chỉ xem được qua app theo phân quyền và mỗi lần xem đều ghi nhật ký.

### Thư mục trên Drive

| Nội dung | Thư mục |
|---|---|
| Bản quét công văn đến, tệp kèm báo cáo kết quả | `_CÔNG VĂN / Đến / <năm>` |
| Bản quét, bản ban hành công văn đi | `_CÔNG VĂN / Đi / <năm>` |
| Nội dung đề xuất (tệp văn bản), ảnh và tệp đính kèm | `_ĐỀ XUẤT NỘI BỘ / <năm> / <mã đề xuất>` |

Tệp bản quét được đặt tên `<năm>-<số vào sổ>_<số ký hiệu>_<tên tệp gốc>` để tìm trên Drive vẫn ra đúng văn bản. Mỗi đề xuất có một thư mục riêng; tệp `NoiDung_<mã>.txt` trong đó được **ghi lại mỗi khi đề xuất thay đổi** (duyệt, từ chối, báo cáo kết quả), nên mở Drive là thấy toàn bộ hồ sơ đề xuất kể cả khi không vào app.

### Sổ công văn đến – đi

**Ai dùng:** Quản trị, Giám đốc, TCHC thấy toàn bộ sổ; TCHC và Quản trị là **văn thư** (đăng ký vào sổ, tải bản quét, sửa, lưu hồ sơ, xóa). Người được giao xử lý một văn bản cũng thấy mục này, nhưng **chỉ thấy đúng những văn bản giao cho mình**.

**Luồng công văn đến:** văn thư đăng ký vào sổ (số đến tự tăng theo năm, sửa được, không cho trùng số) → tải bản quét → lãnh đạo **bút phê**, chọn người xử lý chính, người phối hợp và **hạn xử lý** → người xử lý **báo cáo tiến độ / kết quả**, kèm được tệp (biên bản, văn bản trả lời) → văn thư **lưu hồ sơ**. Trạng thái: Mới đăng ký → Đã giao xử lý → Đang xử lý → Đã xử lý → Lưu hồ sơ.

**Luồng công văn đi:** đăng ký vào sổ (số đi đánh riêng, cũng theo năm) → tải bản đã ký; hệ thống tự chuyển sang **Đã ban hành** → lưu hồ sơ. Công văn đi không có bước giao xử lý.

**Quá hạn:** văn bản đã giao mà qua ngày hạn vẫn chưa báo cáo xong thì hiện nhãn đỏ *Quá hạn*; tab **"Việc của tôi"** của người xử lý hiện dòng cảnh báo tổng số văn bản quá hạn. Chấm đỏ trên menu đếm số văn bản đang chờ chính mình xử lý.

**Khóa sau khi lưu hồ sơ:** văn bản đã lưu hồ sơ thì không sửa, không giao lại, không xóa được (chỉ Quản trị sửa được trong trường hợp nhầm lẫn). Chỉ xóa được văn bản còn ở trạng thái mới đăng ký / chờ ban hành.

### Đề xuất nội bộ

**Ai dùng:** mọi CBCNV có tài khoản gắn hồ sơ. Mỗi người thấy: đề xuất của mình, đề xuất chờ mình duyệt, đề xuất giao mình thực hiện, đề xuất mình được cho theo dõi. Quản trị, Giám đốc và các vai trò xem toàn bộ thấy tất cả.

**Ba thứ phân loại một đề xuất:**

| | Là gì | Ai đặt ra |
|---|---|---|
| **Nhóm loại đề xuất** | Gom các loại cùng tính chất để dễ tìm, dễ lọc (Tài chính – thanh toán, Vật tư – thiết bị, Nhân sự – chế độ, Thi công – kỹ thuật, Khác) | **Quản trị** |
| **Loại đề xuất** | Quyết định **trình tự duyệt**, người theo dõi, người thực hiện, tiền tố mã, có số tiền hay không | **Quản trị** |
| **Bộ phận đề xuất** | Phòng ban hoặc công trình đứng tên đề xuất; lấy từ danh mục **Phòng ban & công trình** của Công ty | **Người lập** chọn khi lập (mặc định là bộ phận của mình) |

**Lập đề xuất:** chọn **loại đề xuất** (ô chọn gom sẵn theo *Nhóm › Loại*) → **bộ phận đề xuất** → nhập tiêu đề, nội dung, số tiền dự kiến (với loại có tiền), thời hạn cần có, độ khẩn. Ngay trong lúc nhập, màn hình cho biết **nhóm và tiền tố mã** của loại đang chọn, và **hiện trước trình tự duyệt** sẽ áp dụng — trình tự này **đổi theo số tiền** nếu cấp duyệt có đặt ngưỡng. Đính kèm được ảnh hiện trường, báo giá, bản vẽ ngay khi lập hoặc bổ sung sau (tối đa 20 tệp mỗi đề xuất). Lưu nháp được để hoàn thiện rồi trình sau.

**Mã đề xuất:** mỗi loại khai một **tiền tố** (2–8 chữ in hoa, không dấu) — mã là tiền tố + số thứ tự 4 chữ số chạy riêng cho tiền tố đó, ví dụ `TT0147`, `TUL0032`, `VT0268`. Loại không khai tiền tố thì mã chỉ là số thứ tự (`0004`). Cách đánh mã này giữ nguyên thói quen đánh số tờ trình theo loại của văn phòng.

**Tìm và lọc:** danh sách đề xuất lọc được theo **nhóm**, **loại**, **bộ phận đề xuất** và **trạng thái**; các thẻ *Chờ tôi duyệt · Tôi thực hiện · Tôi đề xuất · Tất cả* giữ nguyên bộ lọc đang đặt.

**Giám đốc là cấp duyệt cao nhất:** hệ thống **tự gắn cấp "Giám đốc phê duyệt" vào cuối mọi loại đề xuất** — Quản trị không cần khai và không bỏ được cấp này. Trong trình sửa, cấp đó hiện sẵn ở cuối với ghi chú "cấp chốt tự động". Nghĩa là mọi đề xuất, dù qua bao nhiêu cấp soát xét, đều kết thúc ở Giám đốc.

**Duyệt:** đề xuất đi lần lượt qua các cấp của loại đó. Cấp mà người duyệt chính là người đề nghị được bỏ qua; hai cấp liền nhau cùng người duyệt chỉ tính một lần. **Giám đốc và Quản trị duyệt vượt cấp** — duyệt ở bất kỳ cấp nào là đề xuất được duyệt luôn, màn hình báo trước điều này. Từ chối ở cấp nào cũng kết thúc đề xuất và bắt buộc nhập lý do. Không ai tự duyệt đề xuất của chính mình.

**Hai cách duyệt** (đặt riêng cho từng loại):
- **Duyệt theo trình tự** — đi lần lượt từng cấp, đúng như trên. Dùng cho hầu hết các loại.
- **Chỉ cần một người duyệt** — gộp tất cả người duyệt của loại đó (kể cả Giám đốc) thành **một cấp**; ai duyệt trước thì đề xuất được duyệt. Dùng cho việc nhỏ, cần nhanh.

**Thực hiện:** sau khi được duyệt, **người thực hiện** (cấu hình sẵn theo loại) nhận thông báo, cập nhật tiến độ và báo hoàn thành. Trạng thái: Nháp → Chờ duyệt → Đã duyệt → Đang thực hiện → Hoàn thành (hoặc Từ chối / Đã hủy).

**Người theo dõi, người thực hiện:** mỗi loại đề xuất khai sẵn trong cấu hình; người lập đề xuất và người duyệt **thêm được người khác** cho từng đề xuất cụ thể.

### Cài đặt nhóm và loại đề xuất — **chỉ Quản trị**

Vào **Quy trình & phê duyệt**. Giám đốc và các vai trò khác **xem được** hai mục này để nắm thẩm quyền, nhưng **chỉ Quản trị sửa và lưu được** (màn hình ghi rõ "Chỉ Quản trị sửa được"; máy chủ cũng chặn, không chỉ ẩn nút).

**Mục 4 – Nhóm loại đề xuất:** thêm, đổi tên, đổi thứ tự nhóm. Không xóa được nhóm còn loại đề xuất bên trong — chuyển các loại đó sang nhóm khác trước.

**Mục 5 – Loại đề xuất & trình tự duyệt.** Với mỗi loại khai:

| Mục | Ý nghĩa |
|---|---|
| Tên loại | Hiện cho CBCNV chọn khi lập đề xuất |
| Nhóm | Thuộc nhóm nào (mục 4) |
| Tiền tố mã | 2–8 chữ in hoa không dấu, ví dụ `TT`, `TUL`, `VT`. Không trùng giữa các loại. Để trống thì mã chỉ là số thứ tự |
| Có số tiền | Bật thì ô "Số tiền dự kiến" là bắt buộc và ngưỡng tiền ở các cấp mới có tác dụng |
| Cách duyệt | *Duyệt theo trình tự* hoặc *Chỉ cần một người duyệt* |
| Đang hoạt động | Tắt thì loại không còn hiện cho CBCNV chọn, nhưng các đề xuất cũ vẫn tra cứu được |
| Các cấp duyệt | Tên cấp, **ngưỡng số tiền** để cấp đó tham gia, và người duyệt: một **vai trò**, một **người cụ thể**, hoặc **trưởng bộ phận / chỉ huy trưởng của người đề nghị** |
| Người theo dõi | Được xem và nhận thông báo, không phải duyệt |
| Người thực hiện | Nhận việc sau khi đề xuất được duyệt, báo cáo kết quả |

Cấp **"Giám đốc phê duyệt"** luôn được gắn tự động vào cuối, không khai và không bỏ được.

Sáu loại dựng sẵn khi cài đặt lần đầu — **tên loại, nhóm, tiền tố và trình tự chỉ là mẫu để Công ty sửa lại theo quy chế của mình**, phần mềm không đặt ra mức tiền hay thẩm quyền nào:

| Nhóm | Loại | Tiền tố |
|---|---|---|
| Tài chính – thanh toán | Đề nghị thanh toán, tạm ứng | TT |
| Tài chính – thanh toán | Tạm ứng lương | TUL |
| Vật tư – thiết bị | Đề nghị cấp vật tư, thiết bị | VT |
| Nhân sự – chế độ | Đề nghị nhân sự (tuyển dụng, điều động, bổ nhiệm) | NS |
| Thi công – kỹ thuật | Đề xuất biện pháp, khối lượng phát sinh | TC |
| Khác | Đề xuất, kiến nghị khác | (không) |

Mọi ngưỡng số tiền mặc định để **0** (luôn phải qua cấp đó). Hãy đặt ngưỡng theo đúng phân cấp thẩm quyền trong quy chế tài chính của Công ty.

**Đề xuất đã trình giữ nguyên trình tự tại thời điểm trình** — sửa cấu hình về sau không làm lệch các đề xuất đang chạy.

## 4y2. AI: 4 nhà cung cấp, model tổng hợp, phân tích nhóm chạy song song (v3.15)

- **Nhà cung cấp:** Google Gemini, OpenAI GPT, Anthropic Claude, DeepSeek. Quản trị vào Phân tích nhóm (AI) → **Cài đặt AI**: chọn nhà cung cấp, dán khóa API (lưu được khóa của nhiều nhà cung cấp; chỉ nhà cung cấp đang chọn được dùng).
- **Model chính** (các đợt nhật ký, đọc ảnh, trợ lý, đánh giá…) và **model tổng hợp** (tùy chọn – bức tranh tổng thể, đánh giá nhân sự / giải pháp). Gợi ý: model nhanh làm model chính, model mạnh làm model tổng hợp.
- Nút **Lấy danh sách model**: hỏi nhà cung cấp danh sách model khóa API đang dùng được, đưa vào ô chọn. Model bị nhà cung cấp ngừng → máy chủ tự chọn bản thay thế cùng dòng (mini / haiku / sonnet / flash…), gọi lại và lưu.
- Mặc định: `gemini-flash-latest`, `gpt-5-mini`, `claude-haiku-4-5-20251001`, `deepseek-chat`. GPT (dòng gpt-5 / o): giảm mức suy luận (`reasoning_effort: low`) cho kịp giới hạn 60 giây; model không nhận tham số nào thì tự bỏ tham số đó.
- **DeepSeek không đọc ảnh:** việc cần đọc ảnh (CCCD, công văn) tự dùng nhà cung cấp khác đã có khóa; phân tích nhóm thì bỏ ảnh chứng từ. Máy chủ DeepSeek đặt tại Trung Quốc – cân nhắc với dữ liệu nhạy cảm.
- **Phân tích nhóm nhanh hơn:** các đợt nhật ký chạy **3 đợt cùng lúc**; nhật ký thi công (Google Docs, PDF) lập **song song** với bước tổng hợp → báo cáo; bước "Gộp kết quả" ghi tệp JSON cuối cùng. Ước tính với 7 đợt: khoảng 3,5 phút → 2 phút (tùy tốc độ model). Lỗi giữa chừng: bấm "Chạy lại từ đợt lỗi", các bước đã xong được giữ.

## 4z. Công văn: AI đọc văn bản, Giám đốc duyệt chia sẻ, trao đổi xử lý (v3.7)

### AI đọc bản quét và điền hộ vào sổ

Khi văn thư bấm **Vào sổ công văn đến / đi**, ô tải tệp nằm **ngay đầu form**. Chọn hoặc kéo thả bản quét (ảnh, PDF, Word) là app **tự gửi cho AI đọc** và điền vào các ô: **số ký hiệu, ngày văn bản, nơi gửi, nơi nhận (kính gửi), loại văn bản, độ khẩn, trích yếu, người ký, thời hạn trả lời**, kèm một đoạn **tóm tắt nội dung** 3–5 câu.

Nguyên tắc để không sai sổ:
- **AI chỉ ghi những gì nhìn thấy trong văn bản.** Văn bản không ghi hạn trả lời, hoặc bản quét mờ không đọc được, thì ô đó **để trống** — AI không đoán, không tự tính hạn. Loại văn bản và độ khẩn chỉ nhận đúng các giá trị có trong danh mục của sổ.
- **Không đè lên ô văn thư đã nhập.** Ô nào văn thư gõ (hoặc tự chọn loại, độ khẩn) trước khi AI đọc xong thì giữ nguyên.
- Các ô do AI điền được **tô vàng**, kèm mức **độ tin cậy** (cao / trung bình / thấp). Độ tin cậy *thấp* nghĩa là bản quét khó đọc — phải kiểm từng ô.
- **Văn thư vẫn chịu trách nhiệm soát lại trước khi bấm Lưu.** AI chỉ giúp gõ nhanh.
- PDF, Word được chuyển tạm sang Google Docs để lấy chữ rồi **xóa ngay** bản tạm; ảnh gửi thẳng cho AI. Tệp gốc chỉ lưu vào `_CÔNG VĂN` khi văn thư bấm Lưu.

Cần có khóa AI (Gemini hoặc Claude) — Quản trị cài ở **Phân tích nhóm (AI) → Cài đặt AI**, dùng chung cho mọi tính năng AI. Chưa có khóa thì form báo "Máy chủ chưa bật AI" và văn thư nhập tay như trước. Trên **bản dùng thử** (mở index.html không có máy chủ), AI trả về **dữ liệu mẫu có ghi rõ "NỘI DUNG MẪU"** chỉ để xem cách hoạt động.

Màn hình chi tiết văn bản hiện thêm **Hạn trả lời** và **Tóm tắt (AI)**.

### Giám đốc duyệt chia sẻ văn bản cho người liên quan

Ngoài những người được **giao xử lý** (bút phê), văn bản có thể cần cho thêm người khác đọc để nắm thông tin. Việc này đi qua **Giám đốc duyệt**:

1. Văn thư (hoặc Giám đốc, Quản trị) bấm **Trình chia sẻ** → chọn người được đọc (từng người, cả vai trò, hoặc trưởng bộ phận) và ghi lý do → trạng thái **Chờ duyệt**; Giám đốc nhận thông báo.
2. Giám đốc bấm **Duyệt chia sẻ** → xem tóm tắt, **sửa lại danh sách nếu cần** → **Duyệt & chia sẻ**. Những người trong danh sách **nhận thông báo và từ đó mới đọc được** văn bản cùng bản quét.
3. Không đồng ý thì bấm **Không chia sẻ** — **bắt buộc ghi lý do**; văn thư thấy lý do và có thể trình lại.
4. Giám đốc có thể chia sẻ thẳng mà không cần văn thư trình (nút **Chia sẻ**).

**Trước khi Giám đốc duyệt, người trong danh sách đề nghị không đọc được văn bản.** Người được giao xử lý thì vẫn đọc được như trước. Trong sổ, cột trạng thái hiện thêm nhãn *Chưa chia sẻ / Chờ duyệt / Đã chia sẻ / Không chia sẻ*.

### Trao đổi xử lý ngay dưới mỗi văn bản

Cuối màn hình chi tiết mỗi công văn đến/đi có mục **Trao đổi xử lý văn bản** — giống một nhóm chat riêng của văn bản đó:
- **Ai đọc được văn bản thì vào trao đổi được**: văn thư, lãnh đạo, người được giao xử lý, người được Giám đốc chia sẻ. Người khác không thấy.
- Gửi chữ và **đính kèm tệp** (dự thảo trả lời, ảnh hiện trường, bảng tính…). Tệp lưu cùng thư mục `_CÔNG VĂN / Đến|Đi / <năm>` của văn bản.
- Mỗi lời trao đổi báo cho những người liên quan. Trong sổ, mỗi dòng hiện **"N lời trao đổi"**.
- Văn bản đã **Lưu hồ sơ** thì khóa trao đổi (trừ Quản trị). Xóa văn bản thì xóa luôn trao đổi và tệp kèm.

## 5. Quy trình lập và phê duyệt lương hằng tháng

Mặc định mỗi kỳ lương đi qua 6 bước như bảng dưới. Bước, tên bước và người xử lý thay đổi được tại mục 4c. Chỉ người phụ trách bước hiện tại mới sửa được phần dữ liệu của bước đó; menu "Chấm công & Tính lương" hiện số kỳ đang chờ mình xử lý.

| Bước | Ai làm | Việc | Dữ liệu được sửa |
|---|---|---|---|
| 1. Chấm công | Kế toán công nợ | Chấm công theo ngày (mục 4k), giờ tăng ca ở tab Tổng hợp (hoặc nhập tổng / Excel / **Lấy công từ đơn nghỉ**) → **Gửi: Kiểm tra công nợ** | Ngày công, nghỉ HL/KL, OT, ghi chú chấm công |
| 2. Công nợ | Kế toán công nợ | Đối chiếu, kê tạm ứng theo loại (mục 4l), khấu trừ công nợ → **Chuyển TCHC duyệt** | Tạm ứng, khấu trừ công nợ |
| 3. TCHC duyệt công | TCHC | Duyệt chấm công, ghi khấu trừ theo quy chế (vi phạm, mất mát, bồi thường) → **Chuyển: Lập bảng lương** | Vi phạm – khấu trừ quy chế |
| 4. Tính lương | Kế toán tiền lương | Nhập truy lĩnh → **Tính lương** → kiểm tra → **Chuyển TCHC duyệt bảng lương** | Truy lĩnh, công chuẩn |
| 5. TCHC duyệt lương | TCHC | Kiểm tra bảng lương → **Trình Giám đốc phê duyệt** | (không sửa) |
| 6. Chờ duyệt | Giám đốc | Xem bảng lương, tổng hợp → **Phê duyệt & khóa** | (không sửa) |
| Đã duyệt | — | Kết quả lưu, khóa, có mã kiểm tra; xuất Excel, in bảng lương có chữ ký, in phiếu lương; CBCNV xem phiếu lương của mình | (khóa) |

- **Giám đốc phê duyệt ngay**: ở bất kỳ bước nào, Giám đốc bấm nút này để duyệt và khóa luôn, bỏ qua các bước còn lại. Hệ thống tính lại bảng lương theo dữ liệu chấm công hiện có và ghi "bỏ qua bước …" vào lịch sử phê duyệt.
- Từ bước Tính lương, mỗi lần lưu bảng lương được tính lại ngay, nên số liệu luôn khớp với dữ liệu mới nhất.
- **Trả lại**: người phụ trách bước hiện tại có thể trả về bất kỳ bước trước nào, bắt buộc ghi lý do. Lý do hiện ở đầu màn hình cho người nhận.
- **Mở khóa** kỳ đã phê duyệt: chỉ Quản trị, bắt buộc ghi lý do; kết quả phê duyệt và mã kiểm tra cũ bị hủy, CBCNV không xem được đến khi duyệt lại.
- Không chuyển được khỏi bước Tính lương nếu còn người đang làm việc trong kỳ mà chưa có chấm công.
- Khoản thưởng gắn kỳ lương phải được duyệt **trước khi trình Giám đốc**; sau đó hệ thống từ chối.
- Mọi lượt gửi, chuyển, trả lại, phê duyệt, mở khóa đều ghi vào sheet `PheDuyet` (tab "Lịch sử phê duyệt"). Bảng lương in ra lấy tên và thời gian xác nhận của từng bước làm phần ký.
- **Mã kiểm tra**: khi Giám đốc phê duyệt, hệ thống lưu số người, tổng thu nhập, tổng thực lĩnh và một mã 12 ký tự tính từ bảng lương. Nếu ai đó sửa tay số liệu trong Google Sheet, số trên bản in sẽ không còn khớp mã.
- Bảng lương xuất/in khi chưa duyệt đều ghi **"BẢN NHÁP – CHƯA PHÊ DUYỆT"**.

### Vi phạm theo quy chế

Bộ luật Lao động 2019 cấm phạt tiền, cắt lương thay cho xử lý kỷ luật (Điều 127). Vì vậy app khấu trừ vi phạm theo 2 hình thức hợp lệ, phải được quy định sẵn trong Nội quy / Quy chế lương thưởng:

1. **Giảm thưởng / phụ cấp theo quy chế** (vd. chuyên cần, an toàn lao động): trừ vào thu nhập, tối đa bằng thưởng + phụ cấp của tháng; phần thu nhập bị giảm cũng không tính thuế TNCN.
2. **Bồi thường thiệt hại** (làm hư hỏng, mất tài sản): khấu trừ sau thuế, tối đa 30% tiền lương thực trả sau BH và thuế (Điều 102). Phần vượt được ghi chú trên bảng lương để TCHC ghi tiếp ở kỳ sau.

Các khoản đã giữ lại qua nhiều kỳ lương, hoặc người lao động không đồng ý với mức bồi thường, nên được Kế toán trưởng và bộ phận pháp chế xem xét.

### Các tình huống thường gặp

| Tình huống | Cách làm |
|---|---|
| Kế toán phát hiện thiếu/sai công | Trả lại bước "Chấm công", ghi lý do; TCHC sửa rồi gửi lại |
| Giám đốc không đồng ý một khoản | Trả lại bước liên quan (vd "Công nợ"), ghi lý do |
| Cần sửa sau khi đã duyệt | Quản trị mở khóa về bước cần sửa; sau đó đi lại các bước |
| Có người mới vào giữa tháng sau khi TCHC đã gửi | Trả lại "Chấm công" để TCHC bổ sung |

### Ý kiến của CBCNV về phiếu lương

Sau khi kỳ lương được phê duyệt, mỗi CBCNV mở **Phiếu lương của tôi** → **Gửi ý kiến về phiếu lương này**, chọn nội dung và ghi đề nghị. Ý kiến tự chuyển tới bộ phận phụ trách:

| Nội dung | Bộ phận tiếp nhận |
|---|---|
| Ngày công, nghỉ, làm thêm giờ | TCHC |
| Lương, phụ cấp, thưởng | Kế toán tiền lương |
| Thuế TNCN, bảo hiểm | Kế toán tiền lương |
| Tạm ứng, công nợ | Kế toán công nợ |
| Vi phạm, khấu trừ theo quy chế | TCHC |
| Khác | TCHC |

- Bộ phận xử lý vào menu **Ý kiến CBCNV** (số trên menu = ý kiến chưa xử lý xong của bộ phận mình), mở ý kiến, bấm "Xem phiếu lương" để đối chiếu, rồi trả lời và chọn trạng thái: Đang xử lý / Đã giải quyết / Không chấp nhận. Có thể **chuyển** sang bộ phận khác (vd TCHC chuyển Kế toán lương). Đóng ý kiến bắt buộc ghi kết quả.
- Giám đốc và Quản trị xem, trả lời được mọi ý kiến. Trang kỳ lương đã duyệt hiện số ý kiến của kỳ đó.
- CBCNV thấy toàn bộ trao đổi ngay dưới phiếu lương; có phản hồi mới thì menu "Phiếu lương của tôi" hiện số. Chưa đồng ý với kết quả thì CBCNV ghi lý do để **mở lại**.
- Kỳ lương đã khóa không sửa trực tiếp. Nếu đề nghị đúng, khoản chênh lệch được ghi **truy lĩnh** (thiếu) hoặc **khấu trừ** (thừa) ở kỳ lương sau, và bộ phận xử lý nêu rõ trong câu trả lời. Trường hợp cần sửa ngay, Quản trị mở khóa kỳ lương.
- Toàn bộ trao đổi lưu ở sheet `YKien`. Mỗi CBCNV có tối đa 5 ý kiến chưa xử lý cho một kỳ, mỗi tin tối đa 2.000 ký tự.

### Xin nghỉ phép, nghỉ ốm, nghỉ không lương

| Loại nghỉ | Lương | Quỹ phép | Cột chấm công |
|---|---|---|---|
| Nghỉ phép năm | Công ty trả | Trừ quỹ phép năm | Nghỉ HL |
| Nghỉ việc riêng có lương (kết hôn, tang…) | Công ty trả | Không trừ | Nghỉ HL |
| Nghỉ ốm (BHXH chi trả) | Công ty không trả; BHXH chi trả theo hồ sơ | Không trừ | Nghỉ ốm |
| Nghỉ không lương | Không trả | Không trừ | Nghỉ KL |

**Quy trình đơn nghỉ**
1. CBCNV vào **Nghỉ phép của tôi** → **Tạo đơn xin nghỉ**: chọn loại, từ ngày – đến ngày (nghỉ 1 ngày có thể chọn buổi sáng/chiều), lý do, chứng từ. App tự đếm số ngày công nghỉ (không tính Chủ nhật và các ngày lễ khai trong Cài đặt) và báo phép năm còn lại. TCHC có thể **tạo đơn thay** người đã báo nghỉ qua điện thoại.
2. **Duyệt theo quy trình đã cài** (mặc định: **Trưởng phòng / Chỉ huy trưởng** của bộ phận người xin nghỉ → **TCHC**; bộ phận chưa có TP/CHT hoặc chính TP/CHT xin nghỉ thì TCHC duyệt; xem mục 4c). Khi tạo đơn, CBCNV thấy trước trình tự duyệt. Giám đốc duyệt vượt cấp được. Không ai tự duyệt đơn của chính mình. Từ chối bắt buộc ghi lý do.
3. CBCNV tự hủy được đơn đang chờ, hoặc đơn đã duyệt mà chưa tới ngày nghỉ. TCHC và Giám đốc hủy được mọi đơn.
4. Cuối tháng, ở bước **Chấm công**, Kế toán công nợ bấm **Lấy công từ đơn nghỉ**. App điền nghỉ HL / nghỉ ốm / nghỉ KL theo các đơn **đã duyệt** trong tháng và tính ngày công = công chuẩn − ngày nghỉ. Kiểm tra, sửa nếu cần rồi **Lưu**. App cảnh báo nếu còn đơn chờ duyệt trong tháng.

**Ràng buộc tự động**
- Không cho tạo đơn trùng thời gian với đơn đang chờ hoặc đã duyệt.
- Phép năm không đủ thì chặn và hướng dẫn xin nghỉ không lương cho phần vượt. Đơn phép năm không được kéo sang năm sau.
- Không tạo, duyệt hay hủy đơn rơi vào kỳ lương đã qua bước Chấm công (trả kỳ lương về bước Chấm công nếu cần sửa).
- Nghỉ ốm + nghỉ không lương từ 14 ngày làm việc trở lên trong tháng thì tháng đó không đóng BHXH, BHYT, BHTN.

**Quỹ phép năm** = 12 ngày (Điều 113 BLLĐ 2019) + 1 ngày cho mỗi 5 năm làm việc (Điều 114). Năm vào làm hoặc nghỉ việc tính theo tỷ lệ số tháng làm việc; vào làm sau ngày 15 thì không tính tháng đó. Với công việc nặng nhọc (14 ngày), phép tồn chuyển sang, hoặc thỏa thuận riêng, nhập **Số ngày phép năm (ghi đè)** trong hồ sơ. Màn hình **Nghỉ phép CBCNV** có 4 tab: Chờ duyệt, Tất cả đơn, Lịch nghỉ tháng (bảng màu theo từng người, từng ngày) và Quỹ phép năm của từng người.

> Nghỉ ốm: app chỉ ghi nhận ngày nghỉ và số chứng từ. Hồ sơ hưởng chế độ ốm đau gửi cơ quan BHXH vẫn làm theo quy trình hiện tại của TCHC. Nghỉ thai sản cũng ghi theo loại "Nghỉ ốm (BHXH chi trả)".

## 6. Công thức tính lương

### 6.1. Công thức lương do Quản trị lập (menu Hệ thống → Công thức lương)

Phần **thu nhập** gồm các **khoản lương** do Quản trị tự lập bằng công thức kiểu Excel.
- Khoản tính lần lượt từ trên xuống. Khoản sau dùng được kết quả khoản trước qua **mã** của nó.
- Phép tính: `+ − * / ( )`; so sánh `> < >= <= = <>`; hàm `MIN MAX ROUND FLOOR CEIL ABS IF AND OR NOT`. Dấu thập phân là dấu chấm (0.3).
- Mỗi khoản có các thiết lập sau:
  - **Loại**: thu nhập (lương / phụ cấp / tăng ca / thưởng / khác) hoặc khấu trừ khác sau thuế, vd đồng phục.
  - **Thuế**: chịu toàn bộ / không chịu / miễn tối đa X đ / phần chịu thuế theo công thức. Trong công thức phần chịu thuế, `GT` là giá trị của chính khoản đó.
  - **Được giảm trừ vi phạm quy chế** hay không.
  - **Tạm tắt** khoản đó.
- Ô **Chạy thử**: chọn một CBCNV và một kỳ để xem ngay từng khoản, thuế và thực lĩnh với công thức đang sửa, trước khi lưu.
- App kiểm tra lỗi cú pháp, biến chưa có, mã trùng; còn lỗi thì không cho lưu.
- Khi lưu, có thể đánh dấu **tính lại các kỳ lương đang xử lý**. Kỳ đã phê duyệt không bị ảnh hưởng. Mỗi kỳ lưu lại bộ công thức đã dùng để tính.

**Biến dùng được** (bấm vào ô công thức rồi bấm biến ở cột phải để chèn):
- Hệ thống:
  - Lương và ngày công: `LUONG_CHINH, LUONG_BH, LUONG_NGAY, LUONG_GIO, CONG_CHUAN, NGAY_CONG, NGHI_HL, NGAY_LE, NGHI_KL, NGHI_OM, CONG_HL`.
  - Làm thêm giờ: `OT150, OT200, OT300, HS_OT150…`; giờ được miễn thuế: `OT_GIO, OT_GIO_LUYKE, OT_GIO_MIEN, OT_TL_MIEN` (xem 6.1c).
  - Thưởng, truy lĩnh: `THUONG, THUONG_CHIU_THUE, TRUY_LINH`.
  - Thông tin người lao động và mức chung: `SO_NPT, THAM_NIEN_THANG, LA_THU_VIEC, LUONG_CO_SO, LUONG_TOI_THIEU_VUNG`.
- **Khoản cố định theo hồ sơ** `PC_<MÃ>`: mức của từng người nhập trong hồ sơ, vd `PC_NHAO`, `PC_DIENTHOAI`, `PC_XANGXE`, `PC_ANCA`, `PC_KPI`. Thêm khoản mới ở Cài đặt → "Khoản cố định theo hồ sơ".
- **Biến nhập theo kỳ**: số liệu riêng từng tháng, vd `OT_DEM` (giờ làm đêm, TCHC nhập), `KPI_PT` (% hoàn thành KPI, Kế toán lương nhập), số ca công trường xa… Mỗi biến thành một cột trong bảng chấm công ở đúng bước nhập.

**Nạp mẫu**: ô "Nạp mẫu…" có 2 bộ dựng sẵn, nạp xong vẫn sửa tự do rồi mới Lưu:
- **Bảng lương TAPACO 2026** (mặc định) — dựng lại đúng logic sheet "Tinh" của file `LƯƠNG CHUẨN VY 2026.xlsx` (xem 6.1b).
- **Chuẩn theo BLLĐ** — lương chính ÷ công chuẩn × công, phụ cấp theo công, OT theo hệ số 150/200/300%, làm đêm, KPI (bộ mẫu của bản 1.6).

**Giá trị trung gian**: loại khoản "Giá trị trung gian (không in phiếu)" dùng để tính các số phụ (công chuẩn, hệ số…) cho khoản sau dùng lại; không làm tròn, không cộng vào thu nhập. Một khoản trung gian có thể được chọn **hiển thị làm ngày công chuẩn** và một khoản **làm ngày công hưởng lương** trên bảng lương, phiếu lương.

**Nơi làm việc** (Cài đặt → Nơi làm việc): mặc định `VP` – Văn phòng (công chuẩn 24, ăn ca 30.000 đ/ngày) và `DA` – Dự án – công trường (công chuẩn 28, ăn ca 70.000 đ/ngày). Khai trong hồ sơ CBCNV; tháng nào làm nơi khác thì TCHC gõ mã vào cột **Nơi LV** khi chấm công. Biến: `NLV_VP`, `NLV_DA` (1 = đúng nơi làm việc trong kỳ), `CONG_CHUAN_NLV`, `TIEN_AN_NGAY`.

### 6.1c. Thuế TNCN và bảo hiểm đối với tiền làm thêm giờ (v3.12)

- **Bảo hiểm (BHXH, BHYT, BHTN, KPCĐ):** tính trên **lương đóng BH** (`LUONG_BH` ở mẫu TAPACO; lương chính + phụ cấp chức vụ, trách nhiệm ở mẫu chuẩn). Tiền làm thêm giờ **không** nằm trong lương đóng BH (Luật BHXH 2024 Điều 31, NĐ 158/2025: chỉ tính lương, phụ cấp và khoản bổ sung xác định được mức, trả thường xuyên). App đã đúng từ trước; bản 3.12 có kiểm thử khẳng định lương đóng BH không đổi dù có hay không có làm thêm giờ.
- **Thuế TNCN – quy định mới, áp dụng từ kỳ tính thuế 2026 với cá nhân cư trú:** Luật Thuế TNCN số 109/2025/QH15 (khoản 8 Điều 4) và Nghị định 253/2026/NĐ-CP (Điều 26) miễn thuế **toàn bộ** tiền lương làm thêm giờ, làm việc ban đêm, nếu làm thêm đúng điều kiện và giới hạn thời gian của pháp luật lao động; **phần vượt giới hạn tính vào thu nhập chịu thuế**. Trước đây (Luật 2007) chỉ miễn phần trả cao hơn lương giờ thường.
- **Cách app tính:** mỗi kỳ lấy tổng giờ làm thêm (`OT_GIO`) và số giờ đã làm các tháng trước trong năm (`OT_GIO_LUYKE`); số giờ được miễn `OT_GIO_MIEN` = không quá **giới hạn tháng** (mặc định 40 giờ) và không quá phần còn lại của **giới hạn năm** (mặc định 200 giờ; ngành, công việc được làm đến 300 giờ thì sửa ở Cài đặt). Khoản làm thêm giờ chịu thuế phần `GT × (1 − OT_TL_MIEN)`, với `OT_TL_MIEN = OT_GIO_MIEN ÷ OT_GIO` (giờ vượt chia theo tỷ lệ giờ). Kỳ trước 2026 vẫn theo quy định cũ (chỉ miễn phần trả cao hơn). Phụ cấp làm đêm miễn thuế.
- **Mẫu TAPACO:** tiền làm thêm giờ của người thử việc trước đây gộp vào công tính lương nên chịu thuế toàn bộ. Nay tách thành khoản "Lương làm thêm giờ" (tổng lương không đổi đến từng đồng, chỉ thuế giảm). HĐ chính thức theo file không trả tăng ca riêng (mục 5 dưới đây).
- **Công ty đang dùng bộ công thức cũ:** trang **Công thức lương** hiện khung vàng "Cập nhật thuế TNCN tiền làm thêm giờ" liệt kê những gì sẽ đổi → bấm **Cập nhật theo quy định mới**. Kỳ đã duyệt giữ nguyên; kỳ đang mở được tính lại. Nếu khoản `TONG_TG`/`HS_HUONG` đã bị sửa khác mẫu, app không tự vá mà báo cần sửa tay.
- **Lưu ý tuân thủ:** chấm công phải ghi đúng giờ làm thêm (TCHC / Kế toán công nợ nhập `OT150/OT200/OT300`) — bảng chấm công, bảng lương xuất từ app là căn cứ chứng minh theo NĐ 253/2026. App chưa kiểm giới hạn **theo ngày** (≤ 50% giờ làm bình thường, tổng ≤ 12 giờ/ngày); TCHC cần tự bảo đảm khi duyệt chấm công. Cách chia giờ vượt theo tỷ lệ là cách app chọn khi chưa có hướng dẫn chi tiết — Kế toán nên đối chiếu khi cơ quan thuế có hướng dẫn.

### 6.1d. Loại lương theo hợp đồng (v3.13)

Hồ sơ CBCNV → mục "Thuế – Bảo hiểm – Ngân hàng" → **Loại lương theo hợp đồng** (thay ô "Tham gia BHXH bắt buộc" cũ):

| Loại | Thu nhập | Bảo hiểm | Thuế TNCN |
|---|---|---|---|
| **1. Trả lương + đóng BH tại Công ty** (mặc định) | theo bộ công thức lương | NLĐ 8 / 1,5 / 1 %; DN 17,5 / 3 / 1 % + KPCĐ 2 % | lũy tiến (HĐ ≥ 3 tháng) hoặc khấu trừ 10 % (thử việc / dưới 3 tháng) |
| **2. Không đóng BH tại Công ty** – chọn lý do: | theo bộ công thức lương | NLĐ không đóng | như loại 1, nhưng không có BH để trừ |
| ‣ *Đóng BH nơi khác* | | DN chỉ đóng **BH TNLĐ-BNN 0,5 %** trên lương đóng BH (ghi ở cột BHXH DN) | |
| ‣ *Đã nghỉ hưu* | | không BHXH, BHYT, BHTN, KPCĐ | |
| **3. Lương khoán việc** | **Tiền khoán việc** nhập ở bước Tính lương (trống = "Mức khoán tháng" trong hồ sơ) + thưởng + truy lĩnh; không chạy công thức lương thời gian | không đóng BH cả hai bên | khấu trừ 10 % khi trả từ **5.000.000 đ/lần**; có cam kết 08 thì không khấu trừ; không giảm trừ gia cảnh |

- **Hồ sơ cũ** chưa chọn loại: "Tham gia BH = Không" được hiểu là loại 2 (ghi chú phiếu lương nhắc bổ sung lý do), còn lại là loại 1. App **không** tự chuyển ai sang khoán. Lưu hồ sơ thì cột "Tham gia BH" tự đồng bộ theo loại.
- **Khoán việc:** ở hồ sơ, ô lương đổi thành "Mức khoán tháng" — khoán cố định hằng tháng thì nhập số tiền; khoán theo khối lượng nghiệm thu thì để 0 và mỗi kỳ Kế toán lương nhập cột **Tiền khoán việc** (cột chỉ hiện khi có người hưởng lương khoán, chỉ mở ô cho đúng người đó). Nếu hồ sơ ghi HĐLĐ xác định / không xác định thời hạn mà chọn khoán, phiếu lương cảnh báo: HĐLĐ từ 1 tháng trở lên thuộc diện BHXH bắt buộc.
- **Cài đặt:** "BH TNLĐ-BNN – DN đóng cho người đã đóng BHXH nơi khác (%)", mặc định 0,5. Ngưỡng khấu trừ 10 % mặc định nay là **5.000.000** (NĐ 253/2026 Điều 50). Công ty đã lưu cấu hình từ trước vẫn giữ số cũ 2.000.000 — vào Cài đặt sửa thành 5.000.000 rồi Lưu.
- **Biến công thức mới:** `TIEN_KHOAN`, `LA_KHOAN` (1/0), `KHONG_BH` (1/0) — dùng khi Công ty muốn thêm khoản riêng cho các loại này.
- **Xuất Excel có công thức:** thêm cột "Loại lương"; người loại 2, 3 ghi giá trị BH và khoản khoán trực tiếp (không công thức), cột "Chênh lệch" vẫn bằng 0.
- **Cần Kế toán trưởng xác nhận:** (a) người đã nghỉ hưu: app không đóng BH TNLĐ-BNN; nếu Công ty áp dụng Điều 168.3 BLLĐ (trả thêm khoản tương đương mức đóng BH vào lương) thì thêm một khoản trong Công thức lương với điều kiện `KHONG_BH`; (b) người đóng BH nơi khác: theo Luật BHXH 2024, BHXH và BHYT đóng theo HĐLĐ giao kết đầu tiên, các HĐLĐ còn lại DN đóng BH TNLĐ-BNN; (c) mức 5.000.000 đ/lần và 10 % theo NĐ 253/2026.

### 6.1e. Quy chế tiền ăn ca (v3.13.1)

- **Mức ăn ca theo khối** (Cài đặt → Nơi làm việc, sửa được): Văn phòng (VP) **30.000 đ/ngày**, Dự án – công trường (DA) **70.000 đ/ngày**. Khối của mỗi người lấy theo "Nơi làm việc" trong hồ sơ; tháng nào làm ở nơi khác thì ghi cột "Nơi LV" trên bảng chấm công kỳ. Người có "Mức ăn ca riêng" (khai báo mức hưởng) vẫn ưu tiên mức riêng.
- **Chỉ ngày làm đủ 8 giờ mới có ăn ca; làm nửa ngày không có.** Biến mới `NGAY_DU_CA` = số ô **X** (làm cả ngày) và **CT** (công tác) trên bảng chấm công theo ngày. Ô ½, P½, R½, Ô½, KL½ không tính. Kỳ nhập tổng (không có bảng theo ngày) thì lấy phần nguyên ngày công (23,5 → 23).
- Khoản "Tiền ăn ca" (mẫu TAPACO) = `IF(NGAY_AN > 0, NGAY_AN, NGAY_DU_CA) × TIEN_AN_NGAY`. Ô "Ngày tính ăn ca" ở bước chấm công vẫn dùng được để ghi đè khi cần.
- Chấm công GPS (v3.13.2): giờ làm thực = giờ ra − giờ vào − phần trùng giờ nghỉ trưa (mặc định 12:00–13:30). Đủ **8 giờ** ("Số giờ làm đủ 1 công", cho phép thiếu bằng số phút "Cho phép trễ" – mặc định 10 phút) → **X** (đủ công, có ăn ca); thiếu giờ hoặc thiếu chấm công ra → **½** (nửa công, không ăn ca). Ví dụ VP 07:30–17:00 = 8 giờ; công trường 07:00–17:00 = 8,5 giờ; về 16:00 = 7 giờ → nửa công. Sửa ở Chấm công GPS → Cài đặt. Kiểm tra AI báo "Làm chưa đủ giờ" theo cùng cách tính.
- **Công ty đang dùng bộ công thức cũ:** trang Công thức lương hiện khung vàng "Áp dụng quy chế ăn ca" → bấm **Áp dụng quy chế ăn ca**. Kỳ đã duyệt giữ nguyên; kỳ đang mở được tính lại. Mẫu "Chuẩn theo BLLĐ" (ăn ca theo tháng `PC_ANCA`) không tự đổi.

### 6.1b. Mẫu "Bảng lương TAPACO 2026" – đối chiếu với file Excel

| File Excel (sheet Tinh) | Trong app |
|---|---|
| H Lương thỏa thuận | Hồ sơ → **Lương chính / thỏa thuận** (`LUONG_CHINH`), thay đổi qua quyết định lương |
| I Mức lương theo thời gian (sheet BHXH) | Hồ sơ → **Lương cơ bản đóng BH** (`LUONG_BH`); cũng là lương đóng BH |
| J, K, L, M Tiền nhà, xăng xe, điện thoại, chuyên cần (+ trách nhiệm) | Khoản cố định hồ sơ `PC_NHAO, PC_XANGXE, PC_DIENTHOAI, PC_CHUYENCAN, PC_TRACHNHIEM` |
| Các cột tháng 1–12 sheet "Mức lương", "BHXH" | **Khai báo mức hưởng theo thời gian** (mục 4d) |
| P Phạt | Tab **Vi phạm – quy chế** (bước TCHC kiểm tra) → giảm thưởng/phụ cấp |
| A VP/DA (sheet Công) | **Nơi làm việc** hồ sơ + cột Nơi LV khi chấm công |
| B FT/TV | Loại HĐ: Thử việc = TV, còn lại = FT |
| S Công chuẩn `IF(VP,24,28)` | khoản trung gian `CC` |
| T Công thời gian; U giờ CN; V giờ lễ; W giờ tăng ca | Ngày công; Giờ làm CN (`OT200`); Giờ làm lễ (`OT300`); Giờ tăng ca (`OT150`) |
| X, Y Công tính lương | `CONG_TL`: thử việc = công thực tế; DA làm 26–28 công tính đủ công chuẩn. Từ bản 3.12 phần giờ làm thêm quy đổi (CN×2 + lễ×4 + tăng ca×1,5)/8 của thử việc tách sang khoản **`TANGCA` – Lương làm thêm giờ** = `TONG_OT − TONG_TG` (tổng lương **không đổi** so với file) |
| Z, AA, AB Lễ, phép, việc riêng | Ngày lễ; Nghỉ HL |
| AD Hệ số K (1–1,2) | biến kỳ `HS_K`, Kế toán lương nhập ở bước Tính lương (trống = 1) |
| AE / AF Lương HĐ chính thức / thử việc | `TONG_TG = ROUND(LUONG_CHINH × tỷ lệ HĐ × K × CONG_TL ÷ CC + LUONG_BH × tỷ lệ HĐ ÷ CC × ngày nghỉ có lương, −2)`; tỷ lệ HĐ = 85% khi thử việc |
| (tách để in phiếu) | `LCB` lương cơ bản, `PCNO PCXX PCDT PCCC` phụ cấp, `LNGHI` lương lễ/phép, `THUONGNS` thưởng năng suất = phần còn lại (gồm làm tròn) — tổng các khoản luôn bằng `TONG_TG` |
| AG, AH Ăn ca | `ANCA = IF(NGAY_AN > 0, NGAY_AN, NGAY_CONG) × TIEN_AN_NGAY` (TIEN_AN_NGAY = mức ăn ca riêng của người đó nếu có, không thì theo nơi làm việc); miễn thuế tối đa 1.200.000 như file; `NGAY_AN` là biến kỳ TCHC nhập khi khác ngày công |
| AJ–AR Bảo hiểm | Tính theo luật trên `LUONG_BH`, có trần; cờ "Tham gia BHXH" trong hồ sơ |
| AS, AT, AU, AV Thuế | Giảm trừ 15,5 tr + 6,2 tr/NPT, biểu 5 bậc; thử việc khấu trừ 10% |
| AX Tạm ứng | Bước Công nợ |

Đã kiểm tra trên 29 dòng có số liệu của file: tổng lương thời gian, tiền ăn ca, thuế và thực nhận khớp đến từng đồng (6 dòng file cộng nhầm cột "ngày công tính ăn ca" vào tổng thu nhập, lệch 29,5 đ).

**Các điểm khác với file Excel cần Ban Giám đốc xác nhận:**
1. **Bảo hiểm khi làm dưới 14 công**: file trừ cả phần Công ty đóng (32,5%) vào lương người lao động. App làm theo luật: nghỉ không lương + ốm từ 14 ngày thì không đóng; người lao động chỉ chịu 10,5%.
2. **Phạt**: file trừ phạt vào tổng lương trước khi nhân hệ số K và chia công. App trừ đủ số tiền ở bước vi phạm, chỉ trừ vào thưởng/phụ cấp (Điều 127 BLLĐ cấm phạt tiền, trừ lương).
3. **Ăn ca miễn thuế 1.200.000 đ/tháng**: theo file; mẫu chuẩn cũ dùng 730.000 đ. Kế toán xác nhận căn cứ, sửa tại khoản ANCA nếu cần.
4. **Thử việc khấu trừ 10% khi từ 5 triệu**: từ v3.13 app mặc định 5 triệu, khớp file (NĐ 253/2026 Điều 50). Công ty đã lưu cấu hình trước đó thì sửa ở Cài đặt → "Ngưỡng khấu trừ (đ/lần chi trả)".
5. **Tăng ca của HĐ chính thức** không được trả riêng trong file (đã gồm trong hệ số K). Nếu muốn trả, thêm khoản TANGCA từ mẫu chuẩn.

**Ví dụ khác:** thử việc hưởng 85% lương: `LUONG_NGAY * NGAY_CONG * IF(LA_THU_VIEC, 0.85, 1)`; KPI dưới 70% không thưởng: `IF(KPI_PT < 70, 0, PC_KPI * KPI_PT / 100)`; phụ cấp thâm niên 1% mỗi năm tối đa 10%: `LUONG_CHINH * MIN(FLOOR(THAM_NIEN_THANG / 12), 10) / 100`.

**Lương đóng bảo hiểm** cũng là một công thức: mẫu TAPACO là `LUONG_BH`; mẫu chuẩn là `IF(LUONG_BH > 0, LUONG_BH, LUONG_CHINH + PC_CHUCVU + PC_TRACHNHIEM)`.

**Ngày lễ, Tết**: cột "Ngày lễ" trong chấm công. Nút "Lấy công từ đơn nghỉ" tự điền theo danh sách ngày lễ khai trong Cài đặt.

### 6.2. Phần tính theo luật (không sửa bằng công thức)

1. **Tổng thu nhập** = tổng các khoản thu nhập − giảm trừ vi phạm quy chế (tối đa bằng tổng các khoản được đánh dấu "được giảm trừ").
2. **Bảo hiểm**:
   - BHXH 8% và BHYT 1,5% tính trên lương đóng BH, tối đa 20 × lương cơ sở; BHTN 1% tối đa 20 × lương tối thiểu vùng.
   - Không đóng với loại lương 2, 3 (mục 6.1d), hoặc khi nghỉ không lương + nghỉ ốm từ 14 ngày trở lên trong tháng. Loại 2 – đóng BH nơi khác: DN chỉ đóng TNLĐ-BNN 0,5 %.
   - Đoàn phí 1%, chỉ với đoàn viên.
3. **Thu nhập chịu thuế** = tổng phần chịu thuế của từng khoản, trừ phần đã giảm trừ quy chế.
4. **Thuế TNCN**:
   - HĐ từ 3 tháng trở lên: (3) − BH NLĐ − 15.500.000 − 6.200.000 × số người phụ thuộc, áp biểu lũy tiến 5 bậc: đến 10 tr 5%; 10–30 tr 10%; 30–60 tr 20%; 60–100 tr 30%; trên 100 tr 35%.
   - HĐ thử việc / dưới 3 tháng / thời vụ và lương khoán việc: khấu trừ 10% nếu từ 5.000.000 đ/lần trở lên (NĐ 253/2026 Điều 50), trừ khi có cam kết 08.
5. **Thực lĩnh** = (1) − BH NLĐ − đoàn phí − thuế − bồi thường thiệt hại (≤ 30% lương thực trả) − tạm ứng − khấu trừ công nợ − các khoản khấu trừ khác trong công thức.
6. **Chi phí DN**: BHXH 17,5%, BHYT 3%, BHTN 1%, KPCĐ 2% trên lương đóng BH đã áp trần.

Mọi tỷ lệ, mức trần, giảm trừ, biểu thuế, hệ số OT sửa được trong **Cài đặt**.

**Nên để Kế toán trưởng xác nhận lại trước khi chạy thật**: mức miễn thuế tiền ăn giữa ca (730.000 đ), tỷ lệ/ngưỡng khấu trừ 10% với HĐ dưới 3 tháng theo Luật Thuế TNCN 2025, và mức tham chiếu tính trần BHXH theo Luật BHXH 2024.

## 7. Phân quyền

### Bảng phân quyền tích chọn (v3.8)

Vào **Hệ thống → Tài khoản**, cuối trang là **Bảng phân quyền**. **Giám đốc và Quản trị** tích chọn từng ô (chức năng × chức danh) rồi bấm **Lưu phân quyền**; người khác chỉ xem.

- Có 18 quyền tích chọn được, chia 4 nhóm: *Hồ sơ & nhân sự* (xem toàn bộ, sửa hồ sơ, tài liệu đính kèm, lập / duyệt đề xuất nâng lương – khen thưởng – kỷ luật), *Văn bản & thông tin* (Bảng tin, sổ công văn, văn thư), *Giám sát & AI* (Phân tích nhóm, Chấm công GPS – quản lý, Đánh giá tổng hợp, quy định của trợ lý), *Hệ thống* (tài khoản, đặt lại mật khẩu, nhật ký, quy trình, cài đặt, công thức lương).
- Ô vừa đổi được **tô vàng**; thanh trên và dưới bảng đếm số ô đã đổi, có **Hoàn tác** và **Về mặc định**. Khi lưu, hộp xác nhận liệt kê từng thay đổi và **cảnh báo** nếu cấp quyền nhạy cảm (xem toàn bộ hồ sơ – lương, tài khoản, cài đặt…) cho Trưởng phòng, Chỉ huy trưởng hoặc Nhân viên.
- **Áp dụng ngay** ở máy chủ (mọi thao tác kiểm tra quyền theo bảng mới); người đang mở app thấy menu thay đổi khi tải lại. Mỗi lần lưu ghi **nhật ký** ai đổi, cấp / bỏ quyền gì.
- **Ô khóa 🔒** luôn được giữ để không ai tự khóa mình khỏi hệ thống: Quản trị – tài khoản, cài đặt, công thức lương; Giám đốc – xem toàn bộ, phê duyệt; cả hai – cài đặt quy trình. Quyền **sửa bảng phân quyền** cố định cho Giám đốc và Quản trị.
- Các dòng **Lương & nghỉ phép** (ai chấm công, lập bảng, duyệt từng bước, duyệt đơn nghỉ) không tích ở đây mà đổi người làm ở **Quy trình & phê duyệt**. Các dòng cố định theo chức danh (xem nhân sự bộ phận mình; hồ sơ, phiếu lương của chính mình) cũng chỉ để xem.
- Quyền gắn theo **chức danh**: cấp cho một chức danh thì mọi tài khoản mang chức danh đó đều nhận. Ngoài Quản trị, người có quyền *Quản lý tài khoản* chỉ tạo, sửa được tài khoản Nhân viên.

Bảng dưới đây là **phân quyền mặc định** khi cài mới (bấm *Về mặc định* để quay lại):

| Chức năng | Quản trị | Giám đốc | TCHC | KT lương | KT công nợ | TP / CHT | Nhân viên |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Xem toàn bộ hồ sơ, lương, quá trình | ✓ | ✓ | ✓ | ✓ | ✓ | | |
| Hồ sơ, quá trình, phòng ban – công trình, điều động | ✓ | | ✓ | | | | |
| Lập đề xuất nâng lương, khen thưởng – kỷ luật | ✓ | | ✓ | | | | |
| Phê duyệt nâng lương, khen thưởng – kỷ luật | ✓ | ✓ | | | | | |
| Lương: chấm công, công nợ | ✓ | | | | ✓ | | |
| Lương: TCHC duyệt công, khấu trừ quy chế, duyệt bảng lương | ✓ | | ✓ | | | | |
| Lương: lập bảng lương | ✓ | | | ✓ | | | |
| Lương: phê duyệt & khóa (duyệt ngay ở mọi bước) | ✓ | ✓ | | | | | |
| Mở khóa kỳ đã duyệt | ✓ | | | | | | |
| Duyệt đơn nghỉ của bộ phận mình | | | | | | ✓ | |
| Duyệt đơn nghỉ (sau TP/CHT) | ✓ | ✓ vượt cấp | ✓ | | | | |
| Xem nhân sự bộ phận mình | | | | | | ✓ | |
| Xử lý ý kiến CBCNV | ✓ | ✓ | theo chủ đề | Lương, Thuế | Tạm ứng | | |
| Quản lý tài khoản, gán chức danh | ✓ | | chỉ Nhân viên | | | | |
| Cài đặt tham số, công thức lương | ✓ | | | | | | |
| Cài đặt quy trình & người phê duyệt | ✓ | ✓ | | | | | |
| Xem tài liệu đính kèm của mọi CBCNV | ✓ | ✓ | ✓ | ✓ | | | |
| Đăng Bảng tin | ✓ | ✓ | ✓ | | | | |
| Nhật ký thao tác | ✓ | ✓ | | | | | |
| Xem toàn bộ sổ công văn đến – đi | ✓ | ✓ | ✓ | | | | |
| Văn thư: đăng ký vào sổ, tải bản quét, lưu hồ sơ | ✓ | | ✓ | | | | |
| Bút phê, giao xử lý công văn kèm hạn | ✓ | ✓ | ✓ | | | | |
| Xử lý, báo cáo kết quả công văn được giao | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| **Phân tích nhóm (AI)**: lưu trữ, phân tích, xem báo cáo mọi nhóm | ✓ | ✓ | ✓ | ✓ | ✓ | | |
| Phân tích trò chuyện riêng | ✓ | ✓ | | | | | |
| **Chấm công GPS – quản lý**: xem lượt chấm, ảnh, bản đồ, duyệt lượt, duyệt công tác, kiểm tra AI, kết luận cảnh báo | ✓ | ✓ | ✓ | ✓ | ✓ | | |
| Khai báo điểm chấm công, cài đặt GPS | ✓ | ✓ | ✓ | | | | |
| Chấm công của chính mình, đăng ký công tác | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Dùng AI đọc bản quét khi vào sổ công văn | ✓ | | ✓ | | | | |
| Trình chia sẻ công văn | ✓ | ✓ | ✓ | | | | |
| **Duyệt** chia sẻ công văn cho người liên quan | ✓ | ✓ | | | | | |
| Đọc công văn được chia sẻ, trao đổi xử lý dưới văn bản | ✓ | ✓ | ✓ | khi được giao / chia sẻ | khi được giao / chia sẻ | khi được giao / chia sẻ | khi được giao / chia sẻ |
| Cài đặt nhóm, loại đề xuất & trình tự duyệt | ✓ | xem | | | | | |
| Duyệt cấp cuối cùng của mọi đề xuất | ✓ vượt cấp | ✓ | | | | | |
| Lập đề xuất nội bộ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Duyệt đề xuất | theo cấu hình từng loại | ✓ vượt cấp | theo cấu hình | theo cấu hình | theo cấu hình | ✓ bộ phận mình | |
| Hồ sơ, phiếu lương **của mình**, gửi ý kiến, xin nghỉ, tin nhắn | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

CBCNV chỉ nhận về máy dữ liệu của chính mình và phiếu lương các kỳ **đã phê duyệt**.

## 7b. Tốc độ khi chạy thật (v3.14)

Nguyên nhân chậm ở bản ≤ 3.13: mỗi lần ghi, máy chủ **xóa và ghi lại toàn bộ sheet** đã thay đổi. Sheet ChamCongGPS lớn dần mỗi ngày (150 người × 2 lượt × 26 ngày ≈ 7.800 dòng/tháng) nên mỗi lượt chấm công phải ghi lại hàng trăm nghìn ô, giữ khóa ghi hàng chục giây. Mọi thao tác ghi khác (lưu hồ sơ, chấm công, tạm ứng…) phải xếp hàng chờ.

Bản 3.14:
- **Ghi theo dòng:** chỉ ghi dòng đã sửa, thêm dòng mới vào cuối, xóa đúng dòng bị xóa. Chỉ ghi lại cả bảng khi thứ tự dòng bị đảo hoặc xóa trên 30 dòng một lần.
- **Đọc phần cuối:** chấm công GPS, xem công của tôi, duyệt lượt, tổng hợp công tháng, kiểm tra AI chỉ đọc các lượt từ mốc cần thiết (lượt ghi theo thời gian nên lượt mới luôn ở cuối sheet).
- **Mở app không giữ khóa ghi** (nhiều người mở cùng lúc không phải chờ nhau). Nếu lúc mở app có việc cần ghi (áp quyết định lương đến hạn), app mới giữ khóa và kiểm tra dòng trên sheet chưa bị người khác thay đổi trước khi ghi.
- Giao diện: lưu bộ nhớ trình duyệt gộp lại khi rảnh; hỏi tin nhắn mới mỗi 30 giây khi không mở Tin nhắn (2 phút khi app chạy nền) để bớt lượt chạy đồng thời của Apps Script.

Đo trên bảng tính giả lập cỡ thực tế (150 CBCNV, 6 tháng chấm công GPS ≈ 46.000 lượt; ước lượng theo tốc độ Google Sheets): chấm công GPS ≈ 38 giây → dưới 1 giây; duyệt 1 lượt ≈ 38 giây → dưới 1 giây; mở màn hình chấm công ≈ 7 → 2 giây; lưu hồ sơ ≈ 0,9 → 0,6 giây. Công cụ đo: `test/perf/perf_gas.js`.

**v3.15 – đăng nhập nhanh hơn:** đăng nhập không giữ khóa ghi; máy chủ chỉ gửi chấm công, bảng lương của **12 tháng gần nhất** (và các kỳ chưa duyệt) – dữ liệu gửi về giảm khoảng 3 lần với 150 CBCNV có 12 kỳ lương. Các kỳ đã duyệt cũ hơn được **tải ngầm** ngay sau khi vào app (lệnh `dl.cu`), không bắt chờ; đã có trên máy thì giữ lại, không tải lại.

**Lưu ý vận hành:** không chèn / sắp xếp tay các dòng trong sheet ChamCongGPS (sheet phải giữ thứ tự ghi theo thời gian). Cần sửa dữ liệu thì sửa nội dung ô, không đổi thứ tự dòng.

### Lưu nền & tin nhắn nhanh (v3.23)

- **Sửa hồ sơ CBCNV, thêm / sửa / xóa quá trình công tác**: bấm Lưu là form đóng ngay, thông tin mới hiện ngay; app gửi lên máy chủ chạy ngầm theo hàng đợi (đúng thứ tự thao tác). Góc dưới màn hình hiện "Đang lưu lên máy chủ (n thay đổi)…". Máy chủ báo lỗi (mất mạng, trùng mã…) → app hoàn lại dữ liệu cũ, báo lỗi và **mở lại form còn nguyên dữ liệu vừa nhập** để bấm lưu lại. Đóng trang khi còn thay đổi đang lưu thì trình duyệt hỏi lại.
- Thêm hồ sơ mới vẫn chờ 1 lượt (mã NV do máy chủ cấp) nhưng **không tải lại toàn bộ dữ liệu** như trước; lưu quá trình công tác, phòng ban, điều động, công ty thành viên cũng vậy – máy chủ trả kèm đúng phần dữ liệu thay đổi (trước đây mỗi lần lưu quá trình công tác tải lại cả app: 2 lượt máy chủ + vài MB).
- **Tin nhắn**: lượt hỏi tin mới khi không có gì thay đổi được máy chủ trả lời từ bộ nhớ đệm, không mở Google Sheet (4/5 lượt hỏi). Giảm mạnh tải cho Apps Script, các thao tác ghi (gửi tin, lưu hồ sơ…) không còn phải xếp hàng sau hàng loạt lượt hỏi của mọi người. Khi mở Tin nhắn, app hỏi tin mới 2 giây/lần (trước 3 giây). Gửi tin vẫn hiện ngay trên màn hình.

### Thêm hồ sơ, điều động lưu nền; công văn đọc bằng OCR (v3.24)

- **Thêm hồ sơ CBCNV**: bấm Lưu là mở ngay hồ sơ mới với mã gợi ý, lưu nền; mã vừa bị người khác dùng thì máy chủ tự cấp mã kế tiếp và báo.
- **Điều động CBCNV**: hồ sơ chuyển bộ phận ngay trên màn hình (nếu đã đến ngày), lưu nền; lỗi thì hoàn lại.
- **Công văn**: ảnh JPG/PNG, PDF → nhận dạng chữ tiếng Việt (OCR của Google Drive) → AI phân tích chữ, tóm tắt, điền ô. Mục "Chữ nhận dạng (OCR)" xem / sửa rồi "AI phân tích lại"; chữ OCR lưu kèm văn bản (nút "Nội dung văn bản (chữ OCR)"). Ảnh HEIC/WEBP hoặc quá mờ → AI đọc thẳng từ ảnh.

## 4m2. Quy định theo thời kỳ (trước / sau 01/7/2026) và bảng chấm công thêm giờ (v3.25)

**Quy định theo thời kỳ** – *Cài đặt → bảng "Quy định theo thời kỳ"*. Mỗi kỳ lương tự áp dòng có "Từ kỳ" gần nhất:

| Từ kỳ | Lương cơ sở | LTT vùng I | Giảm trừ bản thân / NPT | Biểu thuế | Miễn thuế tiền làm thêm giờ | Khấu trừ 10% (HĐ < 3 tháng, khoán) |
|---|---|---|---|---|---|---|
| 07/2024 | 2.340.000 (NĐ 73/2024) | 4.960.000 (NĐ 74/2024) | 11 / 4,4 triệu | 7 bậc | chỉ phần trả cao hơn giờ thường (TT 111/2013) | từ 2 triệu/lần |
| 01/2026 | 2.340.000 | 5.310.000 (NĐ 293/2025) | 15,5 / 6,2 triệu (NQ 110/2025) | 7 bậc | phần cao hơn | từ 2 triệu/lần |
| 07/2026 | 2.530.000 (NĐ 161/2026) | 5.310.000 | 15,5 / 6,2 triệu | 5 bậc (Luật 109/2025) | toàn bộ tiền làm thêm trong giới hạn giờ (NĐ 253/2026 Điều 26) | từ 5 triệu/lần |

- Trần BHXH, BHYT = 20 × lương cơ sở của kỳ (46,8 triệu trước 07/2026; 50,6 triệu từ 07/2026); trần BHTN = 20 × lương tối thiểu vùng của kỳ. Tỷ lệ đóng không đổi.
- Phiếu lương, bảng lương ghi dòng "Áp dụng: …" (cột BangLuong.cheDoLuat) và dòng "trong đó: tiền làm thêm giờ miễn thuế". Sheet ThamSo của bảng tính chi tiết Excel lấy đúng tham số của kỳ.
- Ô **"Áp biểu 5 bậc và miễn thuế toàn bộ tiền làm thêm giờ cho cả kỳ tính thuế 2026"**: Luật 109/2025 Điều 29.2 và NĐ 253/2026 ghi quy định về tiền lương "áp dụng từ kỳ tính thuế năm 2026". Mặc định không chọn (kỳ 01–06/2026 khấu trừ theo luật cũ, chênh lệch xử lý khi quyết toán năm). Kế toán trưởng quyết định; xem trước số chênh ở **Báo cáo lương → So sánh thuế 2026** (xuất Excel được).
- Kỳ đã phê duyệt không bị tính lại. Kỳ đang xử lý: bấm Tính lương lại để áp quy định.
- Sửa / thêm mốc khi Nhà nước điều chỉnh; "Khôi phục mốc theo luật" để điền lại 3 mốc trên.

**Bảng chấm công thêm giờ** – *Chấm công & Tính lương → tab "Chấm công thêm giờ"* (người xử lý bước Chấm công).
- Lưới theo ngày, nhóm theo bộ phận; gõ số giờ (2 hoặc 2,5) vào ô ngày, Enter / mũi tên để sang người kế tiếp. Loại giờ tự xếp theo ngày: **ngày thường 150%**, **nghỉ hằng tuần 200%** (Chủ nhật; chọn "Thứ 7 và Chủ nhật" ở Cài đặt → Ngày nghỉ hằng tuần), **lễ, Tết 300%** (theo Lịch nghỉ lễ).
- Cột 150% / 200% / 300% / tháng / lũy kế năm; tô đỏ ô vượt giờ trong ngày (ngày thường > 50% giờ làm việc = 4 giờ, mọi ngày > 12 giờ) và vượt giới hạn tháng, năm (Cài đặt: 40 giờ/tháng, 200 giờ/năm – Điều 107 BLLĐ, NĐ 145/2020). Phần giờ vượt giới hạn không được miễn thuế.
- Tự lưu chạy nền; tổng tự điền vào cột OT150 / OT200 / OT300 ở tab Tổng hợp (khóa nhập tay) – cộng cả giờ làm lễ, làm Chủ nhật đã tính ở bảng chấm công ngày. Không ghi giờ trước ngày vào làm / từ ngày chấm dứt HĐ.
- Xuất Excel "Bảng chấm công làm thêm giờ" có chỗ ký – dùng làm bảng kê thời gian làm thêm giờ để chứng minh miễn thuế (NĐ 253/2026).
- Lưu ý mẫu công thức lương TAPACO: tiền làm thêm giờ chỉ tính cho người **thử việc** (biến OT_QD); người chính thức không phát sinh tiền làm thêm theo công thức này. Muốn trả cho mọi người, sửa khoản "Lương làm thêm giờ" ở màn hình Công thức lương (hoặc dùng mẫu "chuẩn").

## 4m3. Năm trường hợp tính lương – bảo hiểm, thuế; trạng thái tính thuế theo thời điểm (v3.26)

Hồ sơ CBCNV → **Trường hợp tính lương** (chọn 1 trong 5):

| Trường hợp | Bảo hiểm | Thuế TNCN mặc định | Quyết toán |
|---|---|---|---|
| 1. HĐLĐ – đóng BH tại Công ty | NLĐ 8% + 1,5% + 1%; DN 17,5% + 3% + 1% + KPCĐ 2% | Lũy tiến, giảm trừ gia cảnh tại Công ty | Được ủy quyền Công ty |
| 2. HĐ khoán việc theo tháng | Không | Khấu trừ 10% khi trả từ ngưỡng (2 tr/lần trước 01/7/2026, 5 tr/lần từ 01/7/2026); có cam kết 08 + MST thì không khấu trừ | Cá nhân tự quyết toán |
| 3. HĐLĐ – đóng BHXH nơi khác | NLĐ không; DN chỉ BH TNLĐ-BNN 0,5% (BHXH, BHTN theo HĐLĐ đầu tiên) | Lũy tiến, **không** giảm trừ gia cảnh tại Công ty (giảm trừ chỉ ở 1 nơi) | Tự quyết toán (2 nguồn) |
| 4. HĐLĐ với người đã nghỉ hưu | Không thuộc diện BHXH bắt buộc (khoản 7 Điều 2 Luật BHXH 2024); TNLĐ-BNN theo Cài đặt (mặc định 0) | Lũy tiến, có giảm trừ; lương hưu không chịu thuế | Được ủy quyền |
| 5. HĐ thử việc – không đóng BH | Không | Khấu trừ 10% từ ngưỡng; cam kết 08 thì không khấu trừ | Tự quyết toán |

**Trạng thái tính thuế theo thời điểm** – hồ sơ → "trạng thái thuế theo thời điểm" (TCHC, Kế toán lương): mỗi dòng từ kỳ … đổi *Cách tính* (lũy tiến / khấu trừ 10%), *Giảm trừ gia cảnh tại Công ty*, *Cam kết 08*, *Ủy quyền quyết toán*. Ví dụ: thử việc trong HĐLĐ từ 3 tháng → lũy tiến ngay; cá nhân trường hợp 3 không đăng ký giảm trừ ở nơi khác → giảm trừ "Có" (thuế tạm khấu trừ giảm mạnh); cam kết 08 cho năm 2026 → "Có" từ 01/2026, "Không" từ 01/2027. Mỗi kỳ lương áp dòng gần nhất; phiếu lương ghi "Cách tính: …".

**Cài đặt**: *BH TNLĐ-BNN DN đóng cho người đã nghỉ hưu (%)*; *Trả thêm khoản tương đương mức đóng BH (Điều 168.3 BLLĐ)* cho người nghỉ hưu, HĐ thử việc – bật thì tự thêm khoản TRATHAYBH (= lương đóng BH × 21,5%, chịu thuế) vào Công thức lương.

**Báo cáo lương → Quyết toán TNCN năm**: từng người – TN chịu thuế, BH, giảm trừ năm, thuế cả năm (biểu năm theo quy định tháng 12), đã khấu trừ, **được hoàn / nộp thêm**, và gợi ý giảm thuế hợp pháp: ủy quyền Công ty quyết toán (mẫu 08/UQ-QTT-TNCN), làm cam kết 08/CK-TNCN, cá nhân tự quyết toán để được hoàn số đã khấu trừ 10%, đăng ký người phụ thuộc, cảnh báo cam kết không còn đúng / thiếu MST. Chọn giảm trừ "đủ 12 tháng" (quyết toán) hoặc "theo số tháng có lương" (ước tính giữa năm). Xuất Excel.

## 4m4. TAPACO Payroll & Tax 2026 – tối ưu thuế, cam kết, ủy quyền, ghi chú thuế – BHXH (v3.27)

**Vai trò mới "Phụ trách kế toán" (kttruong)** – quyền mới `duyetThue`: phê duyệt cam kết thu nhập, giấy ủy quyền quyết toán, hồ sơ thuế; cập nhật kho biểu mẫu và cơ sở dữ liệu quy định. Quy trình duyệt lương vẫn sửa được ở Quy trình & phê duyệt (thêm bước "Phụ trách kế toán kiểm tra" nếu cần: Nhân sự → Kế toán lương → Phụ trách KT → Giám đốc).

**Hồ sơ thuế – BH của từng người** (cột `NhanVien.hsThue`, nút "Hồ sơ thuế – BH" trong tab Tối ưu thuế TNCN): cư trú; thử việc (vị trí, ngày bắt đầu/kết thúc, lương công việc, tỷ lệ %, thử việc trong HĐLĐ); BH ở đơn vị khác (tên, mã đơn vị, ngày HĐLĐ đầu tiên, từng quỹ, tài liệu, nơi đăng ký giảm trừ bản thân); nghỉ hưu (đang hưởng, từ ngày); khoán (kinh doanh độc lập, 5 dấu hiệu quan hệ lao động); thu nhập nơi khác (lương HĐLĐ, vãng lai bình quân, chứng từ khấu trừ, yêu cầu quyết toán phần vãng lai, điều chuyển sáp nhập, tổng thu nhập và thuế nơi khác); dự kiến thu nhập năm; giảm trừ y tế (≤ 23 tr), giáo dục (≤ 24 tr), hưu trí tự nguyện (≤ 3 tr/tháng).

**Tab "Tối ưu thuế TNCN"** (hồ sơ CBCNV; NLĐ xem được của mình): 10 chỉ tiêu (thu nhập dự kiến năm, đã nhận, dự kiến còn lại, nơi khác, thuế đã khấu trừ, giảm trừ, BH & giảm trừ khác, thuế dự kiến, có thể hoàn, hồ sơ thiếu) và 3 phương án: A khấu trừ mặc định · B không khấu trừ theo cam kết hợp lệ (chỉ khi "Đủ điều kiện") · C quyết toán / hoàn. Tách rõ: tiền chưa bị khấu trừ (tạm hoãn) – thuế được giảm theo luật (NPT, y tế, giáo dục, hưu trí) – thuế nộp thừa được hoàn. Có máy tính Gross ↔ Net.

**Cam kết không khấu trừ 10%** – kiểm tra: thuộc diện khấu trừ 10%, cư trú, MST, chỉ 1 nguồn thu nhập, hồ sơ NPT, thu nhập dự kiến sau giảm trừ ≤ 0, bản ký đã duyệt, mẫu còn hiệu lực. Kết quả: Đủ điều kiện / Không đủ / Cần bổ sung / Chờ xác minh. Mẫu 08/CK-TNCN (TT 80/2021) đang ở trạng thái **Đang xác minh** (TT 80 được thay bằng TT 89/2026; chưa thấy mẫu cam kết trong danh mục TT 89 đã tra) → bản in có dấu "THAM KHẢO – MẪU CHƯA XÁC MINH HIỆU LỰC", hồ sơ gửi vào "Chờ xác minh mẫu", Phụ trách KT phải tích xác nhận mẫu được cơ quan thuế chấp nhận mới duyệt được. Duyệt xong app tự đặt trạng thái thuế "Cam kết: Có" từ 01 của năm, "Không" từ 01 năm sau.

**Ủy quyền quyết toán** (NĐ 253/2026 Điều 51): HĐLĐ từ 3 tháng, đang làm việc lúc quyết toán (hoặc điều chuyển do sáp nhập/cùng hệ thống), cư trú, MST, không có lương HĐLĐ nơi khác, vãng lai nơi khác bình quân ≤ 15 tr/tháng đã khấu trừ 10% và không yêu cầu quyết toán phần này, không đề nghị giảm trừ y tế/giáo dục. Tạo mẫu 08/UQ-QTT-TNCN (TT 89/2026) điền sẵn – In/PDF, Word; tải bản ký; Phụ trách KT duyệt.

**Bảng lương**: thêm cột Loại HĐ / trường hợp, Giảm trừ bản thân, Giảm trừ NPT, BH DN + KPCĐ, Tổng chi phí Công ty, Cam kết, Ủy quyền, **Ghi chú thuế – BHXH** (đỏ: cần xử lý, vàng: cần kiểm tra, xanh: thông tin – bấm để xem giải thích, căn cứ, hồ sơ cần bổ sung, biểu mẫu). Lọc theo trường hợp, chỉ dòng có cảnh báo; viền đỏ/vàng đầu dòng.

**Báo cáo lương**: tab "Ủy quyền & cam kết" (danh sách được ủy quyền, phải tự quyết toán, cần bổ sung; xuất Excel) và tab "Hồ sơ – biểu mẫu thuế" (8 biểu mẫu với mã, căn cứ, ngày áp dụng, phiên bản, tình trạng, đối tượng, hướng dẫn, link nguồn; cơ sở dữ liệu quy định LegalRules với văn bản, điều khoản, hiệu lực, trạng thái xác minh, URL). Quản trị / Phụ trách KT cập nhật được (chỉ nhận URL https).

**Khác**: tiền ăn giữa ca miễn thuế theo thời kỳ (1.200.000 đ/tháng từ 07/2026 – áp cho khoản ANCA đang để 730.000; mẫu TAPACO đã để 1.200.000 giữ nguyên); khoán "kinh doanh độc lập" không khấu trừ 10% tiền công; chặn giảm trừ bản thân 2 nơi; kiểm tra thử việc theo vị trí (180/60/30 ngày, 6 ngày làm việc) và lương ≥ 85%; nhật ký ghi giá trị trước → sau cho hồ sơ CBCNV, trạng thái thuế, hồ sơ thuế, cài đặt, quy định.

**Chưa làm được / giới hạn**: chưa mã hóa từng ô dữ liệu trong Google Sheet (bảo vệ bằng việc không chia sẻ Sheet/Drive và phân quyền); cá nhân không cư trú (20%) chỉ cảnh báo, chưa tự tính; thay đổi hợp đồng giữa tháng tính theo trạng thái của cả kỳ; biểu mẫu 02/QTT, 05/QTT, chứng từ khấu trừ chỉ có link nguồn (lập trên HTKK/eTax).

## 4m5. Khoán việc & thanh toán khoán – CONTRACT-02 (v3.28)

Menu **Quản lý → Khoán việc & thanh toán** (Giám đốc, TCHC, Kế toán, Phụ trách KT; Chỉ huy trưởng / Trưởng phòng thấy hợp đồng của công trình mình). Người nhận khoán là hồ sơ CBCNV loại HĐ "Thời vụ / Khoán việc" (người ngoài Công ty: tạo hồ sơ trước).

**Hợp đồng khoán** (bảng HopDongKhoan): số HĐ, người nhận, loại K1 cá nhân độc lập · K2 tổ đội · K3 bản chất HĐLĐ · K4 cá nhân / hộ kinh doanh; công trình, hạng mục, tổ; thời hạn; hình thức (tháng, ngày/ca, khối lượng, sản phẩm, giai đoạn); giá trị, đơn giá; GROSS / NET.
- **BM-01 – đánh giá bản chất** (11 câu hỏi: giờ làm, phân công, quản lý điều hành, tự chọn nhân công, biện pháp, trả theo nghiệm thu hay lương cố định, sửa chữa, rủi ro chi phí, tính lặp lại, chấm công – kỷ luật, tính nhất quán hồ sơ; câu hỏi đăng ký kinh doanh; câu hỏi tổ đội). Kết quả: Khoán độc lập có cơ sở · Có dấu hiệu QHLĐ – cần kiểm tra · Thuộc QHLĐ – chuyển HĐLĐ · Có dấu hiệu kinh doanh · Chưa đủ dữ liệu – tạm dừng phê duyệt. Luôn kèm lý do, căn cứ (BLLĐ Điều 13, Luật BHXH 2024 Điều 2); không kết luận theo số tháng hay 1 dấu hiệu.
- **BHXH**: không mặc nhiên "không đóng BH" – QHLĐ từ đủ 1 tháng thuộc diện BHXH bắt buộc.
- **Hồ sơ 9 thư mục** trên Drive `TAPACO / PAYROLL / NAM_yyyy / HOP_DONG_KHOAN / mã người /` (01 Hợp đồng … 09 Giải trình). Trạng thái: Chưa cung cấp, Đã tải lên, Chờ kiểm tra, Cần bổ sung, Hợp lệ, Không hợp lệ, Hết hiệu lực. Chỉ Kế toán / Phụ trách KT đặt "Hợp lệ"; 06_Cam_ket_thue không đặt "Hợp lệ" khi mẫu cam kết đang xác minh.

**Đợt thanh toán** (bảng ThanhToanKhoan): tạm ứng · nghiệm thu thanh toán · quyết toán HĐ. Tháng ghi nhận khối lượng tách với **ngày thực trả** (thuế tính theo ngày trả: ngưỡng 2 tr/lần trước 01/7/2026, 5 tr/lần từ 01/7/2026).
- Phải trả (gross) = giá trị nghiệm thu được xác nhận + thu nhập chịu thuế bổ sung − giảm trừ giá trị HĐ. Thuế = thu nhập thuộc diện khấu trừ × thuế suất. Thực nhận = phải trả − thuế − thu hồi tạm ứng − khấu trừ khác.
- Tạm ứng mặc định **không** tính thuế (tính khi nghiệm thu); chọn "có bản chất tiền công đã chi trả" thì tính ngay và lần thu hồi sau không tính lại (tránh thuế 2 lần).
- Không trừ giảm trừ bản thân trước khi khấu trừ 10%; thiếu MST không phải căn cứ miễn khấu trừ; cảnh báo đỏ khi nhiều khoản cùng tháng đều dưới ngưỡng nhưng tổng ≥ ngưỡng (chia nhỏ) và khi khoản vượt ngưỡng chưa xử lý thuế.
- Chế độ thuế: tiền công 10% · kinh doanh (K4 – không khấu trừ 10%, không giảm trừ bản thân; NĐ 68/2026) · không cư trú 20% (cần Phụ trách KT xác nhận) · HĐLĐ (chặn duyệt khoán).
- Tổ đội: TAPACO trả trực tiếp từng thành viên → thuế từng người; đội trưởng nhận trọn → thuế tính trên đội trưởng, không tự chia cho thành viên khi chưa có căn cứ.
- HĐ NET: nhập thực nhận thỏa thuận → app tính ngược giá trị gross.
- **6 bước duyệt**: Ban chỉ huy → Phòng KT-KH → Kế toán → Phụ trách KT → Giám đốc → Kế toán thanh toán (Đã chi). Bước Kế toán / Phụ trách KT kiểm tra hồ sơ 01, 04, 06; tính lại thuế theo trạng thái cam kết lúc duyệt. Lịch sử duyệt lưu từng bước.

**Cam kết không khấu trừ** (dùng chung Hồ sơ thuế – BH): ELIGIBLE / INELIGIBLE / PENDING / LEGAL_REVIEW / EXPIRED / REVOKED. Xét tổng thu nhập cả năm (TAPACO + nơi khác do cá nhân khai); có thu nhập 10% ở nhiều nơi không tự bị từ chối mà cộng gộp, có lương HĐLĐ lũy tiến nơi khác thì không đủ. Chỉ áp dụng khi bản cam kết năm đó đã ký, Phụ trách KT duyệt, có MST, chưa thu hồi / hết hạn. **Thay đổi cả ở bảng lương**: từ 3.28 bảng lương khoán cũng chỉ tạm không khấu trừ khi cam kết hợp lệ (Cài đặt `ckBatBuocDuyet`, mặc định bật).

**Quyền lợi thuế người nhận khoán**: 10 chỉ tiêu; dự báo A theo hợp đồng · B theo lịch sử · C cá nhân khai · D tổng hợp (không giả định TAPACO là toàn bộ thu nhập khi chưa có khai báo); giảm trừ bản thân 186 tr/năm, người phụ thuộc theo tháng đủ điều kiện có hồ sơ. Ủy quyền: chỉ có khoán độc lập → không được ủy quyền theo cơ chế HĐLĐ, hướng dẫn tự quyết toán; chuyển sang HĐLĐ ≥ 3 tháng → xét Điều 51 NĐ 253/2026, Phụ trách KT xác nhận phạm vi.

**Biểu mẫu**: BM-01 bản chất HĐ · BM-02 kê khai thu nhập dự kiến · BM-03 cam kết (mẫu tham khảo – cần xác minh) · BM-04 xác nhận khối lượng · BM-05 bảng thanh toán & khấu trừ · BM-06 tổng hợp thu nhập & thuế đã khấu trừ · BM-07 hướng dẫn quyết toán / hoàn thuế · BM-08 ủy quyền (chỉ khi đủ điều kiện). In/PDF, Word; bảng thanh toán xuất Excel. Link văn bản gốc (Cổng Chính phủ) do Công ty cung cấp được lưu với trạng thái "chưa kiểm tra" – Phụ trách KT mở kiểm tra rồi đổi trạng thái; không bật nút tải mẫu chính thức.

**Kết nối**: tiền khoán đã chi cộng vào quyết toán TNCN năm, tab Tối ưu thuế, ghi chú bảng lương. Số tiền thuế do quy tắc và công thức tính (kiểm toán được), không do AI quyết định.

## 4m6. Nhân sự hưởng lương hưu (v3.29)

**Xác định theo thực tế hưởng lương hưu, không theo tuổi.** Hồ sơ thuế – BH → "Người đã nghỉ hưu": tình trạng (Đang hưởng lương hưu / Đã đủ tuổi nhưng chưa hưởng / Hưởng trợ cấp khác – đã nhận BHXH một lần / Chưa xác minh), ngày bắt đầu hưởng, số quyết định, đơn vị chi trả, lương hưu hằng tháng, thẻ BHYT hưu, công việc nặng nhọc – nguy hiểm. Xác minh (Chờ / Đã xác minh / Không hợp lệ, ngày, người kiểm tra) chỉ do Kế toán lương / Phụ trách KT làm ở tab Tối ưu thuế TNCN, phải có hồ sơ 02 (quyết định hưởng lương hưu). Sửa tình trạng, ngày hưởng, số QĐ, đơn vị chi trả → về "Chờ".

**Bảo hiểm theo từng kỳ**: đang hưởng (đã xác minh) → BHXH, BHYT, BHTN của NLĐ và DN = 0, không trừ 10,5%, không tự trích BH TNLĐ-BNN. Đủ tuổi chưa hưởng / trợ cấp khác / hồ sơ không hợp lệ / kỳ trước tháng bắt đầu hưởng → tính như HĐLĐ đóng BH. Khai "đang hưởng" nhưng chưa xác minh → tạm không trích theo khai báo, ghi chú đỏ "Chưa đủ căn cứ xác định nghĩa vụ BHXH" và **Giám đốc không khóa (phê duyệt) được bảng lương** cho đến khi Kế toán xác minh.

**Tiền trả thêm tương đương bảo hiểm** (EMPLOYER_INSURANCE_CASH_EQUIVALENT, Điều 168.3 BLLĐ): = căn cứ × (17,5% + 3% + 1%, lấy theo tỷ lệ DN trong Cài đặt; BHXH, BHYT áp trần 20 × lương cơ sở, BHTN trần 20 × lương tối thiểu vùng của thời kỳ). Căn cứ: số khai riêng → lương cơ bản (mức đóng BH tham chiếu) → lương HĐ. Phương án A: cộng thêm vào lương thỏa thuận; phương án B: tổng chi trả đã gồm – app tách tiền lương theo công việc và khoản trả thêm, không cộng 21,5% lần hai, cảnh báo nếu lương theo công việc dưới lương tối thiểu vùng. Chưa chọn phương án → chưa tính, cảnh báo "Chưa rõ lương hợp đồng đã bao gồm tiền thay bảo hiểm hay chưa". Khoản này hiện thành dòng/cột riêng trên bảng lương, phiếu lương; chịu thuế TNCN toàn bộ; lưu căn cứ, tỷ lệ, kỳ hiệu lực quy định, căn cứ pháp lý (BangLuong.traThayBHCT).

**Thuế**: HĐLĐ từ 3 tháng → lũy tiến (giảm trừ bản thân 15,5 tr, NPT 6,2 tr; người nghỉ hưu vẫn được giảm trừ; không giảm trừ 2 nơi). Dưới 3 tháng → khấu trừ 10% từ 5 tr/lần, cam kết chỉ khi đủ điều kiện. Không cư trú → nhánh riêng (cảnh báo). Lương hưu do Quỹ BHXH trả là thu nhập miễn thuế, không cộng vào thu nhập TAPACO, không phải "thu nhập tiền lương nơi khác" khi xét ủy quyền. Hàm calculateRetireePIT trả chi tiết từng bậc.

**Ủy quyền** (evaluateRetireeTaxAuthorization): các điều kiện Điều 51 NĐ 253/2026 + câu hỏi "thuộc trường hợp phải tự quyết toán theo khoản 3 Điều 51" (nội dung đang xác minh – Kế toán trưởng trả lời). Không mặc định mọi người nghỉ hưu đều được ủy quyền.

**Phân tích tối ưu thuế** (tab Tối ưu thuế TNCN → Nhân sự hưởng lương hưu): kiểm tra giảm trừ bản thân, NPT chưa đủ hồ sơ, lương hưu không bị tính nhầm, khoản trả thêm tính đúng từng kỳ; A = thuế đã và sẽ khấu trừ cả năm, B = thuế năm; nộp thừa max(0, A − B), nộp thêm max(0, B − A); tách số thuế giảm hợp pháp với số chỉ được hoàn; hướng xử lý (Ủy quyền / Tự quyết toán / Bổ sung giảm trừ / Kiểm tra hồ sơ). Có bảng "Tính lương & thuế kỳ" với tổng tiền nhận từ hai nguồn.

**Bảng lương**: cột "Ghi chú pháp lý – nhân sự hưởng lương hưu" (13 loại ghi chú, bấm xem giải thích, căn cứ). **Hồ sơ hưu trí 10 mục** trên Drive `NhanSu / mã NV / HoSoHuuTri / năm /`. **Biểu mẫu** TAPACO-HT-01 Phụ lục thỏa thuận tiền trả thêm BH, TAPACO-HT-02 Bảng tính lương & thuế năm (In/PDF, Word); dùng chung mẫu 08/UQ, cam kết, kê khai.

## 4m7. Phiếu nhập / xuất kho vật tư trong Tin nhắn (v3.30)

Trong **nhóm** chat (nhóm công trình, nhóm kho): bấm **+** → **Phiếu nhập / xuất kho** (mọi thành viên nhóm lập được, thường là thủ kho / chỉ huy công trường).
- Chọn Phiếu nhập hoặc Phiếu xuất; ngày (không sau hôm nay); kho / công trình (mặc định tên nhóm); nhà cung cấp – người giao (nhập) hoặc đơn vị / tổ đội nhận (xuất); người nhận; hạng mục sử dụng (xuất); số chứng từ (hóa đơn, biên bản, phiếu giao hàng).
- Bảng vật tư: tên (gợi ý từ các phiếu trước của nhóm), quy cách, ĐVT, số lượng (nhận dấu phẩy thập phân, VD 1.250,5), đơn giá (không bắt buộc → thành tiền), ghi chú; tối đa 60 dòng.
- **Ảnh**: nút **Chụp ảnh** mở camera điện thoại; **Tải ảnh lên** chọn nhiều ảnh từ máy. Ảnh nén ≤ 1.600 px, mỗi ảnh có ô mô tả, tối đa 12 ảnh; lưu vào thư mục Drive của nhóm (Lưu trữ nhóm chat / <nhóm> / 1. Tệp & ảnh).
- Gửi → số phiếu tự cấp **PN-yyyymm-nnn / PX-yyyymm-nnn** theo nhóm; phiếu hiện thành thẻ trong cuộc trò chuyện (3 dòng đầu, ảnh thu nhỏ, nút Xem phiếu), có thông báo đẩy.
- Xem phiếu: chi tiết, ảnh lớn (trước / sau, tải về), **In / PDF** (mẫu phiếu có chỗ ký Người lập, Người giao / nhận, Thủ kho, Chỉ huy trưởng, kèm ảnh).
- Phiếu không thu hồi như tin thường; người lập hoặc trưởng / phó nhóm **Hủy phiếu** kèm lý do (ghi vào nhóm, giữ lịch sử, số phiếu không dùng lại) rồi lập phiếu mới.
- **Thông tin nhóm → tab Kho vật tư**: lọc theo tháng, nhập / xuất, gồm phiếu hủy; tổng hợp nhập – xuất theo vật tư (nhập − xuất trong kỳ), danh sách phiếu kèm ảnh; **Xuất Excel** 2 sheet (chi tiết phiếu, tổng hợp vật tư).
- Bảng PhieuKho (Google Sheet) tự tạo; nhật ký ghi lập / hủy phiếu. Chưa có: tồn kho đầu kỳ, thẻ kho, quy trình duyệt phiếu.

## 4m8. Sửa lỗi tiền làm thêm giờ – mẫu công thức TAPACO (v3.31)

**Lỗi:** khoản trung gian OT_QD của mẫu công thức TAPACO là `IF(LA_THU_VIEC, …, 0)` nên chỉ người **thử việc** được quy đổi giờ làm thêm ra tiền; CBCNV HĐLĐ chính thức có nhập giờ tăng ca (bảng chấm công ngày, "Chấm công thêm giờ", tab Tổng hợp) nhưng "Lương làm thêm giờ" = 0.
**Sửa:** chính thức quy đổi theo hệ số làm thêm ở Cài đặt (mặc định ngày thường ×1,5, nghỉ tuần ×2, lễ Tết ×3 – Điều 98 BLLĐ; lương ngày lễ đã trả riêng): công quy đổi = (OT150 × 1,5 + OT200 × 2 + OT300 × 3) ÷ giờ công ngày; tiền = lương thỏa thuận × K × công quy đổi ÷ công chuẩn (VP 24 / DA 28), làm tròn trăm đồng. Thử việc giữ cách cũ (85%, ×1,5 / ×2 / ×4). Miễn thuế: từ 07/2026 toàn bộ trong giới hạn giờ; trước 07/2026 chỉ phần chênh so với giờ thường.
Bộ công thức đã lưu trong Cài đặt được tự vá khi khoản OT_QD còn đúng công thức cũ (khoản đã sửa tay thì giữ nguyên – kiểm tra ở Công thức lương). **Kỳ đang lập phải bấm Tính lương lại**; kỳ đã phê duyệt không tự đổi – nếu cần, Quản trị mở khóa và tính lại, hoặc nhập truy lĩnh ở kỳ sau.

## 4m9. v3.32 – Chấm ăn ca, công nợ, 2 bảng lương, tăng tốc

**Bảng chấm công thêm giờ:** cột Thứ 7, Chủ nhật tô màu đậm riêng (Chủ nhật đậm hơn); ngày lễ màu cam.

**Bảng chấm ăn ca** (tab "Chấm ăn ca", người chấm công nhập ở bước Chấm công): mỗi ô ghi số bữa ăn trong ngày 0, 1, 2, 3 (gõ số là sang ngày kế; Enter / ↑ ↓ sang người khác). Tự lưu chạy nền. Có dòng "Tổng suất ăn" theo ngày, xuất Excel, nút "Điền 1 bữa theo ngày công".
- Cài đặt → Nơi làm việc: thêm cột **Phụ cấp ăn ca/bữa** và **Giá thu/bữa**.
- Phụ cấp ăn ca (thu nhập) = số bữa × phụ cấp/bữa khi nơi làm việc khai mức này; không khai thì vẫn tính theo ngày làm đủ 8 giờ như cũ.
- Tiền ăn thu = số bữa × giá thu/bữa → cột **Ăn ca** ở bước Công nợ, trừ vào lương.

**Kiểm tra công nợ:** ở bước Công nợ chỉ hiện các cột tiền: Tạm ứng · Ăn ca · Thu ứng · Thu khấu trừ khác · **Tổng nợ** (= cộng 4 khoản, tự cập nhật khi nhập) · Ghi chú. Ô "Hiện cột chấm công" để xem lại cột chấm công. Tạm ứng theo bảng kê; Ăn ca theo bảng chấm ăn (người không có bảng chấm ăn nhập tay); Thu ứng = thu hồi khoản ứng / vay kỳ trước. Phiếu lương, bảng lương, Excel có đủ các cột.

**Ăn ca miễn thuế TNCN** = MIN(mức tối đa của kỳ; ngày công thực tế × mức tối đa ÷ 22). Mức tối đa theo bảng Quy định theo thời kỳ (1.200.000 từ 07/2026; 730.000 trước đó). Cài đặt → "Ăn ca miễn thuế TNCN" chọn cố định mức tối đa nếu không muốn tính theo ngày công. Ví dụ: 11 công × 70.000 = 770.000; miễn 11 × 1.200.000 ÷ 22 = 600.000; chịu thuế 170.000.

**Chế độ bảng lương** (Cài đặt → Ngày công – làm thêm giờ):
1. **1 bảng**: lương theo công + lương làm thêm giờ, tăng ca (như trước).
2. **2 bảng độc lập**: Bảng 1 – lương theo công thời gian (không gồm tiền làm thêm giờ; BH, công nợ trừ ở đây); Bảng 2 – lương tăng ca, thêm giờ (giờ 150/200/300 %, tiền, thuế TNCN phân bổ, thực lĩnh). Thuế phân bổ = thuế cả tháng − thuế khi không có phần chịu thuế của tiền làm thêm giờ, nên Bảng 1 + Bảng 2 = thực lĩnh cả tháng. Excel xuất 4 sheet (2 bảng + 2 danh sách chuyển khoản); In bảng lương in đúng bảng đang xem; phiếu lương ghi tách 2 phần. Kỳ tính trước bản 3.32 chưa có số phân bổ: Bảng 2 lấy tiền làm thêm giờ, thuế nằm ở Bảng 1 – bấm Tính lương lại (kỳ chưa duyệt) để phân bổ.

**Tăng tốc duyệt các bước:** khi bấm Chuyển / Trình, số liệu đang nhập gửi cùng lệnh chuyển: 1 lần gọi máy chủ thay cho 2, và bảng lương chỉ tính một lần. Tính lại bảng lương cập nhật tại chỗ, Google Sheet chỉ ghi các dòng thay đổi (trước đây ghi lại cả sheet BangLuong). Chuyển bước không được (ví dụ còn người chưa chấm công) thì số đã nhập vẫn được lưu.

**Tăng tốc tin nhắn:**
- Gửi ảnh / tệp: tải lên và gửi tin trong 1 lần gọi; ghi Drive trước khi lấy khóa ghi nên người khác không phải chờ; gửi nhiều ảnh song song (3 cùng lúc).
- Mỗi ảnh có thêm bản thu nhỏ 480 px. Khung chat chỉ tải ảnh thu nhỏ, tải gộp 6 ảnh mỗi lần gọi và chỉ tải khi cuộn tới. Ảnh gốc chỉ tải khi mở xem. Ảnh cũ chưa có bản thu nhỏ thì dùng ảnh thu nhỏ của Drive.
- Máy chủ đọc ảnh theo đường nhanh: kiểm tra phiên và thành viên nhóm qua bộ nhớ đệm, đọc thẳng Drive, không mở Google Sheet.
- Ảnh đã xem lưu trên máy (Cache Storage của app) nên mở lại không tải lại; xóa khi đăng xuất.

## 4m10. v3.33 – Lịch sử hợp đồng lao động

**Hồ sơ CBCNV → tab "Hợp đồng lao động"** (TCHC, Quản trị):
- Mỗi lần ký hoặc đổi hợp đồng thêm **1 dòng mới**, không ghi đè. Mỗi dòng gồm: số HĐ, loại HĐ, từ ngày – đến ngày, lương HĐ, lương đóng BH, hình thức trả, file HĐ đính kèm.
  - Loại HĐ: thử việc / xác định thời hạn / không xác định thời hạn / HĐ dưới 3 tháng / thời vụ – khoán việc / cộng tác viên.
  - Hình thức trả: lương + đóng BH; đóng BH nơi khác / đã về hưu; lương khoán.
  - File HĐ được lưu vào Tài liệu hồ sơ, loại "Hợp đồng lao động".
- Nhập sai thì bấm **Hủy dòng**, ghi lý do. Dòng hủy vẫn lưu (gạch ngang) và không dùng để tính lương.
- Mỗi lần thêm dòng, app tự ghi một mục "Thay đổi HĐLĐ" vào Quá trình công tác.
- Người chưa có lịch sử HĐ dùng nút "Tạo dòng HĐ đầu tiên từ hồ sơ hiện tại". Khi chưa có dòng nào, bảng lương vẫn tính theo hồ sơ như trước.
- Khi đã có lịch sử HĐ, các ô loại HĐ, số HĐ, ngày ký / hết hạn, trường hợp tính lương trong "Sửa hồ sơ" chỉ để xem và tự cập nhật theo HĐ đang hiệu lực.
- Lương HĐ, lương đóng BH có hiệu lực từ ngày bắt đầu HĐ. Quyết định lên lương hoặc khai báo mức hưởng có ngày hiệu lực sau đó vẫn áp sau.

**Tính lương, thuế, BH theo HĐ của kỳ:**
- Bảng lương tháng nào lấy HĐ có hiệu lực trong tháng đó.
- Đổi HĐ giữa tháng thì chia giai đoạn:
  - Mỗi giai đoạn tính theo lương và loại HĐ của giai đoạn đó.
  - Ngày công tách theo bảng chấm công ngày; không có bảng ngày thì chia theo số ngày làm việc.
  - Thưởng, truy lĩnh, công nợ, vi phạm tính ở giai đoạn cuối.
- **BH:**
  - Thử việc, khoán việc – thời vụ, CTV: không trích BH bắt buộc (kể cả khi chọn "lương + đóng BH").
  - Đóng BH nơi khác / về hưu: không trích (DN chỉ đóng BH TNLĐ-BNN cho người đóng nơi khác).
  - HĐ còn lại: trích theo lương đóng BH.
  - Đổi HĐ giữa tháng: BH tính theo HĐ có hiệu lực cuối tháng, vì BH đóng theo tháng.
- **Thuế TNCN:**
  - HĐ từ 3 tháng trở lên: biểu lũy tiến, có giảm trừ gia cảnh.
  - HĐ dưới 3 tháng (theo thời hạn ghi trên HĐ), thử việc, khoán, CTV: khấu trừ theo tỷ lệ cố định.
  - Đổi HĐ giữa tháng: gộp thu nhập cả tháng theo từng cách tính.
- Phiếu lương ghi HĐ áp dụng; tháng có đổi HĐ thì ghi từng giai đoạn (ngày, số công, lương).
- Kỳ đã phê duyệt không tự tính lại; khi thêm HĐ ảnh hưởng kỳ đã duyệt, app báo để mở khóa hoặc truy lĩnh.

**Cảnh báo HĐ sắp hết hạn:**
- Cài đặt → "Báo TCHC trước khi HĐLĐ hết hạn" (mặc định 30 ngày).
- Danh sách CBCNV và trang Tổng quan hiện các HĐ sắp / đã hết hạn mà chưa có HĐ tiếp theo.
- Mỗi sáng (từ 7:00), máy chủ gửi thông báo đẩy cho vai trò TCHC, mỗi HĐ 1 lần.

**Quy định theo thời kỳ** (Cài đặt), mỗi mốc có "Từ kỳ":
- Thêm cột tỷ lệ khấu trừ cố định %, BH NLĐ (BHXH · BHYT · BHTN), BH DN (BHXH · BHYT · BHTN · KPCĐ · TNLĐ), bên cạnh lương cơ sở, lương tối thiểu vùng, giảm trừ gia cảnh, biểu thuế.
- Điền sẵn theo mức hiện hành: NLĐ 8 / 1,5 / 1; DN 17,5 / 3 / 1, KPCĐ 2, TNLĐ 0,5 (Luật BHXH 2024); khấu trừ 10%; biểu 5 bậc, giảm trừ 15,5 / 6,2 triệu, lương cơ sở 2.530.000 từ 07/2026.
- Luật đổi thì thêm mốc mới; các kỳ từ mốc đó tự áp dụng.

## 4m11. v3.34 – Cột Ăn ca ở bước Công nợ sửa được
- Cột Ăn ca tự lấy từ bảng chấm ăn ca (số bữa × giá 1 bữa), ô nền xanh, và **sửa được**.
- Ô đã sửa tay có nền cam, kèm ghi chú số theo bảng chấm ăn; bảng lương dùng số đã sửa. Chấm ăn lại không ghi đè số đã sửa.
- Xóa trống ô thì quay về số theo bảng chấm ăn.
- Người không có bảng chấm ăn: nhập tay như trước.
- Tổng nợ = Tạm ứng + Ăn ca + Thu ứng + Thu khấu trừ khác, trừ hết vào lương kỳ này.

## 4m12. v3.35 – In lý lịch cá nhân
- Hồ sơ CBCNV (TCHC, quản lý) và "Hồ sơ của tôi" (CBCNV) có nút **In lý lịch**. In khổ A4 dọc; trên điện thoại chọn "Lưu dưới dạng PDF".
- Các mục:
  - I. Thông tin chung: cá nhân, CCCD, liên hệ, trình độ, công ty, phòng ban, chức vụ, nơi làm việc, ngày vào, thâm niên, HĐ hiện hành.
  - II. Quá trình công tác.
  - III. Lịch sử hợp đồng lao động.
  - IV. Khen thưởng – kỷ luật (mặc định không in).
- Tùy chọn kèm thông tin lương (lương hiện hưởng, lương HĐ, diễn biến lương) và MST, số tài khoản – mặc định KHÔNG in. Khi không chọn in lương, số tiền ghi trong quá trình công tác cũng được ẩn.
- Phần cuối có ô ảnh 3×4, lời cam đoan, chỗ ký "Người khai" và "Xác nhận của Công ty".

## 4m13. v3.36 – Lý lịch: hồ sơ kèm theo, BH, thuế TNCN, hưu trí; xuất PDF và Word
Nút **In lý lịch** (hồ sơ CBCNV / Hồ sơ của tôi): chọn mục rồi bấm **In / Lưu PDF** hoặc **Tải Word (.docx)**.

Các mục được đánh số tự động theo mục đã chọn:
- I. Thông tin chung.
- II. Quá trình công tác.
- III. Lịch sử HĐLĐ.
- IV. Bảo hiểm xã hội, y tế, thất nghiệp: trường hợp tính lương, có tham gia BH bắt buộc tại Công ty không, tỷ lệ đóng của kỳ hiện tại, quá trình đóng BH tại Công ty theo bảng lương (từng giai đoạn, số tháng), đơn vị đóng BH nơi khác.
- V. Thuế TNCN: MST, cư trú, cách tính thuế kỳ hiện tại, số người phụ thuộc, cam kết 08, các hồ sơ thuế đã lập (cam kết, ủy quyền quyết toán, NPT…).
- VI. Hưu trí: tình trạng, xác minh, ngày hưởng, số QĐ, đơn vị chi trả; người không hưởng lương hưu ghi "Không thuộc diện".
- VII. Hồ sơ kèm theo: danh sách tài liệu trong hồ sơ (loại, tên, số hiệu, ngày, hết hạn, tình trạng) kèm **link mở trong app** (#tl=…). Người mở phải đăng nhập và có quyền xem; Drive Công ty không chia sẻ trực tiếp.
- VIII. Khen thưởng – kỷ luật (mặc định không in).

Tùy chọn **kèm thông tin lương** (mặc định không) để in thêm: lương HĐ, lương đóng BH, chi tiết BH 12 kỳ, thu nhập chịu thuế và thuế đã khấu trừ trong năm, mức lương hưu. Số tài khoản ngân hàng cũng là tùy chọn riêng.

File Word có bảng, link bấm được, chỗ ký "Xác nhận của Công ty" và "Người khai"; dùng phông Times New Roman, khổ A4.

## 4m14. v3.37 – Đọc chữ (OCR) tài liệu hồ sơ

**Quy trình:**
1. Tải tài liệu lên (hoặc ảnh CCCD mặt trước / mặt sau). Nếu là ảnh chụp hoặc PDF, app tự đọc chữ; trong form tải lên có ô bật / tắt, mặc định bật.
2. Văn bản đọc được lưu thành **tệp .txt riêng**, cùng thư mục với tệp gốc trên Drive. Tài liệu ghi "Đã đọc chữ"; nút **Văn bản OCR** để xem lại.
3. App nhận dạng loại giấy tờ và đề xuất điền:
   - CCCD: số CCCD, ngày sinh, giới tính, quê quán, nơi thường trú, ngày cấp, nơi cấp, hạn thẻ.
   - Bằng tốt nghiệp / chứng chỉ: trình độ, chuyên ngành, số hiệu, ngày cấp.
   - Thông báo MST: mã số thuế.
   - Mọi tài liệu: số hiệu, ngày ký.
4. Màn hình soát gồm 3 cột: Đang có | Đọc được (sửa được) | Cập nhật.
   - Ô đang trống được tích sẵn. Ô khác hồ sơ không tích sẵn.
   - Có cảnh báo khi họ tên khác hồ sơ hoặc số CCCD trùng người khác.
   - Bấm **Lưu vào lý lịch**. Lý lịch in / Word lấy ngay dữ liệu mới.
5. Chỉ TCHC (quyền sửa hồ sơ) được lưu vào lý lịch. CBCNV tự tải thì chỉ xem được văn bản đọc được; TCHC soát khi xác nhận.

**Nơi đọc chữ (không gửi dữ liệu cho AI bên ngoài):**
- **Google Drive OCR trên máy chủ Apps Script** (khuyến nghị: tốt với tiếng Việt, đọc cả PDF scan). Ảnh được chuyển tạm sang Google Docs trong Drive Công ty, lấy chữ rồi xóa bản tạm. Cách bật:
  - Mở dự án Apps Script → **Dịch vụ (Services) "+"** → chọn **Drive API** → **Thêm**.
  - **Triển khai → Quản lý bản triển khai → Phiên bản mới**.
  - Lần đầu chạy, cấp quyền Google Docs / Drive khi được hỏi.
- **Đọc trên máy** (Tesseract) khi máy chủ chưa bật Drive API: chạy ngay trong trình duyệt. Lần đầu tải bộ đọc tiếng Việt khoảng 10 MB, cần mạng. PDF có sẵn lớp chữ thì lấy chữ trực tiếp; PDF scan đọc tối đa 6 trang.
- Cài đặt → "Tự đọc chữ (OCR) ảnh / scan khi tải hồ sơ": Có / Không.
- Máy đọc chữ có thể sai dấu, sai số. Luôn soát với tài liệu gốc trước khi lưu.

## 4m15. v3.38 – Bảng phân quyền: Lương, Nghỉ phép, Theo chức danh tích chọn được

Tài khoản → Bảng phân quyền. Mặc định giữ đúng như trước.

**Lương — các bước lập & duyệt; Nghỉ phép — các cấp duyệt đơn**
- Mỗi bước của quy trình hiện hành là một dòng. Tích = giao bước đó cho chức danh (vai trò).
- Lưu xong, thay đổi ghi vào **Quy trình & phê duyệt** (cùng một dữ liệu). Có tùy chọn áp dụng cho kỳ lương đang xử lý; kỳ đã duyệt giữ nguyên.
- Không bỏ được người xử lý cuối cùng của một bước. Quản trị luôn xử lý được mọi bước (🔒). Không giao bước cho vai trò Nhân viên.
- Người giao riêng theo tài khoản, "Trưởng bộ phận của người nộp", thêm / bớt / đổi tên bước: vẫn làm ở Quy trình & phê duyệt (dòng ghi "thêm: …").
- Sửa các dòng này cần cả quyền "Sửa bảng phân quyền" và "Cài đặt quy trình & người duyệt".

**Theo chức danh**
- Xem nhân sự bộ phận mình: mặc định Trưởng phòng, Chỉ huy trưởng. Cấp cho chức danh khác thì người đó xem danh sách, quá trình công tác, đơn nghỉ của phòng ban mình (menu "Bộ phận của tôi"); không mở hồ sơ đầy đủ, lương của người khác.
- Hồ sơ, phiếu lương, đơn nghỉ, ý kiến của chính mình: luôn có cho mọi CBCNV (🔒 – Điều 95 BLLĐ: NLĐ được thông báo bảng kê trả lương).
- Sửa bảng phân quyền: mặc định Quản trị, Giám đốc (🔒, không bỏ được). Cấp thêm cho chức danh khác được (có cảnh báo – quyền rất nhạy cảm vì người đó tự cấp được mọi quyền); không cấp cho Nhân viên.
- "Về mặc định" đưa các quyền về như ban đầu; người làm từng bước Lương / Nghỉ phép giữ nguyên (khôi phục ở Quy trình & phê duyệt).

## 4m16. v3.39 – Ăn ca miễn thuế TNCN: công thức mở

Cài đặt → Ngày công – làm thêm giờ → **Ăn ca miễn thuế TNCN – công thức (sửa được)**.
- Mặc định: `MIN(MUC_MIEN, NGAY_CONG * MUC_MIEN / 22)` (như trước).
- Biến: MUC_MIEN (mức tối đa của thời kỳ – bảng Quy định theo thời kỳ: 730.000 trước 07/2026, 1.200.000 từ 07/2026), NGAY_CONG, SO_BUA, GT (tiền ăn ca của tháng) và mọi biến của công thức lương. Hàm MIN, MAX, ROUND, IF…
- Ví dụ: chia 26 công `MIN(MUC_MIEN, NGAY_CONG * MUC_MIEN / 26)`; theo đơn giá ngày `MIN(MUC_MIEN, NGAY_CONG * 40000)`; theo số bữa `MIN(MUC_MIEN, SO_BUA * 30000)`; cố định `MUC_MIEN`.
- Dưới ô có **Xem thử** với 22 / 15 / 5 công, cập nhật khi gõ; biến sai hoặc sai cú pháp báo ngay, máy chủ từ chối lưu.
- Kết quả luôn kẹp trong 0 – MUC_MIEN: công thức không làm miễn thuế vượt mức luật.
- Lưu ghi nhật ký (trước / sau). Bảng lương Excel xuất công thức tương ứng. Kỳ đã tính: bấm "Tính lại" để áp công thức mới.

## 4m17. v3.40 – Phiếu lương tách 2 phiếu khi chọn "2 bảng độc lập"

Khi Cài đặt → Chế độ bảng lương = 2 bảng độc lập và CBCNV có tiền làm thêm giờ trong kỳ, phiếu lương gồm 2 phiếu (xem, in / lưu PDF: mỗi phiếu 1 trang):
- **Phiếu 1/2 – Lương theo công thời gian:** thu nhập không gồm tiền làm thêm giờ; BH, công nợ, tạm ứng, khấu trừ trừ ở phiếu này. Thuế: thu nhập chịu thuế tính cả tháng (gồm tăng ca); dòng "Thuế TNCN cả tháng", "trừ: phân bổ cho tiền làm thêm giờ", "Thuế khấu trừ ở phiếu này". Cuối phiếu ghi còn nhận Phiếu 2 và tổng cả 2 phiếu.
- **Phiếu 2/2 – Lương làm thêm giờ, tăng ca:** giờ làm thêm ngày thường / ngày nghỉ / lễ; tiền làm thêm từng khoản (ghi phần miễn thuế); thuế TNCN phân bổ (= thuế cả tháng − thuế khi không có tăng ca); còn nhận.
- Còn nhận Phiếu 1 + Phiếu 2 = thực lĩnh cả tháng, khớp Bảng 1 và Bảng 2.
- Không có tiền làm thêm giờ, hoặc chế độ 1 bảng: 1 phiếu như cũ. CBCNV xem ở "Phiếu lương" cũng thấy 2 phiếu.

## 4m18. v3.41 – Mẫu phiếu lương & bảng lương (kéo thả)

Cài đặt → thẻ **Mẫu phiếu lương & bảng lương** (người có quyền Cài đặt). Kéo ⋮⋮ (hoặc ▲▼ trên điện thoại) để đổi chỗ; bỏ tích để ẩn; sửa chữ để đổi tên. Chỉ đổi cách trình bày, không đổi số tiền.

**Phiếu lương**
- Tiêu đề phiếu; tiêu đề Phiếu 1 / Phiếu 2 khi tách 2 bảng.
- Bố cục 15 khối từ trên xuống: đầu phiếu, tiêu đề, thẻ thông tin, hợp đồng, ngày công, Thu nhập, BH, Thuế TNCN, Khấu trừ khác, Còn nhận, bằng chữ, ghi chú, chuyển khoản, BH công ty đóng, chữ ký. Mục I, II… tự đánh số theo thứ tự hiện; dòng Còn nhận ghi công thức theo mục hiện.
- Luôn hiện (🔒): Tiêu đề – họ tên, Thu nhập, Còn nhận. Ẩn BH / Thuế / Khấu trừ: số còn nhận không đổi, phiếu ghi "đã trừ cả …" để NLĐ biết.
- Thẻ thông tin: chọn và sắp xếp (lương chính, lương BH, công, người phụ thuộc, bộ phận, chức vụ, ngày vào, MST, số TK, thực lĩnh).
- Chữ ký: mỗi dòng một người (mặc định Người lao động, Kế toán tiền lương).

**Bảng lương**
- Mỗi cột có 2 ô: **In** (in / PDF) và **Excel**. Mặc định như cũ.
- In / PDF: đúng thứ tự, ẩn / hiện, tên cột theo mẫu; thêm được Mã NV, Bộ phận, Công chuẩn, Lương chính, Đoàn phí, Tổng khấu trừ, Ngân hàng, Số TK.
- Excel có công thức: cột bỏ tích được **ẩn** (không xóa) nên công thức cộng vẫn đúng; đổi tên cột; nhóm khấu trừ theo thứ tự mẫu (Tổng khấu trừ tự lấy từ cột đầu đến cột cuối nhóm). Các khoản thu nhập theo thứ tự Công thức lương.
- Chữ ký cuối bảng: để trống = theo người duyệt của Quy trình lương (in) / Người lập biểu – Kế toán trưởng – Giám đốc (Excel).

**Xem thử** phiếu lương và bảng lương trước khi lưu (dùng dữ liệu kỳ đang chọn). **Lưu mẫu** ghi nhật ký; **Về mẫu mặc định** đưa về như cũ. CBCNV xem phiếu của mình theo mẫu đã lưu.

## 4m19. v3.42 – Ô chọn CBCNV gõ tìm; tìm công văn; tách tab Công nợ

**Ô chọn CBCNV gõ tìm (mọi màn hình)**: quá trình công tác, đề xuất lương, khai báo mức hưởng, đơn nghỉ thay CBCNV, khoán việc, biểu mẫu, bộ lọc "Tất cả CBCNV"… Ô danh sách CBCNV thành ô gõ có kính lúp:
- Bấm vào ô: hiện danh sách; gõ tên / mã / chức vụ / phòng ban → lọc ngay, người khớp nhất lên đầu.
- Gõ không dấu được ("nguyen van hop"), sai 1–2 ký tự vẫn ra ("ngyen"), viết tắt chữ đầu ("nvh" → Nguyễn Văn Hợp). Gõ số / mã: chỉ khớp đúng chuỗi.
- Chọn bằng chuột hoặc ↑ ↓ Enter; Esc đóng.

**Công văn đến – đi**: ô tìm dưới các tab Đến / Đi / Việc của tôi. Tìm theo số vào sổ, số – ký hiệu, nơi gửi / nhận, trích yếu, tóm tắt, người ký, loại, ghi chú, đoạn đầu chữ quét (OCR). Gần đúng như trên; văn bản khớp nhất lên đầu; tích "Mọi năm" (mặc định) để tìm cả các năm.

**Chấm công & Tính lương**: tab cũ "Tổng hợp công & công nợ" tách thành:
- **Tổng hợp công** (giữ chỗ cũ): cột chấm công, tính lương.
- **Công nợ** (sau tab Tạm ứng): Tạm ứng · Ăn ca · Thu ứng · Thu khấu trừ khác · Tổng nợ · Ghi chú. Kế toán công nợ nhập ở đây; ở tab kia có dòng nhắc chuyển tab.

## 4m20. v3.43 – App điện thoại (Android .apk; iPhone thêm vào màn hình chính)

Thư mục **5_AppDienThoai**: `TAPACO-NhanSu-x.yy.apk` (bản mới nhất) + `HUONG-DAN-CAI-APP.txt`.

**Android**: app "TAPACO Nhân sự" (gói vn.tapaco.nhansu, Android 7.0+). Giao diện `index.html` đóng gói trong app (đã gắn link máy chủ), chạy trong WebView tại địa chỉ an toàn https nên dùng được GPS, máy ảnh.
- Phát hành: tải .apk lên Drive Công ty, chia sẻ link → dán vào Cài đặt → **Link tải app Android (.apk)**. CBCNV xem ở menu tài khoản → **📱 Cài app điện thoại** (link + các bước cài; hướng dẫn iPhone).
- Trong app: tải Excel / Word / PDF → lưu Download/TAPACO và tự mở; In / Lưu PDF → hộp thoại in Android; PDF đính kèm → nút "Mở PDF trên điện thoại"; chụp ảnh hồ sơ → mở camera; nút Quay lại của máy lùi trong app; link ngoài mở bằng trình duyệt.
- Chưa có: thông báo đẩy khi tắt app (dùng Chrome nếu cần). Nâng cấp giao diện = phát hành .apk mới, cài đè.
- Khóa ký: gói nguồn, `android/khoa/` — giữ kín, sao lưu.

**iPhone / iPad**: Safari → Chia sẻ → Thêm vào MH chính (Apple không cho cài .apk).

**Chưa thử trên máy thật** — cài thử 1–2 máy Android trước khi phát hành rộng.

## 4m21. v3.44 – Sửa / xóa mức hưởng ngay trong hồ sơ

Hồ sơ CBCNV → tab **Lương & mức hưởng**:
- **Mức đang hưởng**: bấm vào một mức (✎) → mở "Khai báo thay đổi mức hưởng", con trỏ đặt sẵn ở khoản đó.
- **Diễn biến mức hưởng**: bấm vào **ô bất kỳ** để sửa khoản đó tại mốc đó:
  - mốc là khai báo → mở "Sửa khai báo" của mốc, chọn sẵn ô khoản đó (ô chưa có trong khai báo: nhập để thêm);
  - mốc từ đề xuất lương đã duyệt → hỏi lập khai báo mới cùng ngày; mốc theo HĐLĐ → chuyển tab Hợp đồng lao động;
  - dòng **"Trước … (mức gốc)"** → **Sửa mức gốc**: mức trước mốc thay đổi đầu tiên (ghi vào mức cũ của mốc đầu; khoản chưa có mốc thì sửa thẳng hồ sơ).
- Cuối mỗi dòng: **✎** sửa cả mốc, **✕** xóa mốc (khoản trở về mức trước đó). Trong hộp Sửa khai báo có nút **Xóa khai báo**; xóa trắng một ô để bỏ khoản đó khỏi mốc.
- Sửa / xóa khai báo đã hiệu lực và sửa mức gốc: **Quản trị, Giám đốc** (trước đây chỉ Quản trị). TCHC vẫn lập khai báo mới, sửa / xóa khai báo chờ duyệt.
- Lưu xong các kỳ lương chưa duyệt được tính lại; kỳ đã duyệt giữ nguyên. Mọi thay đổi ghi nhật ký (mh.save, mh.delete, mh.goc).

## 4m22. v3.45 – Tải ảnh nhanh (Bảng tin, Tin nhắn)

**Bảng tin công ty**
- Đăng ảnh: app tự thu nhỏ ảnh gốc ≤ 1600px và tạo thêm **ảnh thu nhỏ 640px** (≈ 40–60 KB) lưu cạnh ảnh gốc trên Drive.
- Danh sách bài: hiện ảnh thu nhỏ, **gộp tối đa 12 ảnh / 1 lần gọi máy chủ** (trước: mỗi ảnh 1 lần gọi, tải ảnh gốc), **chỉ tải khi cuộn tới**. Bài 1 ảnh: hiện ảnh thu nhỏ trước rồi thay ảnh nét. Bấm xem phóng to: ảnh gốc.
- Ảnh đã xem **lưu trên máy** (bộ nhớ app): mở lại không tải lại; xóa khi đăng xuất.
- Bài cũ (đăng trước 3.45) chưa có ảnh thu nhỏ riêng: máy chủ dùng ảnh thu nhỏ của Google Drive.

**Tin nhắn**: tải gộp 10 ảnh / lần gọi (trước 6).

**Máy chủ (Code.gs)**: ảnh thu nhỏ đọc từ Drive được giữ trong bộ nhớ đệm Apps Script 6 giờ (đọc Drive 0,3–1 giây / tệp → đọc đệm vài chục ms); ảnh thu nhỏ vừa tải lên được đệm sẵn. Chỉ tệp ≤ ~70 KB; tệp bị xóa → xóa khỏi đệm. Quyền xem kiểm tra như cũ.

## 4m23. v3.46 – Tự cắt, nắn thẳng ảnh căn cước

Hồ sơ CBCNV → **Ảnh căn cước công dân** → Tải ảnh lên (mặt trước / mặt sau), cả khi CBCNV tự tải:
- App tự tìm khung thẻ trong ảnh chụp (tách thẻ khỏi nền bàn, tay…), hiện **4 chấm góc** (chấm xanh = góc trên-trái của thẻ) và **ảnh sau khi cắt**.
- Lệch thì **kéo chấm** vào đúng 4 góc thẻ (chuột hoặc ngón tay); **⟳ Xoay 90°** nếu thẻ bị nằm dọc / ngược; **Tự nhận lại** để tìm lại.
- **Cắt và tải lên**: ảnh được nắn thẳng (khử nghiêng, khử phối cảnh) về đúng tỷ lệ thẻ 85,6 × 54 mm, 1280 × 807 px — rõ chữ cho OCR / AI đọc thẻ. **Dùng ảnh gốc**: tải ảnh nguyên (thu nhỏ ≤ 1600px). **Hủy**: không tải.
- Không nhận được khung (ảnh mờ, nền cùng màu thẻ): 4 chấm đặt sẵn gần mép ảnh, kéo tay.
- Xử lý hoàn toàn trên máy (trình duyệt / app), không gửi ảnh ra ngoài.
- Mẹo chụp: đặt thẻ trên nền tối, khác màu thẻ, chụp đủ 4 góc, không lóa đèn.

## 4m24. v3.47 – Ảnh thẻ 3×4; Bảng xác định hệ số năng suất tháng

**Ảnh thẻ (Hồ sơ CBCNV → Thông tin → Ảnh thẻ 3×4)**
- **Tải ảnh thẻ**: chọn ảnh chân dung → kéo khung 3×4 (kéo khung để di chuyển, kéo chấm góc dưới-phải để phóng to / thu nhỏ) → Lưu (600 × 800 px).
- **Lấy từ ảnh CCCD**: khung đặt sẵn ở ảnh chân dung bên trái thẻ CCCD mặt trước (ảnh đã nắn thẳng từ 3.46); chỉnh rồi lưu.
- Lưu như tài liệu hồ sơ loại "Ảnh thẻ" (riêng tư trên Drive, xem theo quyền, ghi nhật ký); thay ảnh thì xóa ảnh cũ. CBCNV tự tải được (chờ TCHC xác nhận) nếu được phép tự tải CCCD.
- Dùng làm **ảnh đại diện hồ sơ** và chèn vào **Lý lịch cá nhân** (In / PDF: ô ảnh 3×4; Word: ảnh bên phải phần họ tên). Chưa có ảnh thẻ → lý lịch **tự cắt ảnh chân dung trên CCCD mặt trước**; chưa có cả hai → ô trống "Ảnh 3 × 4".

**Hệ số năng suất (Chấm công & Tính lương → tab "Hệ số năng suất")**
- Nhập ở bước **Lập bảng lương** (người làm bước đó). Mỗi CBCNV: điểm 0–100 từng tiêu chí; **Chuyên cần** tự tính = (công + phép + lễ) ÷ công chuẩn (gõ để sửa).
- App tự tính **điểm** bình quân theo tỷ trọng → **xếp loại** → **hệ số**; sửa tay hệ số được (để trống = theo xếp loại). Nút "Điền 100 điểm cho ô trống".
- **Lưu bảng hệ số**: hệ số ghi vào biến kỳ của công thức lương (mẫu TAPACO: **HS_K** – hệ số K nhân vào lương), bảng lương tự tính lại. Lưu người chấm, thời điểm.
- **Xuất Excel** (cột Điểm có công thức SUMPRODUCT theo dòng tỷ trọng, dòng chữ ký) và **In bảng** (Người lập · Trưởng bộ phận · Kế toán trưởng · Giám đốc).
- **Cài đặt → Hệ số năng suất tháng**: biến nhận hệ số, tiêu chí + tỷ trọng (thêm / bớt, đánh dấu "chuyên cần theo công"), thang xếp loại → hệ số. Số đặt sẵn là **mẫu** (A ≥ 90 → 1,2; B ≥ 80 → 1,1; C ≥ 65 → 1; D → 1) – sửa theo quy chế lương, thưởng Công ty.

## 4m25. v3.48 – Thêm cột biến tính lương vào bảng lương

Cài đặt → Mẫu phiếu lương & bảng lương → **Bảng lương**: danh sách cột có thêm (ẩn sẵn, tích **In** / **Excel** để thêm vào bảng; kéo ⋮⋮ để đặt vị trí):
- Thu nhập: Lương đóng BH, Truy lĩnh, **Tiền ăn ca (nhận)**.
- Thuế: **TN không chịu thuế**, **Ăn ca miễn thuế**, **TN chịu thuế**, **Giảm trừ gia cảnh** (bản thân + người phụ thuộc), **TN tính thuế**, Thuế phân bổ tăng ca.
- Khấu trừ: BHXH / BHYT / BHTN (NLĐ) từng khoản, **Tổng nợ (công nợ)**, Tổng khấu trừ (đã có: Tạm ứng, Ăn ca, Thu ứng, KT khác, Bồi thường).
- Khác: Thực lĩnh tăng ca, BH công ty đóng, KPCĐ công ty đóng.
- **Mọi khoản của Công thức lương** (tên kèm "(CT)") và **biến kỳ** (vd "Hệ số thưởng K [HS_K]").
- Ô **Lọc cột** (gõ không dấu: "thue", "an ca", "giam tru", "tam ung"…) để tìm nhanh.
- In / PDF: hiện đúng các cột đã tích. Excel có công thức: các cột thêm ghi giá trị đã tính (không công thức), đặt trước cột "Ký nhận"; cột tiền có dòng tổng.

## 4m26. v3.49 – Cài app trên điện thoại (sửa APK, iPhone qua web https)
- **Lỗi APK 3.43–3.48**: apktool nén `resources.arsc` và không zipalign → Android 11+ (targetSdk ≥ 30) từ chối cài. Sửa: `apktool.yml` doNotCompress thêm `arsc`; `android/zipalign.py` ghi lại zip (resources.arsc, png STORED, đệm extra cho dữ liệu thẳng hàng 4 byte, .so 4096) **trước** khi ký; apksigner ký v1+v2+v3; `zipalign.py --kiem` chạy cuối `build_apk.sh` (sai thì dừng). Test `node test/t67.js <tệp.apk>`.
- **iPhone**: PWA cần https. Hướng dẫn `6_DangWeb_iPhone/HUONG-DAN-DANG-WEB.txt` (GitHub Pages / Netlify Drop). `cfg.linkWeb` (cfg.save chỉ nhận `https://…`, áp cả `linkAPK`); `caiAppDT()` hiện link + Sao chép, iPhone xếp phần iPhone lên trước, cảnh báo file://; `beforeinstallprompt` → `window._bip`, nút "Cài app ngay" (`caiNgayPWA`). Thẻ `apple-touch-icon` tĩnh `./icon-180.png` (bản file:// vẫn dùng icon vẽ canvas), `apple-mobile-web-app-status-bar-style=default`. sw.js SHELL thêm icon-180.png, CACHE `tapaco-hr-v349`.
- Chưa thử trên máy thật (đang xác minh).

## 4m27. v3.50 – Báo máy chủ Apps Script cũ hơn app
- `login` và `load` trả `data.svVer = VERSION`. Client `soPB(v)`, `mcCu(can)` (thiếu svVer = máy chủ trước 3.50), `mcBao()` toast đỏ 1 lần/phiên cho admin.
- Lưu ảnh thẻ: máy chủ < 3.47 không nhận `anhThe` → lỗi "Chọn loại tài liệu." → client đổi thành hướng dẫn cập nhật Code.gs (Quản lý bản triển khai → Sửa → Phiên bản mới). Toast lỗi dài hiển thị lâu hơn (70 ms/ký tự, tối thiểu 6 s). Test smoke74.

## 4m28. v3.51 – Nguồn webapp.txt trên GitHub
- `WEBAPP_NGUON` (client) = [raw.githubusercontent.com/vuhuylong1976-maker/HCNS/refs/heads/main/webapp.txt]. `docWebapp()`: thử GitHub trước (thêm `?_=<phút>` tránh CDN, timeout 5 s, CORS `*`), rồi `webapp.txt` cùng thư mục (chỉ khi http/https), ghi `APP.webappNguon`. Chạy cả khi file:// và trong APK.
- Chỉ nhận URL khớp `WEBAPP_RE` (script.google.com/…/exec). sw.js bỏ qua raw.githubusercontent.com (không cache). Test smoke75 (giả lập GitHub, máy chủ web cục bộ).

## 4m29. v3.52 – Quá trình công tác từ lý lịch tải lên
- Core `ocrQT(text, toanBo)`: tìm tiêu đề "quá trình công tác / quá trình làm việc / kinh nghiệm làm việc" (hoặc cả bài khi là Sơ yếu lý lịch), mỗi dòng bắt đầu bằng khoảng ngày (`QT_RE`: dd/mm/yyyy, mm/yyyy, yyyy, "tháng m năm y", "nay") → `{tuNgay, denNgay, noiDung, donVi, chucVu}`; dòng không có ngày (≤3) nối vào dòng trên; dừng ở "khen thưởng, kỷ luật, đặc điểm lịch sử, quan hệ gia đình, tự nhận xét, cam đoan…". `ocrTach` trả thêm `qt`. Tối đa 80 dòng.
- `tl.ocr` nhận Word (.docx) khi kèm `text` (máy khách tách bằng `docxVanBan`: đọc ZIP + DecompressionStream('deflate-raw'), hàng bảng nối " | "), đánh `ghiChu` cho dòng đã có / tuNgay ≥ ngayVao. `tl.ocrluu` nhận `qt[]` → QuaTrinh loại `Theo lý lịch` (kiểm ngày, bỏ trùng), trả `qt, qtMoi, qtBo`.
- `LOAI_QT` thêm 'Theo lý lịch'; `QT_PHU` = Đào tạo, Khác, Theo lý lịch – không khép quá trình đang mở, không đổi vị trí hiện tại (qt.save, client). Test t68, smoke76 (mẫu test/mau_lylich.docx).

## 4m30. v3.53 – Google Maps JavaScript API
- `cfg.gps.ggKey` (gps.caidat kiểm `^[A-Za-z0-9_-]{30,60}$`, nhật ký không ghi khóa). Client `napGG(key)` nạp `maps/api/js?key&language=vi&region=VN&callback=__ggOk`; `gm_authFailure` → `GM_LOI`, xóa bản đồ Google, render lại bằng Leaflet.
- Lớp tương thích: `GWMap` (setView, panTo, fitBounds{padding,maxZoom}, on('click'), attributionControl), `GWLop` (addTo, remove, getLatLng, bindPopup → InfoWindow, bindTooltip, openPopup, on('dragend')), `GL.circle/circleMarker/marker/polyline/divIcon`; `lx(m)` chọn GL hoặc L. Bấm vòng tròn chuyển tiếp thành bấm bản đồ.
- Tab Điểm chấm công: `gqBamThem(lat,lng)` → chọn bộ phận → `gqDiemModal` với điểm mới đang chọn. sw.js bỏ qua maps.googleapis.com / maps.gstatic.com (không cache). Test smoke77 (giả lập API: test/fake_gmaps.js).

## 4m31. v3.54 – Đề xuất lương duyệt nhiều cấp
- `QT_LL_MD` [tbp, tchc, gd]; `cfg.quyTrinhLenLuong` (quytrinh.save kiểm tên, người). LenLuong thêm `luongDuyet` (JSON cấp, ảnh chụp lúc lập), `capDuyet`, `lichSu`.
- `capDuyetLL(o,cfg)`: quyTbp theo người được đề xuất; bỏ cấp có người duyệt = người lập / người được đề xuất (cấp cuối luôn giữ); 2 cấp liền nhau cùng người giữ cấp sau. `chuanHoaLL`: bản ghi cũ chờ duyệt → đặt ở cấp cuối.
- `ll.save`: quyền propose, hoặc tự đề xuất (maNV = mình), hoặc TP/CHT của bộ phận (`llLaQL` qua quanLyCua); không propose bắt buộc lý do, không ghi số QĐ. `ll.decide`: người của cấp hiện tại hoặc Giám đốc/Quản trị (vượt cấp → Đã duyệt); từ chối cần lý do; `bao()` người cấp sau / người lập. `ll.delete`: người lập khi chưa có cấp duyệt.
- load (không view_all): ll gồm của mình đã duyệt + tự lập + được giao + mình lập. Client: `llCaps/llCap/llCanDecide/llTT/llTienTrinh/llNVs`, view `myll` (vMyLL), badge, Quy trình mục 2b. Test t69, smoke78.

## 4m32. v3.55 – Trợ lý quy định: bộ đệm, dự phòng, tự kiểm
- gas `env.dem` (CacheService 'd_'+khóa, ≤ 95.000 ký tự, ≤ 6 giờ). Core: `botDemGhi/botDemDoc` (chữ văn bản chia khúc 90.000, khóa `bvt_<id>_<textLuc>`), `botPhienBan()` (băm id/capNhat/trạng thái quy định + văn bản + cfg.ai), đệm câu trả lời `bh_<hash(câu hỏi bỏ dấu|phiên bản)>` khi daTraLoi.
- `bot.hoi`: văn bản tối đa 30.000 ký tự, maxTok 1500; aiGoi lỗi → `traCuuTuKhoa(lý do)` + `loiAI`; trả `ms`, `daDem`. `bot.kiem` (READ_ONLY, quyền botqd): thời gian từng bước.
- Client `botGui`: đồng hồ chờ `#botCho`, lỗi mạng → hỏi lại 1 lần sau 2,5 s, thông báo rõ. `botKiem()` ở màn Quản lý. Test t70, smoke79.

## 4m33. v3.56 – Tải ảnh từ điện thoại
- APK `MainActivity.onActivityResult` (.registers 10): kết quả hủy → null; camera (req 2) → Uri ảnh; chọn tệp → `Intent.getClipData()` (nhiều tệp, duyệt getItemAt/getUri) rồi mới `getData()`. Trước đây dùng `FileChooserParams.parseResult` chỉ đọc getData() → chọn nhiều ảnh trả null.
- Client `anhMo(file)`: createImageBitmap(imageOrientation) → createImageBitmap → <img>+canvas; báo HEIC. Dùng ở btNen (bảng tin), chNen (tin nhắn), catMoAnh (CCCD). `btAddFiles` báo khi files rỗng. Test smoke80.

## 4m34. v3.57 – Trích nội dung mục quy định từ tệp
- Client `bqTaiTep(input)`: docx → `docxVanBan`; pdf → `ocrMay(f,'application/pdf',bao,{trangChu:60,trangOCR:10})`; ảnh → `ocrMay` (Tesseract trên máy); txt → `f.text()`; `vbLamSach` ghép dòng; điền bqNoiDung (Thay/Thêm), bqTieuDe, bqVanBan (Số …), bqHieuLuc ("hiệu lực từ ngày dd/mm/yyyy"); `bqDem` đếm ký tự.
- Core: `BOT_ND_MAX` 40.000, `botSD` (giữ xuống dòng) cho noiDung; `botDoanLQ(t, tuHoi, max)` chọn đoạn (theo dòng trống / Điều, Chương, Khoản…) liên quan nhất, giữ thứ tự gốc; prompt mỗi mục ≤ min(6000, 30000/số mục); tra cứu không AI trích ≤ 2.500 ký tự. Test t71, smoke81 (test/mau_quyche.docx).

## 8. Bảo mật

- Mật khẩu lưu dạng băm SHA-256 kèm chuỗi muối riêng từng tài khoản; không ai đọc được mật khẩu gốc.
- Phiên đăng nhập hết hạn sau 6 giờ không thao tác. Sai mật khẩu 5 lần khóa đăng nhập 10 phút.
- Khóa tài khoản người nghỉ việc tại **Tài khoản** (trạng thái "Khóa"); tài khoản bị khóa lập tức mất phiên.
- Quên mật khẩu admin: mở Apps Script → chạy hàm `resetAdminPassword` → mật khẩu trở về `Tapaco@2026`.
- Mọi thao tác ghi vào sheet `NhatKy`.

## 9. Nâng cấp phiên bản

1. Dán `Code.gs` mới đè lên bản cũ → Lưu → chạy `upgradeSchema` (chỉ thêm cột/sheet còn thiếu, không xóa dữ liệu).
2. **Triển khai → Quản lý triển khai → Sửa (bút chì) → Phiên bản: Mới → Triển khai**. URL giữ nguyên.
3. Thay `index.html` trên hosting; tăng `CACHE` trong `sw.js` (vd `tapaco-hr-v2`) để điện thoại tải bản mới.
22. Từ bản 3.1.x lên 3.2 (kiểm tra chấm công bằng AI): dán Code.gs mới → chạy `upgradeSchema` (tạo sheet **KiemTraCC**, **CanhBaoCC**) → Triển khai phiên bản mới; thay index.html và **sw.js** (CACHE tapaco-hr-v27). Dùng chung khóa AI đã cài cho "Phân tích nhóm (AI)".
23. Từ bản 3.2 lên 3.2.1 (sửa lỗi camera, định vị, bản đồ): chỉ thay index.html và **sw.js** (CACHE tapaco-hr-v28); Code.gs chỉ đổi số phiên bản, không đổi dữ liệu. Thay đổi: CSS bản đồ Leaflet nhúng sẵn trong app (không tải từ CDN); nền bản đồ không tải được thì hiện lưới + thông báo, ghim / vòng / tọa độ vẫn dùng được; dòng "Đang đặt" hiện tọa độ điểm đang ghim; lỗi tìm địa chỉ báo tiếng Việt; camera bị chặn hiện hướng dẫn ngay trong khung chụp (Android / iPhone, camera bận, không có camera) kèm nút Thử lại; nhận biết trình duyệt nhúng của Zalo, Facebook, Instagram, WeChat, LINE, TikTok và khung nhúng chặn quyền (Permissions-Policy) để báo đúng nguyên nhân; bản xem thử trên Claude tự chuyển sang mô phỏng định vị, dùng ảnh mẫu khi camera bị chặn.
24. Từ bản 3.2.x lên 3.3 (kiểm tra vị trí khi nhắn tin): dán Code.gs mới → chạy `upgradeSchema` (tạo sheet **KiemTraCong**) → Triển khai phiên bản mới; thay index.html và **sw.js** (CACHE tapaco-hr-v29). Mặc định **đang bật**, chỉ trong giờ làm việc: nếu chưa phổ biến quy chế, TCHC tắt tạm ở tab Cài đặt ngay sau khi nâng cấp.
25. Từ bản 3.3 lên 3.4 (đánh giá tổng hợp CBCNV): dán Code.gs mới → chạy `upgradeSchema` (tạo sheet **DanhGiaTH**, **DanhGiaCN**) → Triển khai phiên bản mới; thay index.html và **sw.js** (CACHE tapaco-hr-v30). Báo cáo Word/PDF cần đã chạy `caiDatLuuTru`.
21. Bản 3.1.1: thay index.html và **sw.js** (CACHE tapaco-hr-v26), Code.gs (chỉ đổi số phiên bản). Sửa: thông báo chặn quyền vị trí bằng tiếng Việt; bố cục danh sách điểm chấm công và hộp chọn chế độ công tác trên điện thoại.
20. Từ bản 3.0.x lên 3.1 (thông báo đẩy, bản đồ, công tác):
    - Dán Code.gs mới → chạy `upgradeSchema`. Lệnh này tạo sheet **CongTac**, **ThietBiPush** và thêm cột `congTac` cho ChamCongGPS.
    - **Triển khai → Quản lý triển khai → Phiên bản mới.** Lần chạy đầu Apps Script có thể xin thêm quyền gọi dịch vụ bên ngoài; bấm cho phép.
    - Chạy lại `caiDatNhacTuDong` để nhắc hẹn từ trình kích hoạt cũng gửi thông báo đẩy.
    - Thay index.html và **sw.js** (CACHE tapaco-hr-v25). sw.js mới có phần nhận thông báo đẩy, **bắt buộc thay**.
    - Quản trị cấu hình Firebase (mục 4s).
19. Bản 3.0.1 (giao diện điện thoại): thay index.html và **sw.js** (CACHE tapaco-hr-v24). Không đổi Code.gs và cấu trúc sheet.
    - Bảng tin không còn bị cắt mép phải. Lịch nghỉ lễ hiển thị 2 tháng mỗi hàng, số ngày to hơn. Danh sách ngày nghỉ chuyển thành thẻ gọn.
    - Ô nhập liệu dùng chữ 16 px để iPhone không tự phóng to khi chạm. Nhãn chữ nhỏ được nâng tối thiểu 12 px.
18. Từ bản 2.9 lên 3.0 (chấm công GPS + khuôn mặt):
    - Dán Code.gs mới → chạy `upgradeSchema` (tạo sheet **ChamCongGPS**, **KhuonMat**) → **Triển khai → Quản lý triển khai → Phiên bản mới**.
    - Thay index.html và **sw.js** (CACHE tapaco-hr-v23).
    - Kiểm tra app đang mở qua https.
    - Khai báo điểm chấm công cho từng công trình, rồi hướng dẫn CBCNV đăng ký khuôn mặt.
17. Từ bản 2.8 lên 2.9:
    - Dán Code.gs mới → **Triển khai → Quản lý triển khai → Phiên bản mới**. Không đổi cấu trúc sheet.
    - Chạy lại 1 lần `caiDatLuuTru` để Apps Script xin quyền mới (xuất PDF).
    - Thay index.html và sw.js (CACHE tapaco-hr-v22).
    - Báo cáo lập bằng bản 2.8 vẫn mở được (thiếu các phần mới).
16. Từ bản 2.7 lên 2.8:
    - Dán Code.gs mới → **Triển khai → Quản lý triển khai → Phiên bản mới**.
    - Chạy 1 lần hàm **`caiDatLuuTru`** (chấp nhận các quyền mới: Google Drive, Google Docs, gọi dịch vụ bên ngoài). Hàm này chạy `upgradeSchema` (thêm cột `luuTru` cho HoiThoai), tạo thư mục **“TAPACO – Lưu trữ nhóm chat (KHÔNG CHIA SẺ)”** cạnh file Google Sheet, cài lại trình kích hoạt 5 phút (nhắc hẹn + lưu trữ mỗi đêm) và ghi vào nhật ký thực thi đường dẫn thư mục cùng cảnh báo nếu thư mục đang bị chia sẻ.
    - Quản trị vào **Phân tích nhóm (AI) → Cài đặt AI** nhập khóa API.
    - Thay index.html và **sw.js** (CACHE tapaco-hr-v21).
15. Từ bản 2.6 lên 2.7:
    - Dán Code.gs mới, chạy `upgradeSchema` (thêm sheet **ChatMuc**, **ChatTraLoi**).
    - Chạy 1 lần `caiDatNhacTuDong` để bật nhắc hẹn và báo cáo định kỳ.
    - Thay index.html và **sw.js** (CACHE tapaco-hr-v19, thêm xử lý bấm vào thông báo).
14. Từ bản 2.5.x lên 2.6: thay Code.gs, index.html (CACHE tapaco-hr-v19); không đổi cấu trúc sheet. Mã NV cũ giữ nguyên; mã mới tự sinh tiếp theo số lớn nhất (VD NV009 → NV0010). Tài khoản cũ giữ tên đăng nhập.
13. Bản 2.5.1: thêm file mẫu Excel tạm ứng và nhập từ Excel; tăng CACHE (tapaco-hr-v17); không đổi cấu trúc sheet.
12. Từ bản 2.4 lên 2.5: chạy `upgradeSchema` (BangLuong không thêm cột mới ngoài bản 2.4); tăng CACHE (đã đổi sẵn thành tapaco-hr-v16). Phiếu lương chi tiết đầy đủ áp dụng cho các kỳ tính lại từ bản này (kỳ cũ vẫn xem được, thiếu một số dòng chi tiết). Thư viện Excel đổi sang xlsx-js-style (có định dạng, màu, viền).
11. Từ bản 2.3 lên 2.4: chạy `upgradeSchema` (thêm cột chiTietNgay, chiTietTamUng cho ChamCong; ctTamUng, ctViPham, ctCong cho BangLuong); tăng CACHE (đã đổi sẵn thành tapaco-hr-v15). Danh sách ngày lễ cũ trong Cài đặt tự chuyển sang Lịch nghỉ lễ. Sau đó vào **Lịch nghỉ lễ** → Nạp lịch theo luật năm 2026, 2027; vào **Cài đặt → Danh mục vi phạm** kiểm tra nhóm của từng mục.
10. Từ bản 2.2 lên 2.3: chạy `upgradeSchema` (thêm cột `boPhan` cho Users); tăng CACHE (đã đổi sẵn thành tapaco-hr-v14). Tài khoản cũ tự đổi vai trò: Nhân sự → Tổ chức – Hành chính, Kế toán → Kế toán tiền lương, Công nợ → Kế toán công nợ; người xử lý `r:hr`, `r:ketoan`, `r:congno` trong Quy trình cũng tự chuyển. Phòng ban cũ trong Cài đặt tự chuyển sang danh mục Phòng ban & công trình. Sau khi nâng cấp, Quản trị: (1) giữ đúng 1 tài khoản Quản trị, gán chức danh cho các tài khoản khác; (2) gán Trưởng phòng / CHT kèm bộ phận; (3) thêm các công trình và điều động người; (4) xem lại **Quy trình & phê duyệt**. Quy trình lương mặc định mới có 6 bước như mục 5; ai đang dùng quy trình đã lưu từ bản cũ thì giữ nguyên, muốn đổi sang quy trình mới bấm **Khôi phục mặc định** ở đó.
9. Từ bản 2.1 lên 2.2: chạy `upgradeSchema` (thêm cột noiBat, noiBatDen cho BaiViet); tăng CACHE trong sw.js (đã đổi sẵn thành tapaco-hr-v13).
8. Từ bản 2.0 lên 2.1: chạy `upgradeSchema` (thêm cột thuMuc cho HoiThoai; msgId, seq, w, h, thuMuc cho TepChat). Sau đó Quản trị vào Tin nhắn → Kho tệp → **Xếp tệp cũ vào thư mục theo nhóm**.
7. Từ bản 1.9 lên 2.0: chạy `upgradeSchema` để tạo sheet **HoiThoai**, **ThanhVien**, **TinNhan**, **TepChat**.
6. Từ bản 1.8 lên 1.9: chạy `upgradeSchema` để tạo sheet **BaiViet** và **TuongTac**.
5. Từ bản 1.7 lên 1.8: chạy `upgradeSchema` để tạo sheet **MucHuong** và cột `anCa` trong NhanVien (bắt buộc, nếu không app báo thiếu sheet). Mẫu công thức TAPACO có thêm khoản "Phụ cấp trách nhiệm" (PCTN); ai đã lưu mẫu 1.7 thì nạp lại mẫu hoặc tự thêm khoản `PC_TRACHNHIEM * HS_HUONG` và cộng PCTN vào công thức THUONGNS.
4. Từ bản 1.6 lên 1.7: bộ công thức đã lưu được giữ nguyên. Muốn dùng mẫu TAPACO: Công thức lương → Nạp mẫu → Bảng lương TAPACO 2026 → Chạy thử → Lưu (tự thêm khoản "Tiền chuyên cần" vào danh mục). Sau đó khai Nơi làm việc và Lương cơ bản đóng BH cho từng hồ sơ (có thể nhập bằng Excel).

## 10. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| "Không kết nối được máy chủ" | Kiểm tra URL kết thúc bằng `/exec`, quyền truy cập "Bất kỳ ai", đã triển khai phiên bản mới sau khi sửa code |
| "Thiếu sheet …" | Chạy hàm `setup` hoặc `upgradeSchema` |
| "Hệ thống đang bận" | Nhiều người lưu cùng lúc; thử lại sau vài giây |
| CCCD/SĐT mất số 0 đầu khi sửa tay trong Sheet | Không sửa tay trong Sheet; nếu cần, định dạng cột là Văn bản thuần (setup đã đặt sẵn) |
| CBCNV không thấy phiếu lương | Kỳ lương chưa được Giám đốc phê duyệt |
| Không bấm được nút chuyển bước | Chỉ người phụ trách bước hiện tại thấy nút; xem dòng "Bước hiện tại" |
| Người làm lâu năm hiện "quá hạn nâng bậc" | Nhập "Ngày hưởng bậc hiện tại" trong hồ sơ |
