# MASTER PROMPT · TẠO LẠI APP "TAPACO NHÂN SỰ – TIỀN LƯƠNG" (v3.26)

> Dán toàn bộ nội dung dưới đây cho Claude để dựng lại hệ thống từ đầu, hoặc làm tài liệu đặc tả khi nâng cấp.

---

## 0. Vai trò & mục tiêu

Bạn là kỹ sư phần mềm xây dựng công cụ nội bộ cho **Công ty Cổ phần TAPACO Việt Nam** (xây dựng hạ tầng, Hà Nội). Hãy tạo ứng dụng quản lý nhân sự – tiền lương **chạy thật được**, không phải bản demo:

1. Theo dõi **hồ sơ** và **quá trình công tác** của cán bộ công nhân viên (CBCNV).
2. Theo dõi **lên lương** (đề xuất → duyệt → tự áp dụng theo ngày hiệu lực), cảnh báo **đến hạn nâng bậc**, **HĐLĐ sắp hết hạn**.
3. Theo dõi **khen thưởng – kỷ luật**; thưởng đã duyệt gắn kỳ lương tự vào bảng lương.
4. **Tính lương hằng tháng** đúng quy định Việt Nam 2026 (BHXH/BHYT/BHTN có trần, thuế TNCN biểu 5 bậc) theo **quy trình 6 bước**: TCHC chấm công → Kế toán làm lương → Kế toán công nợ → TCHC kiểm tra, khấu trừ vi phạm theo quy chế → Giám đốc phê duyệt → lưu kết quả, khóa dữ liệu, xuất in; CBCNV chỉ xem lương sau khi duyệt.
5. **Tài khoản riêng cho từng CBCNV**: mỗi người chỉ xem hồ sơ, quá trình, lịch sử lương, phiếu lương của mình, và **gửi ý kiến / đề nghị xem xét lại ngay trên phiếu lương**, được chuyển tới đúng bộ phận để trả lời.
6. **Xin nghỉ phép năm / việc riêng có lương / nghỉ ốm (BHXH chi trả) / nghỉ không lương**: CBCNV tạo đơn, TCHC duyệt (đơn > 3 ngày thêm bước Giám đốc), quỹ phép năm tự tính, đơn đã duyệt đổ vào chấm công.
9. **Công thức lương do Quản trị lập** (lương cơ bản, phép, lễ Tết, phụ cấp, ăn ca, tăng ca, làm đêm, KPI…) bằng biểu thức kiểu Excel; BH và thuế giữ theo luật.
8. **Tài liệu đính kèm hồ sơ** (HĐLĐ, CCCD, MST, quyết định…) lưu Google Drive riêng tư, xem qua app theo quyền, có nhật ký xem.
7. **Quy trình & người phê duyệt cài đặt được** bởi Quản trị và Giám đốc: các bước lương, các cấp duyệt nghỉ, người tiếp nhận ý kiến theo chủ đề, trưởng bộ phận / công trường.

Toàn bộ giao diện, thông báo, tài liệu bằng **tiếng Việt**.

## 1. Kiến trúc (bắt buộc giữ)

- `index.html`: PWA **một file**, vanilla JS + thư viện CDN (chỉ `xlsx 0.18.5` từ cdnjs). Font Barlow / Barlow Condensed (Google Fonts). Không build tool.
- `Code.gs`: Google Apps Script; **Google Sheets là CSDL**; `doPost` nhận JSON `{action, token, ...}` gửi với `Content-Type: text/plain` (tránh CORS preflight); trả `{ok, data}` hoặc `{ok:false, error, auth}`.
- **Lõi nghiệp vụ dùng chung** `HRCORE` (schema, cấu hình mặc định, động cơ tính lương, phân quyền, mọi handler API) viết một lần, được nhúng nguyên văn vào **cả** `Code.gs` và `index.html`. Lõi nhận `env` gồm: `db.table(name)` (mảng object, sửa tại chỗ), `db.dirty(name)`, `hash`, `uuid`, `now`, `today`, `sessions.get/put/del`, `attempts.get/put`, `log`, `readLog`.
  - Trên Apps Script: `SheetDB_` đọc cả sheet → object (Date → `yyyy-MM-dd`); v3.14 ghi theo dòng (giữ `goc` = đối tượng / ô / số dòng lúc đọc; dòng sửa → setValues gộp liền nhau, dòng xóa → deleteRows từ dưới lên, dòng mới → nối cuối; chèn giữa / đảo thứ tự / xóa > 30 dòng → ghi cả bảng); `db.tail(name, truong, tu)` đọc khúc 1.500 dòng từ cuối tới khi gặp truong < tu, đọc cả bảng sau đó thì ghép phần đầu vào cùng mảng; core `gpsTu(tu)` dùng cho gps.cham / cuatoi / ds / duyet (req.tu) / tonghop / ccktQuet / ctCapNhatLuot; `load` thuộc READ_ONLY, thao tác chỉ đọc mà phát sinh ghi thì lấy khóa rồi `flush(true)` kiểm cột id trước khi ghi; getCfg nhớ trong 1 yêu cầu; kiểm thử `test/t38_gas.js`, đo `test/perf/perf_gas.js`; `CacheService` giữ phiên 6 giờ (gia hạn mỗi lần gọi); `LockService` cho mọi thao tác ghi; `NhatKy` dùng `appendRow`.
  - Trong trình duyệt khi **chưa cấu hình URL máy chủ**: chế độ dùng thử chạy cùng lõi trên `localStorage` (SHA-256 đồng bộ viết tay), có dữ liệu mẫu và nút đăng nhập nhanh 5 vai trò, mật khẩu `123456`.
- Quy trình build: `src/core.js` + `src/gas.js` + `src/client.html` (placeholder `/*__CORE__*/`) → `dist/index.html`, `dist/Code.gs`; thêm `manifest.webmanifest`, `sw.js` (chỉ cache giao diện, bỏ qua script.google.com), `icon-192/512.png`.
- `upgradeSchema()` chỉ **thêm** sheet/cột thiếu, không xóa dữ liệu; cột chữ đặt định dạng `@` (giữ số 0 đầu CCCD, SĐT, STK; ngày dạng `yyyy-MM-dd`), cột tiền `#,##0`. `setup()` tạo sheet, cấu hình mặc định, tài khoản `admin / Tapaco@2026` (bắt buộc đổi mật khẩu). `resetAdminPassword()` khôi phục admin.
- Thương hiệu: navy `#0F1E3D`/`#16294F`, cam `#F97316`; giao diện **sáng**, bảng dày đặc, chữ to rõ cho dùng ngoài công trường; sidebar trên desktop, ngăn kéo trên mobile; logo vẽ bằng canvas; manifest tạo lúc chạy nếu không có file ngoài.
- PDF/in: `window.print()` với vùng `#print` (không dùng jsPDF, để tiếng Việt chuẩn).

## 2. Lược đồ dữ liệu (sheet → cột)

- **ChamCong** (v2.4) thêm: chiTietNgay (JSON {ngày: ký hiệu}), chiTietTamUng (JSON). **BangLuong** thêm: ctTamUng, ctViPham, ctCong.
- **Users**: username, passHash, salt, role, maNV, hoTen, active (Có/Không), mustChange, lastLogin, createdAt, boPhan (v2.3)
- **NhanVien**: maNV, hoTen, gioiTinh, ngaySinh, cccd, ngayCapCCCD, noiCapCCCD, dienThoai, email, diaChi, queQuan, trinhDo, chuyenNganh, phongBan, chucVu, donVi, ngayVao, loaiHD, soHD, ngayKyHD, ngayHetHanHD, trangThai, ngayNghi, ngach, bac, **ngayHuongBac**, luongChinh, luongBH, phuCap (JSON `{maPC: số tiền}`), thamGiaBH, doanVien, mst, soNPT, camKet08, nganHang, soTK, ghiChu, capNhat, loaiLuong, lyDoKhongBH, **congTy** (mã công ty thành viên, v3.20)
- **QuaTrinh**: id, maNV, loai, tuNgay, denNgay, phongBan, chucVu, donVi, soQD, nguoiKy, noiDung, taoBoi, taoLuc
- **LenLuong**: id, maNV, loai, hieuLuc, ngachCu, bacCu, luongCu, luongBHCu, ngachMoi, bacMoi, luongMoi, luongBHMoi, lyDo, soQD, trangThai (Chờ duyệt/Đã duyệt/Từ chối), apDung (Có khi đã áp vào hồ sơ), nguoiDeXuat, ngayDeXuat, nguoiDuyet, ngayDuyet, ghiChuDuyet
- **KhenThuong**: id, maNV, loai, hinhThuc, ngay, soTien, chiuThue, kyLuong (YYYY-MM), noiDung, soQD, trangThai, nguoiDeXuat, ngayDeXuat, nguoiDuyet, ngayDuyet
- **ChamCong**: id (`ky|maNV`), ky, maNV, ngayCong, nghiHuongLuong, nghiKhongLuong, ot150, ot200, ot300, tamUng, khauTruKhac (khấu trừ công nợ), truyLinh, ghiChu, ghiChuCongNo, chiTietPhat (JSON `[{ma, ten, loai, soLan, mucTien, thanhTien, ghiChu}]`), ghiChuTCHC
- **BangLuong**: id, ky, maNV, hoTen, phongBan, chucVu, luongChinh, ngayCongChuan, ngayCongHL, luongThoiGian, phuCap, phuCapChiTiet (JSON), tienOT, thuong, truyLinh, viPham, giamThuong, tongThuNhap, luongDongBH, bhxh, bhyt, bhtn, doanPhi, tnChiuThue, giamTru, tnTinhThue, thueTNCN, boiThuong, tamUng, khauTruKhac, thucLinh, bhxhDN, bhytDN, bhtnDN, kpcdDN, ghiChu, ctTamUng, ctViPham, ctCong, **congTy** (công ty hạch toán chốt lúc tính lương, v3.20)
- **KyLuong**: ky, trangThai (Chấm công / Tính lương / Công nợ / TCHC kiểm tra / Chờ duyệt / Đã duyệt), ngayCongChuan, nguoiTinh, ngayTinh, nguoiDuyet, ngayDuyet, soNguoi, tongThuNhap, tongThucLinh, maKiemTra, capNhat (bản cũ 'Nháp' → 'Chấm công', 'Đã chốt' → 'Chờ duyệt')
- **NghiPhep**: id, maNV, hoTen, loai, tuNgay, denNgay, buoi (Cả ngày/Sáng/Chiều — chỉ khi 1 ngày → 0,5), soNgay, lyDo, chungTu, trangThai (Chờ duyệt / Chờ GĐ duyệt / Đã duyệt / Từ chối / Đã hủy), nguoiTao, taoLuc, nguoiDuyet1, ngayDuyet1, nguoiDuyet, ngayDuyet, ghiChuDuyet. NhanVien thêm `ngayPhepNam` (ghi đè quỹ phép). ChamCong thêm `nghiOm`.
- **TaiLieu**: id, maNV, loai, ten, soHieu, ngay, hetHan, lienKet ('qt:id' | 'll:id' | 'kt:id'), nvXem (Có/Không), trangThai (Đã xác nhận / Chờ xác nhận), tenFile, mimeType, kichThuoc, fileId (Drive), nguoiTai, taiLuc, nguoiXacNhan, ghiChu. Cấu hình: loaiTaiLieu, taiLieuBatBuoc (mặc định HĐLĐ + CCCD), loaiNVTuTai, dungLuongToiDaMB 10.
- **YKien**: id, ky, maNV, hoTen, chuDe, boPhan (role), trangThai (Mới / Đang xử lý / Đã giải quyết / Không chấp nhận), noiDung, thread (JSON `[{tu, ten, vaiTro, luc, noiDung, ghiChu}]`), taoLuc, capNhat, nguoiXuLy. Chủ đề → bộ phận: Ngày công, nghỉ, làm thêm giờ → hr; Lương, phụ cấp, thưởng → ketoan; Thuế TNCN, bảo hiểm → ketoan; Tạm ứng, công nợ → congno; Vi phạm, khấu trừ theo quy chế → hr; Khác → hr.
- **PheDuyet**: id, ky, tu, den, hanhDong (Gửi/Chuyển/Trình duyệt/Phê duyệt/Trả lại/Mở khóa), nguoi, vaiTro, thoiGian, ghiChu
- **CauHinh**: key, value (JSON)
- **NhatKy**: thoiGian, user, action, chiTiet

Danh mục: loại HĐ = Không xác định thời hạn / Xác định thời hạn / Thử việc / HĐ dưới 3 tháng / Thời vụ – Khoán việc. Trạng thái NV = Đang làm việc / Thử việc / Tạm nghỉ / Đã nghỉ việc. Loại quá trình = Tiếp nhận, Điều động, Bổ nhiệm, Miễn nhiệm, Thay đổi HĐLĐ, Đào tạo, Tạm nghỉ, Nghỉ việc, Khác. Loại lên lương = Nâng bậc định kỳ, Nâng lương trước hạn, Điều chỉnh lương, Giảm lương. Loại KT = Thưởng tháng, Thưởng lễ – Tết, Thưởng công trình / dự án, Lương tháng 13, Thưởng đột xuất, Kỷ luật (hình thức: Khiển trách, Kéo dài thời hạn nâng lương ≤ 6 tháng, Cách chức, Sa thải — **không có tiền phạt**, Điều 127 BLLĐ 2019).

## 3. Cấu hình mặc định (sửa được trong Cài đặt)

- Lương cơ sở 2.530.000 (NĐ 161/2026, từ 01/7/2026); lương tối thiểu vùng I 5.310.000 (NĐ 293/2025, từ 01/01/2026).
- Trần BHXH/BHYT = 20 × lương cơ sở; trần BHTN = 20 × lương tối thiểu vùng.
- NLĐ: BHXH 8, BHYT 1,5, BHTN 1 (%). DN: BHXH 17,5 (gồm TNLĐ-BNN), BHYT 3, BHTN 1, KPCĐ 2. Đoàn phí 1% (tối đa 10% lương cơ sở).
- Giảm trừ bản thân 15.500.000; người phụ thuộc 6.200.000 (kỳ tính thuế 2026).
- Biểu thuế lũy tiến (Luật 109/2025/QH15): [10tr 5%], [30tr 10%], [60tr 20%], [100tr 30%], [trên 35%].
- HĐ thử việc/dưới 3 tháng/thời vụ và lương khoán việc: khấu trừ 10% nếu thu nhập ≥ 5.000.000 đ/lần (NĐ 253/2026 Điều 50), trừ khi có cam kết 08.
- Loại lương theo hợp đồng (`NhanVien.loaiLuong`): `bh` trả lương + đóng BH (mặc định); `khongbh` không đóng BH tại Công ty, `lyDoKhongBH` = "Đóng BH nơi khác" (DN đóng TNLĐ-BNN `tyLeDN.tnld` 0,5% trên lương đóng BH, ghi vào bhxhDN) | "Đã nghỉ hưu" (không BH hai bên); `khoan` lương khoán việc: thu nhập = `ChamCong.tienKhoan` (nhập bước Tính lương; trống → lương chính = mức khoán tháng) + thưởng + truy lĩnh, không chạy công thức lương, không BH, khấu trừ 10% như vãng lai. Hồ sơ cũ: thamGiaBH "Không" → khongbh, còn lại → bh; nv.save đồng bộ thamGiaBH, bắt buộc lý do khi khongbh.
- Công chuẩn: cố định 26 hoặc "số ngày trong tháng trừ Chủ nhật"; 8 giờ/ngày; OT 150/200/300%.
- Nghỉ không lương ≥ 14 ngày/tháng → không đóng BH. Chu kỳ nâng bậc 24 tháng, cảnh báo trước 45 ngày.
- `ngayPhepCoBan` 12, `thamNienCongPhep` 5, `nghiCanGDDuyet` 3, `ngayLe` [] (yyyy-mm-dd, cùng Chủ nhật không tính vào đơn nghỉ).
- `tyLeBoiThuongToiDa` 30; `danhMucViPham` `[{ma, ten, mucTien, loai}]` để trống (Công ty nhập theo Nội quy).
- Phụ cấp mặc định (mã, đóng BH, theo công, miễn thuế): chucvu (có, không, 0), trachnhiem (có, không, 0), congtruong (không, có, 0), anca (không, có, 730.000), dienthoai (không, không, 0), xangxe (không, không, 0). `mienThue`: 0 = chịu thuế toàn bộ, >0 = miễn tối đa, −1 = miễn toàn bộ.
- Danh mục phòng ban, chức vụ, đơn vị/công trường; thang bảng lương `[{ngach, bac, muc}]` (để trống, Công ty tự nhập — không bịa số).

## 4. Động cơ tính lương `tinhPhieuLuong(nv, cc, cfg, {luong, luongBH, ncc, thuong, thuongChiuThue})`

1. luongTG = round(L/C × (ngayCong + nghiHuongLuong)).
2. Phụ cấp: theoCong → mức × min(ngayCong, C)/C; khác → mức × min(công HL, C)/C. Cộng phần đóng BH theo **mức gốc**; phần chịu thuế trừ `mienThue`.
3. tienOT = lương giờ × (h150×1,5 + h200×2 + h300×3), lương giờ = L/C/8; phần chịu thuế = lương giờ × tổng giờ (phần cao hơn được miễn).
4. Tổng = luongTG + PC + OT + thưởng + truy lĩnh.
5. Lương đóng BH = luongBH (nếu > 0) hoặc L + PC đóng BH; = 0 nếu thamGiaBH = Không hoặc (nghỉ KL + nghỉ ốm) ≥ ngưỡng. Ngày nghỉ ốm không tính vào công hưởng lương. Áp trần riêng BHXH-BHYT và BHTN.
6. TNCT = luongTG + PC chịu thuế + OT chịu thuế + thưởng chịu thuế + truy lĩnh.
7. Thuế: vãng lai → 10% (theo ngưỡng/cam kết); còn lại → lũy tiến trên max(0, TNCT − BH NLĐ − giảm trừ).
8. Vi phạm (chiTietPhat): "Giảm thưởng / phụ cấp theo quy chế" → giamThuong = min(tổng đề nghị, thưởng + PC tháng), trừ vào Tổng và TNCT (tối đa thưởng chịu thuế + PC chịu thuế); "Bồi thường thiệt hại" → boiThuong = min(đề nghị, 30% × (Tổng − BH NLĐ − thuế)). Phần vượt ghi vào ghiChu. Không có hình thức "phạt tiền" (Điều 127 BLLĐ 2019).
9. Thực lĩnh = Tổng − BH NLĐ − đoàn phí − thuế − bồi thường − tạm ứng − khấu trừ công nợ. Chi phí DN tính trên mức đã áp trần.
- **Lương tại kỳ** `luongTaiKy`: quyết định đã duyệt mới nhất có hiệu lực ≤ ngày cuối kỳ; nếu không có, lấy `luongCu` của quyết định sớm nhất sau kỳ; nếu không có, lấy hồ sơ.
- **Hạn nâng bậc**: mốc = max(ngayHuongBac hoặc ngayVao, hiệu lực lần nâng bậc/nâng trước hạn đã duyệt gần nhất) + chu kỳ; +6 tháng cho mỗi kỷ luật "Kéo dài thời hạn nâng lương" sau mốc. Không áp cho thử việc/HĐ vãng lai.
- Kiểm thử chuẩn: L 20.000.000, PC chức vụ 2.000.000, ăn ca 730.000, đủ 26/26, 0 NPT → tổng 22.730.000; BH 2.310.000; TN tính thuế 4.190.000; thuế 209.500; thực lĩnh 20.210.500.

## 5. Phân quyền (kiểm tra tại máy chủ, giao diện chỉ ẩn/hiện) — v2.3

ROLES (một chức danh / tài khoản): admin (chỉ 1 tài khoản), giamdoc, tchc, ktluong, ktcongno, truongphong, cht, nhanvien. ROLE_CU đổi tên cũ: hr→tchc, ketoan→ktluong, congno→ktcongno; normRole() áp dụng khi đăng nhập, publicUser, token `r:` trong quy trình. ROLE_QL = [truongphong, cht].

| Quyền | Vai trò |
|---|---|
| view_all | admin, giamdoc, tchc, ktluong, ktcongno |
| edit_nv (hồ sơ, quá trình, phòng ban – công trình, điều động) | admin, tchc |
| propose (lên lương, KT-KL) | admin, tchc |
| approve (duyệt lên lương, KT-KL) | admin, giamdoc |
| payroll / payroll_approve | admin, ktluong / admin, giamdoc |
| users | admin (mọi chức danh), tchc (chỉ nhanvien) |
| config, congthuc | admin |
| quytrinh, log | admin, giamdoc |
| tailieu | admin, giamdoc, tchc, ktluong |
| bantin | admin, giamdoc, tchc |

- Quy trình lương mặc định (QT_LUONG_MD): Chấm công (r:ktcongno) → Công nợ (r:ktcongno) → TCHC duyệt công / vipham (r:tchc) → Tính lương (r:ktluong) → TCHC duyệt lương / kiemtra (r:tchc) → Chờ duyệt / pheduyet (r:giamdoc). kiemTraQTLuong: bước 1 chamcong, có tinhluong, cuối pheduyet.
- **Giám đốc phê duyệt ngay**: ky.move to=Đã duyệt từ bất kỳ bước nào khi role giamdoc/admin → recalc, khóa kỳ, PheDuyet ghi "bỏ qua bước: …".
- Đơn nghỉ mặc định (QT_NGHI_MD): 'tbp' (TP/CHT có Users.boPhan = bộ phận người xin; người xin không tự duyệt; dự phòng cfg.truongBoPhan rồi r:tchc) → r:tchc. Giám đốc duyệt vượt cấp.
- Ý kiến (QT_YKIEN_MD): Lương → r:tchc + r:ktluong; Thuế → ktluong; Tạm ứng → ktcongno; còn lại → tchc.
- upsertUser: TCHC chỉ tạo/sửa nhanvien; nhanvien/TP/CHT bắt buộc maNV; chỉ 1 admin; boPhan phải có trong danh mục; TP/CHT phải có bộ phận. tenDangNhap(hoTen): tên + chữ đầu họ đệm, không dấu ("Lê Văn Thắng" → thanglv), trùng thêm số, tối thiểu 3 ký tự; user.bulk dùng cách này.
- loadAll cho TP/CHT: nhánh không view_all + bpNV (CBCNV bộ phận, trường cơ bản), bpQT, np của người trong bộ phận; qlBP = {bộ phận: [TP/CHT]} gửi cho mọi người.

Vai trò `nhanvien` chỉ nhận: hồ sơ mình, quá trình mình, lên lương & KT-KL **đã duyệt** của mình, bảng lương của mình ở kỳ **đã duyệt**, cấu hình rút gọn. Không trả `passHash/salt` ra ngoài.

## 5b. Quy trình cấu hình được (v1.4)

- Mã người xử lý: `r:<vai trò>` | `u:<tên đăng nhập>` | `tbp` (trưởng bộ phận của người đề nghị, tra `cfg.truongBoPhan[donVi] || [phongBan]`, không có thì `r:hr`). `laNguoiXuLy(s, nguoi)`: admin luôn đúng.
- `cfg.quyTrinhLuong` `[{id, ten, loai, nguoi}]`, loai ∈ chamcong / tinhluong / congno / vipham / kiemtra / pheduyet. Ràng buộc: bước 1 chamcong, bước 2 tinhluong, bước cuối pheduyet; đúng 1 chamcong, tinhluong, pheduyet; congno, vipham ≤ 1; không dùng tbp; bước nhập liệu không giao cho tài khoản vai trò nhanvien. Trường nhập theo loại: FIELDS_LOAI. KyLuong.trangThai = id bước hoặc 'Đã duyệt'; KyLuong.quyTrinh = ảnh chụp các bước (kèm nguoiTen) lúc mở kỳ. Mặc định id = tên cũ để tương thích. Tính lại tự động khi ở các bước từ Tính lương đến trước Phê duyệt.
- `cfg.quyTrinhNghi` `[{id, ten, nguoi, tren, loai[]}]`: cấp áp dụng khi soNgay > tren và (loai rỗng hoặc chứa loại nghỉ); khi tạo đơn chụp vào `luongDuyet` (đã quy tbp về người cụ thể, bỏ cấp mà người duyệt là người nộp, bỏ cấp trùng người với cấp liền trước), `capDuyet` = chỉ số cấp hiện tại, `lichSu` = [{cap, nguoi, luc, ketQua, ghiChu}]. Giám đốc / admin duyệt vượt cấp → Đã duyệt. Đơn cũ 'Chờ GĐ duyệt' được chuẩn hóa.
- `cfg.quyTrinhYKien` `[{chuDe, nguoi}]`; YKien lưu `xuLy` (mã người) + `boPhan` (tên hiển thị); chuyển xử lý bằng `chuyen` (mã r:/u:).
- `quytrinh.save {quyTrinhLuong, quyTrinhNghi, quyTrinhYKien, truongBoPhan, apDungKyDangXuLy}` – quyền `quytrinh` (admin, giamdoc); trả cfg đã gắn nguoiTen, danh sách kỳ được cập nhật / giữ cũ.
- loadAll: tài khoản CBCNV được nhận thêm đơn nghỉ / ý kiến có giao cho `u:<mình>`; Giám đốc nhận danh sách tài khoản để chọn người xử lý.
- Màn hình "Quy trình & phê duyệt": sơ đồ luồng (thẻ bước → mũi tên), bảng bước có ô tên, chọn loại, chip người xử lý + ô "thêm người xử lý", nút ↑ ↓ ✕, khôi phục mặc định, bảng trưởng bộ phận cho mọi phòng ban / đơn vị.

## 5c. Tài liệu đính kèm (v1.5)

- `env.files` {put({name, mime, data(base64), maNV, hoTen}) → fileId, get(id) → {data}, del(id)}. Apps Script: thư mục gốc "TAPACO – Tài liệu hồ sơ nhân sự (KHÔNG CHIA SẺ)" (id lưu ScriptProperties TL_ROOT), thư mục con "<maNV> – <họ tên>", file giữ quyền riêng tư mặc định, xóa = setTrashed. Bản dùng thử: localStorage 'tphr_f_<id>'.
- Quyền mới `tailieu` (admin, giamdoc, hr, ketoan) xem mọi tài liệu; `edit_nv` tải / sửa / xác nhận / xóa; CBCNV tự tải các loại trong loaiNVTuTai (→ Chờ xác nhận), chỉ xem tài liệu của mình có nvXem ≠ Không, xóa được tài liệu của mình khi còn chờ xác nhận. MIME cho phép: pdf, jpeg/png/webp/heic, doc/docx, xls/xlsx; giới hạn dung lượng theo cấu hình. `tl.get` trả base64 và ghi nhật ký 'tl.xem'. Xóa hồ sơ (admin) thì xóa luôn file.
- Giao diện: tab Tài liệu trong hồ sơ (chip giấy tờ bắt buộc ✓/…/✗, thẻ tài liệu theo loại có biểu tượng PDF/ẢNH/DOC/XLS, số hiệu, ngày, hạn, gắn với, nút Xem / Xác nhận / Sửa / Xóa); ô kéo thả, tự nén ảnh > 600 KB (cạnh dài 2200 px, JPEG 0,85); xem trước ảnh/PDF trong modal bằng blob URL, nút Tải về; 📎 trên dòng thời gian; tab "Rà soát giấy tờ" (3 thẻ số, bảng chờ xác nhận, bảng sắp hết hạn, ma trận NV × giấy tờ bắt buộc); cột "Giấy tờ n/m" trong danh sách; cảnh báo hết hạn ở Tổng quan.

## 5d. Công thức lương (v1.6)

- Bộ tính biểu thức tự viết (không eval): số, biến [A-Za-z_][A-Za-z0-9_]* (quy về IN HOA), + − * / ( ), so sánh > < >= <= = <> (trả 1/0), hàm MIN MAX ROUND(x,n) FLOOR CEIL ABS IF(c,a,b) AND OR NOT; chia 0 = 0; AST được cache; `bienDung` liệt kê biến để kiểm tra.
- cfg.congThucLuong `[{ma, ten, nhom: thunhap|khautru, phanLoai: luong|phucap|tangca|thuong|khac, congThuc, thue: toanbo|khong|mientoida|congthuc, thueGT, giamTruQC, an}]` (mặc định CT_LUONG_MD 15 khoản), cfg.bienKy `[{ma, ten, buoc: chamcong|tinhluong, macDinh}]` (mặc định OT_DEM, KPI_PT), cfg.congThucDongBH. Danh mục phuCap chỉ còn {ma, ten} → biến PC_<MÃ>; thêm mặc định nhao, kpi.
- Biến hệ thống BIEN_HT (LUONG_CHINH, LUONG_BH, CONG_CHUAN, LUONG_NGAY, LUONG_GIO, GIO_CONG_NGAY, NGAY_CONG, NGHI_HL, NGAY_LE, NGHI_KL, NGHI_OM, CONG_HL, OT150/200/300, HS_OT150/200/300, THUONG, THUONG_CHIU_THUE, TRUY_LINH, SO_NPT, THAM_NIEN_THANG, LA_THU_VIEC, LUONG_CO_SO, LUONG_TOI_THIEU_VUNG) + PC_* + biến kỳ + mã các khoản phía trên; trong công thức phần chịu thuế có thêm GT.
- tinhPhieuLuong: tính lần lượt các khoản (làm tròn đồng), cộng thu nhập/phần chịu thuế theo từng khoản, giảm trừ quy chế tối đa bằng tổng khoản giamTruQC, lương đóng BH theo công thức rồi áp trần; BH/thuế/bồi thường/tạm ứng như cũ; khấu trừ công thức trừ sau thuế. Trả thêm chiTiet JSON [{ma, ten, nhom, phanLoai, gt, thue}], ktCongThuc; các trường tổng hợp luongThoiGian/phuCap/tienOT/thuong/truyLinh = tổng theo phân loại. ChamCong thêm ngayLe, bienKy JSON (cc.save chỉ nhận biến của bước hiện tại). KyLuong.congThuc = ảnh chụp bộ công thức lúc tính.
- Hành động `congthuc.save {congThucLuong, bienKy, congThucDongBH, tinhLaiKyMo}` và `congthuc.test {…, maNV, ky}` – quyền `congthuc` (admin); kiemTraCT kiểm mã (IN HOA 2–20 ký tự, không trùng biến), biến chưa có, cú pháp, chỉ tham chiếu khoản phía trên.
- Giao diện "Công thức lương": thẻ từng khoản (số thứ tự, mã, tên, loại, ô công thức đơn cách sau dấu =, chọn thuế, cờ giảm trừ / tạm tắt, giá trị chạy thử bên phải), ↑ ↓ ✕, khôi phục mẫu; công thức lương đóng BH; bảng biến nhập theo kỳ; cột phải dính: chip biến (hệ thống, PC_, biến kỳ, khoản đã lập, hàm) chèn vào ô đang chọn, khối Chạy thử (bảng khoản / thuế / tổng / thực lĩnh, danh sách giá trị biến). Bảng lương, Excel, phiếu lương hiển thị theo từng khoản của kỳ.

## 5e. Mẫu TAPACO, nơi làm việc, giá trị trung gian (v1.7)

- Khoản `nhom: 'trunggian'` {ma, ten, congThuc, vaiTro: ''|congchuan|conghl, an, ghiChu}: tính không làm tròn, gán vào biến, không cộng thu nhập, không vào chiTiet; vaiTro ghi đè BangLuong.ngayCongChuan / ngayCongHL (mỗi vai trò tối đa 1 khoản).
- cfg.noiLamViec `[{ma, ten, congChuan, anCa}]` (mặc định VP 24/30.000, DA 28/70.000); NhanVien.noiLamViec; ChamCong.noiLamViec (bước chấm công, trống = theo hồ sơ, cc.save chặn mã lạ, lưu IN HOA). Biến NLV_<MÃ> (1/0), CONG_CHUAN_NLV, TIEN_AN_NGAY.
- CT_MAU {tapaco, chuan}; mặc định = tapaco: CC, CONG_TL, TL_HD, HS_HUONG, NGAY_NGHI_CL, TONG_TG (trung gian, TONG_TG = ROUND(LUONG_CHINH*HS_HUONG + LUONG_BH*TL_HD/CC*NGAY_NGHI_CL, -2)); LCB, PCNO, PCXX, PCDT, PCCC, LNGHI, THUONGNS (= TONG_TG − tổng các khoản trên), ANCA (miễn tối đa 1.200.000), THUONGKT, TRUYLINH; biến kỳ HS_K (tinhluong, 1), NGAY_AN (chamcong, 0); lương đóng BH = LUONG_BH. Danh mục phụ cấp thêm chuyencan; congthuc.save/test tự thêm PC_ còn thiếu nếu có trong danh mục mặc định.
- luongTaiKy: quyết định không ghi lương BH thì giữ mức BH trước đó. Giao diện: "Nạp mẫu…", thẻ trung gian (viền xám, chọn vai trò, hiện giá trị chạy thử), chip NLV_*, bảng Nơi làm việc trong Cài đặt, cột "Nơi LV" khi chấm công; nhãn OT đổi thành Giờ tăng ca / Giờ làm CN / Giờ làm lễ (nhập Excel vẫn nhận tên cũ).
- Kiểm thử: test/t3.js đối chiếu với các dòng của file Excel (thu nhập, ăn ca, thuế, thực nhận); test/t4.js 14 kiểm tra v1.7.

## 5f. Mức hưởng theo thời gian (v1.8)

- Sheet MucHuong: id, maNV, hieuLuc, giaTri JSON (khóa luongChinh, luongBH, soNPT, anCa, pc.<mã>; chỉ khoản thay đổi), cu JSON (mức ngay trước hiệu lực của các khóa đó), soQD, lyDo, trangThai (Chờ duyệt | Hiệu lực | Từ chối), nguoiTao, taoLuc, nguoiDuyet, ngayDuyet, ghiChuDuyet. NhanVien thêm anCa (mức ăn ca riêng đ/ngày).
- mhNguon(nv, LenLuong, MucHuong): gộp khai báo hiệu lực + quyết định lên lương đã duyệt (gt luongMoi/luongBHMoi>0, cu luongCu/luongBHCu), sắp theo hieuLuc rồi thời điểm duyệt. mucHuong(nv, ll, mh, ngày): với từng khóa lấy mốc cuối ≤ ngày; nếu chưa có thì lấy cu của mốc sớm nhất sau ngày; không có mốc nào thì lấy hồ sơ. nvTaiKy(nv, ll, mh, ky) = hồ sơ thay mức tại cuối kỳ → dùng trong recalc và congthuc.test (SO_NPT, PC_*, giảm trừ, TIEN_AN_NGAY = anCa > 0 ? anCa : mức nơi làm việc).
- Hồ sơ đồng bộ mức hôm nay (dongBoMucHuong, gọi khi load/login, sau mh.*, ll.decide/delete). nv.save: nếu đã có diễn biến thì chặn đổi luongChinh, luongBH, soNPT, anCa, phuCap.
- Hành động: mh.save {row:{id?, maNV, hieuLuc, giaTri, soQD, lyDo}} (edit_nv hoặc approve; khóa lương bởi người không có approve → Chờ duyệt; đã hiệu lực chỉ admin sửa), mh.decide {id, approve, note} (approve), mh.delete {id} (hiệu lực: admin), mh.import {rows} (sắp theo ngày, báo lỗi từng dòng). Bỏ mốc: mốc kế tiếp cùng khóa nhận cu của mốc bị bỏ; không còn mốc nào thì hồ sơ về cu. Lưu/duyệt/xóa → tính lại các kỳ đang xử lý có cuối kỳ ≥ hiệu lực. load: nhân viên nhận mh của mình (Hiệu lực).
- Giao diện: hồ sơ tab "Lương & mức hưởng" (mức đang hưởng, ma trận diễn biến theo mốc – ô đổi tô màu, danh sách khai báo, lịch sử lên lương); form khai báo (placeholder "đang hưởng" theo ngày hiệu lực); Lương & nâng bậc → tab "Khai báo mức hưởng" (lọc người/trạng thái, duyệt, Tải mẫu Excel 2 sheet, Nhập Excel); huy hiệu menu cộng khai báo chờ duyệt. Mẫu TAPACO thêm PCTN = PC_TRACHNHIEM * HS_HUONG.
- Kiểm thử test/t5.js (21 kiểm tra).

## 5g. Bảng tin công ty (v1.9)

- **Dữ liệu**
  - Sheet BaiViet: id, loai (BT_LOAI: Thông báo | Chính sách – Quy định | Tin nội bộ | Truyền thông – Hoạt động | Vinh danh – Khen thưởng), tieuDe, noiDung, anh JSON [{fileId, ten, mime, w, h}], dinhKem JSON [{fileId, ten, mime, kichThuoc}], doiTuong JSON (phòng ban / đơn vị; rỗng = toàn công ty), ghim, xacNhan, soVanBan, hieuLuc, trangThai (Đăng | Nháp), tacGia, tenTacGia, taoLuc, dangLuc, capNhat.
  - Sheet TuongTac: id, baiId, loai (thich | xacnhan | binhluan), username, hoTen, noiDung, thoiGian.
- **Quyền**
  - PERM.bantin = admin, giamdoc, hr.
  - Ai được thấy bài: người có quyền bantin thấy mọi bài; người khác chỉ thấy bài đã Đăng, có đối tượng rỗng hoặc chứa phongBan/donVi trong hồ sơ của họ.
  - loadAll trả bt, tt theo quyền.
- **Hành động**
  - bt.upload: ảnh jpeg/png/webp/gif hoặc PDF/Word/Excel, ≤ dungLuongToiDaMB; GAS lưu vào thư mục "_BẢNG TIN CÔNG TY".
  - bt.save: sửa được khi là tác giả / giamdoc / admin; tệp bị bỏ khi sửa thì chuyển vào thùng rác.
  - bt.delete: xóa cả tệp và tương tác.
  - bt.pin.
  - bt.file: chỉ trả tệp thuộc bài người đó được thấy.
  - bt.like (bật/tắt), bt.ack (một lần, chỉ bài Đăng có xacNhan), bt.comment (≤ 2000 ký tự), bt.cdel (của mình, hoặc người có quyền bantin).
- **Giao diện**
  - Màn hình mặc định cho người không có view_all.
  - Phần đầu và thanh lọc: khung tiêu đề (hero), khung soạn nhanh, các nút lọc chuyên mục / "Cần tôi xác nhận" / "Bản nháp", ô tìm kiếm.
  - Thẻ bài: tác giả, thời gian tương đối, đối tượng, chuyên mục có màu, Ghim, Nháp, Mới; lưới ảnh 1–4 (+N) có xem phóng to; tệp đính kèm theo màu định dạng; thanh xác nhận đọc; Hữu ích; bình luận (2 gần nhất, mở rộng).
  - Cột phải: Cần bạn xác nhận, Đang ghim, Chính sách – quy định.
  - Thống kê xác nhận x/y, xuất Excel.
  - Ảnh tải qua bt.file và lưu đệm bằng object URL; bài mới tính theo localStorage tphr_btseen_<user>.
- **Kiểm thử**: test/t6.js (18 kiểm tra), test/smoke10.js.

## 5h. Nhắn tin nội bộ (v2.0)

- **Dữ liệu**
  - HoiThoai: id, loai (rieng | nhom), ten, mau, moTa, taoBoi, taoLuc, capNhat, tinCuoi JSON, trangThai ('' | Đã giải tán).
  - ThanhVien: id, hoiThoai, username, vaiTro (truong | thanhvien), daDoc (seq), ghim, tatTB, thamGia, roi.
  - TepChat: fileId, hoiThoai, ten, mime, kichThuoc, nguoiTai, thoiGian.
  - TinNhan: id, hoiThoai, username, hoTen, loai (text | anh | tep | hethong | thuhoi | camxuc), noiDung, tep JSON, traLoi, thoiGian.
  - Thu hồi và cảm xúc là dòng sự kiện, traLoi = id tin gốc; cảm xúc bật/tắt theo số lần.
- **env.chat** (Apps Script: chatEnv_, trình duyệt / kiểm thử: HRCORE.memChat)
  - Ghi đọc TinNhan:
    - TinNhan **không** đi qua SheetDB;
    - seq = số dòng;
    - append ghi 1 dòng định dạng '@';
    - last, get(seq), since(seq, max), before(seq, n) đọc theo khoảng dòng.
  - Tín hiệu qua CacheService:
    - bump(users), sig(u): báo danh sách hội thoại thay đổi;
    - online(u), onlineOf(us): trực tuyến 75 giây;
    - typing / typingOf: đang soạn tin 6 giây.
- **Hành động**
  - Đọc (không khóa): chat.init, chat.poll, chat.history, chat.file.
  - chat.init trả về: danh bạ, hội thoại kèm thành viên và daDoc, 3000 dòng cuối thuộc các hội thoại của mình, cur, hv.
  - chat.poll {cur, hv, typing, online}:
    - chỉ đọc các dòng sau cur;
    - trả ht khi có tin mới hoặc hv đổi;
    - trả reset khi cur > last.
  - Gửi và tương tác: chat.send {ht, noiDung ≤ 5000, traLoi, tep (phải là tệp của mình trong cùng ht)}, chat.upload, chat.react, chat.recall (tin của mình; xóa tệp), chat.read, chat.pref, chat.open.
  - Nhóm: chat.group, chat.gupdate, chat.gadd (mọi thành viên), chat.gremove, chat.grole (trưởng nhóm / admin), chat.gleave (tự chuyển trưởng nhóm), chat.gdisband (người tạo / admin). Mỗi thay đổi nhóm ghi một tin hethong.
- **Giao diện**
  - Bố cục: danh sách bên trái (tìm kiếm kèm danh bạ, ghim, tắt thông báo, số chưa đọc, đang soạn tin), khung chat bên phải. Nền kẻ ô kiểu giấy bản vẽ; bong bóng của mình màu navy, của người khác màu trắng; gom tin liên tiếp trong 5 phút; vạch phân cách theo ngày.
  - Nội dung tin: trích dẫn trả lời, ảnh và tệp, cảm xúc, đã xem, gửi lạc quan (hiện tin ngay khi đang gửi), kéo thả và dán ảnh.
  - Thông tin nhóm: thành viên và quyền, ảnh và tệp đã gửi.
  - Hỏi tin mới định kỳ: 3 giây / 15 giây / 45 giây (chạy nền).
  - Thông báo: thông báo trình duyệt, ô thông báo nhỏ, số chưa đọc trên tiêu đề tab.
  - Điện thoại: một khung, chạm vào tin để hiện thao tác.
  - Bản dùng thử: người nhận tự trả lời các tin nhắn riêng.
- **Kho tệp (v2.1)**
  - Dữ liệu: HoiThoai.thuMuc (mã thư mục Drive); TepChat thêm msgId, seq, w, h, thuMuc.
  - env.files thêm folder(tên) / renameFolder / move / folderUrl.
  - Thư mục Drive: chat.upload tạo thư mục con "_TIN NHẮN NỘI BỘ/Tên nhóm [6 ký tự mã]" (chat riêng: "Riêng – A & B [mã]") khi gửi tệp lần đầu; chat.gupdate đổi tên thư mục.
  - chat.send gắn msgId / seq vào TepChat; một tệp chỉ gửi được 1 lần.
  - chat.files {ht, loai anh|tep, nguoi, tu, den, q, tu0, n}: chỉ trả tệp trong hội thoại mình đang tham gia; kèm tổng số tệp, dung lượng, coThem; admin nhận thêm thuMucUrl, chuaXep.
  - chat.xep (admin): chuyển tệp cũ vào thư mục nhóm, 150 tệp mỗi lần.
  - Giao diện Kho tệp nằm trong khung chat:
    - bộ lọc: cuộc trò chuyện, loại (Tất cả / Ảnh / Tài liệu), người gửi, tháng, tìm tên tệp;
    - nội dung nhóm theo tháng: lưới ảnh và danh sách tài liệu;
    - xem ảnh phóng to; nút Đến tin nhắn tải lùi lịch sử tới khi gặp tin đó.
- **Kiểm thử**: test/t7.js (41), test/t8_gas.js (giả lập Apps Script, 10), test/smoke11.js, test/smoke12.js.

## 5i. Điện thoại & tin nổi bật (v2.2)

- **Thanh điều hướng dưới (#tabbar, ≤ 900px)**
  - Mục theo vai trò: bt, chat, rồi slip + mynp (có maNV) hoặc dash + luong (view_all), cuối cùng là "Thêm".
  - "Thêm" mở menu bên (#side), số đếm gộp các mục còn lại; badge theo badge().
- **Chat trên điện thoại**
  - body.vchat: khung chat tràn viền.
  - body.chfull (đang mở hội thoại / Kho tệp): ẩn #top, #banner, #tabbar; .ch cố định toàn màn hình (100dvh), có safe-area.
  - history.pushState khi mở hội thoại; popstate đóng hội thoại hoặc modal.
  - Ô nhập 16px để iOS không tự phóng to.
- **Tin nổi bật**
  - BaiViet.noiBat ('Có'), noiBatDen (ngày); bt.save nhận noiBat, noiBatDen.
  - Cửa sổ #btPop hiện 0,7 giây sau khi vào app, lần lượt các tin nổi bật còn hiệu lực mà người dùng chưa xem (localStorage tphr_btpop_<user> = {id: capNhat}).
  - Nút "Xem chi tiết" / "Đọc và xác nhận" → btOpenPost; "Tin tiếp"; "Để sau".
  - Dải .bt-strip đầu Bảng tin: tin nổi bật + tin ghim, cuộn ngang theo nấc; mỗi thẻ ảnh trên, chữ dưới.
  - Cờ tphr_nopop chỉ để kiểm thử tự động.
- **Kiểm thử**: test/smoke13.js (điện thoại).

## 5j. Phòng ban, công trình & điều động (v2.3)

- cfg.boPhan = [{ma, ten, loai: 'Phòng ban'|'Công trình', diaDiem, batDau, ketThuc, trangThai: 'Hoạt động'|'Đã kết thúc', ghiChu}]; cfg.phongBan suy ra (danh sách tên). Mặc định: BGD, TCHC, TCKT, KT, KTKH. mergeCfg chuyển phongBan cũ thành boPhan.
- NhanVien.phongBan lưu tên bộ phận; Users thêm cột boPhan (bộ phận do TP/CHT quản lý; trống = theo hồ sơ).
- bp.save {boPhan}: đổi tên theo ma lan sang NhanVien, QuaTrinh, Users, BaiViet.doiTuong; chặn xóa / kết thúc bộ phận còn người.
- bp.dieudong {boPhan, maNV[], tuNgay, soQD, chucVu, noiDung}: mỗi người một QuaTrinh 'Điều động' qua qt.save; khép quá trình trước. dongBoViTri() áp quá trình tương lai khi tới ngày. chuyenQuanLy(): Users.boPhan của TP/CHT đi theo người được điều động.
- Giao diện: view `bp` (thẻ theo tab Phòng ban / Công trình / Đã kết thúc; Điều động đến, Sửa, Lịch sử điều động), view `bpme` cho TP/CHT (KPI, đơn chờ duyệt, lịch nghỉ, nhân sự, nút Nhắn tin; tab dưới "Bộ phận"). Form hồ sơ / quá trình chọn bộ phận từ danh mục. Tài khoản: chọn CBCNV → tự gợi ý tên đăng nhập; chức danh; Bộ phận quản lý cho TP/CHT; thẻ ai giữ chức danh; mô tả quyền (ROLES_MOTA); bảng phân quyền.
- Kiểm thử: test/t9.js (28), test/smoke14.js.

## 5k. Lịch lễ, chấm công theo ngày, tạm ứng, nhóm vi phạm (v2.4)

- **Âm lịch**: thuật toán Hồ Ngọc Đức, múi giờ +7; amSangDuong(d, m, y). Kiểm: Tết 2026-02-17, 2027-02-06; Giỗ Tổ 2026-04-26.
- **lichLeTheoLuat(nam)**:
  - Ngày lễ: 1/1; cuối năm Âm lịch + mùng 1–4; 10/3 Âm lịch; 30/4; 1/5; 2/9.
  - Ngày liền kề Quốc khánh: 1/9 nếu 2/9 rơi vào T3, T4 hoặc T7; ngược lại 3/9.
  - 24/11 từ năm 2026.
  - Ngày trùng Chủ nhật thì thêm "Nghỉ bù" vào ngày làm việc kế tiếp chưa nghỉ.
- **Cấu hình lịch lễ**:
  - cfg.ngayLeDs = [{ngay, ten, loai: Lễ, Tết | Nghỉ bù | Công ty cho nghỉ}]; cfg.ngayLe suy ra từ đây. mergeCfg chuyển ngayLe cũ.
  - Action le.save {ds}: quyền edit_nv hoặc config.
  - View `le` (p 'all'): lịch 12 tháng, danh sách ngày, nạp theo luật, thêm theo khoảng ngày.
- **ChamCong.chiTietNgay** = {ngày: ký hiệu}.
  - KY_HIEU: X 1 công; X2 0,5; CT 1; P/R hl 1; O om 1; KL kl 1; KP kl 1 + kp; L le 1; P2/R2/O2/KL2 = 0,5 nghỉ + 0,5 công.
  - ngayTheoDon(): ô khóa từ đơn đã duyệt (bỏ CN, ngày lễ; đơn nửa ngày → mã "2").
  - cleanNgay(ct, ky, khoa); tongHopNgay(ct, ky, cfg) → ngayCong, nghiHuongLuong, nghiKhongLuong, nghiOm, ngayLe, ot300 = ngày làm lễ × giờ; ot200 = ngày làm CN × giờ nếu cfg.chuNhatTinh = 'ot'.
  - cc.save có chiTietNgay thì ghi đè các cột trên.
  - np.decide (duyệt) / np.cancel gọi dongBoNgayDon(): ô đơn bị hủy → X; áp ô khóa mới; tính lại tổng hợp.
- **ChamCong.chiTietTamUng** = [{loai ∈ cfg.loaiTamUng, soTien > 0, ngay, ghiChu}].
  - tamUng = tổng.
  - Đã có bảng kê thì bỏ qua tamUng nhập tay.
  - FIELDS_LOAI.congno gồm chiTietTamUng.
- **chiTietPhat** thêm `nhom` ∈ NHOM_VP: Nghỉ không phép, Vi phạm hợp đồng, Mất mát, hư hỏng tài sản, Vi phạm nội quy – an toàn, Vi phạm khác. Bồi thường mặc định nhóm Mất mát; danhMucViPham cũng có nhom.
- **BangLuong** thêm ctTamUng, ctViPham, ctCong để CBCNV xem chi tiết trên phiếu lương.
- **Giao diện tab lương**: ngay (bảng chấm công ngày: thanh ký hiệu, tô bằng chuột hoặc chạm, tô cả cột / hàng, Điền mặc định, Xuất Excel), cc (tổng hợp), tu (tạm ứng), vp (thẻ tổng theo nhóm, gợi ý ghi vi phạm từ ô KP).
- **Kiểm thử**: test/t10.js (17), test/smoke15.js.

## 5l. Phiếu lương chi tiết & báo cáo Excel có công thức (v2.5)

- **Phiếu lương** (slipHTML) gồm:
  - Ô thông tin: lương chính, lương đóng BH, công hưởng lương, số người phụ thuộc.
  - Dòng ngày công lấy từ BangLuong.ctCong {c, hl, le, om, kl, ot150, ot200, ot300, npt, gtbt, gtnpt, lbhHS}.
  - I. Thu nhập theo phanLoai, ghi phần miễn thuế (gt − thue).
  - II. BH và đoàn phí.
  - III. Thuế: không chịu thuế, chịu thuế, giảm trừ bản thân / người phụ thuộc, BH, tính thuế, thuế.
  - IV. Tạm ứng (ctTamUng), khấu trừ công nợ, bồi thường (ctViPham), khấu trừ công thức.
  - V. Còn nhận; dòng phần doanh nghiệp đóng.
- **Core**:
  - dauVaoTinhLuong(nv, ky, ncc, ll, mh, kt) → {nv, o}; recalc dùng chung.
  - astExcel / ctExcel: AST công thức → Excel.
    - Chia an toàn: IF(r=0,0,l/r).
    - So sánh → (..)*1. IF(c<>0,…). FLOOR → INT. CEIL → -INT(-x). AND/OR/NOT → *1.
- **Workbook spec** {ten, sheets:[{ten, rows:[[cell]], cols, merges, freeze}]}, cell = giá trị | {v, f, z, s}. sangXLSX(spec, XLSX) dùng xlsx-js-style: style title / sub / note / code / head / grp / tot / bold / net / sign / draft, viền bảng, numFmt.
- **xlBangTinh(ky, ds, cfg, ctl, dbh, info)**:
  - Sheet ThamSo: tham số BH / thuế / giảm trừ, biểu thuế; tham chiếu tuyệt đối.
  - Sheet TinhLuong: hàng 4 nhóm cột (merge), hàng 5 tên, hàng 6 mã; dữ liệu từ hàng 7; dòng TỔNG CỘNG.
  - Cột: đầu vào (BIEN_CO_BAN + biến dùng), cờ TG_BH / DOAN_VIEN / VANG_LAI / CAM_KET_08, biến dẫn xuất, trung gian, khoản ROUND(…,0), TONG_TN, MIEN_THUE, DN_GIAM, CO_THE_GIAM, GIAM_QC, TONG, DONG_BH, LUONG_DONG_BH, BHXH / BHYT / BHTN (trần), DOAN_PHI, TNCT, GIAM_TRU, TNTT, THUE (lũy tiến qua bảng bậc: MAX(0, IF(T=0, X, MIN(X, T)) − T_trước) × ts), bồi thường (tối đa %), tạm ứng, công nợ, khoản khấu trừ, CON_NHAN, phần DN, đối chiếu app, chênh lệch.
- **Báo cáo khác**:
  - xlBangLuongThang: nhóm bộ phận, SUBTOTAL(9,…), ký; sheet ChuyenKhoan.
  - xlNam: TongHop 12 tháng + ma trận ThuNhap / ThueTNCN / TNChiuThue / ConNhan.
  - xlCaNhan.
  - xlChiPhi: DuLieu + TheoBoPhan / TheoThang bằng SUMIFS / COUNTIFS, tỷ trọng.
- **View `bc`** (view_all), 5 tab: thang, nam, canhan, chitiet (bcBuoc: modal từng bước tính), chiphi.
- **Kiểm thử**: test/t11_xl.js dựng dữ liệu 6 người (VP / DA, thử việc, HĐ dưới 3 tháng, đoàn viên, NPT, OT, vi phạm, bồi thường, tạm ứng, công nợ), xuất file không có giá trị lưu sẵn → `soffice --convert-to xlsx` → test/xlcheck.py so mọi ô công thức với giá trị của app (mẫu tapaco 611, chuan 620 công thức, 0 lệch). test/smoke16.js tải 5 file từ giao diện.
- **Lưu ý CSS**: không đặt class `.lg` cho phần tử khác vì trùng modal `.md.lg` (lịch lễ dùng `.lgr`).

## 5m. Mẫu Excel tạm ứng (v2.5.1)

- **xlMauTamUng(ky, nvs, loai, info)**:
  - Sheet TamUng: STT, Mã NV, Họ tên, Bộ phận, một cột cho mỗi loại, Tổng = SUM, Ngày, Ghi chú; dòng TỔNG CỘNG.
  - Sheet ChiTiet: Mã NV, Họ tên, Loại, Ngày, Số tiền, Ghi chú.
  - Sheet HuongDan.
- **docTamUng(sheets, nvs, loai, ky)**:
  - Tìm dòng tiêu đề có "Mã NV" hoặc "Họ và tên".
  - Dạng rộng: mỗi cột là một loại, nhận diện bằng nhanLoaiTU (khớp tên, rồi theo từ khóa lương / xăng – đi lại – công tác / tiền ăn – cơm / khác).
  - Dạng dài: cột "Loại" + "Số tiền".
  - Người: theo Mã NV, hoặc theo họ tên nếu trùng duy nhất.
  - docTien nhận 1.500.000, 1,500,000, 1,5tr, 500k, kiểu số. docNgay nhận số ngày Excel, dd/mm/yyyy, dd/mm.
  - Bỏ qua dòng Tổng / Cộng. Trả về {items, loi, canhBao}.
- **Client**:
  - tuMau() tải file mẫu.
  - tuNhap() đọc file → modal xem trước (tổng theo loại, lỗi, cảnh báo, bảng khoản, chế độ Cộng thêm / Thay) → tuNhapLuu() gọi cc.save với chiTietTamUng.
  - Chỉ nhận người đã có dòng chấm công trong kỳ.
- **Kiểm thử**: test/smoke17.js.

## 5n. Mã NV, tài khoản, mật khẩu (v2.6)

- **Mã NV**: maNVMoi() = 'NV' + 4 chữ số tiếp theo số lớn nhất trong các mã NV\d+. nv.save isNew với mã trống, và nv.import với ô Mã trống, đều tự sinh.
- **Tên đăng nhập**: tenDangNhap(hoTen, daCo, ngaySinh) = tên + chữ đầu họ đệm + ddmmyy (thanglv201190); trùng thì thêm số.
- **Mật khẩu ban đầu**: matKhauTheoNV(nv, cfg) = tên không dấu, chữ thường + năm sinh (thang1990). Thiếu ngày sinh hoặc mật khẩu dưới 6 ký tự thì dùng cfg.matKhauMacDinh. mustChange = 'Có'.
- **user.bulk** trả thêm mkBanDau (không lưu).
- **PERM.doimk** = [admin, giamdoc]:
  - user.resetpw {username, password?}: trống thì về mật khẩu mặc định và trả lại mật khẩu đó; không tự đặt cho mình; không đổi tài khoản admin trừ khi chính admin.
  - user.save chỉ đổi mật khẩu tài khoản đã có khi người gọi có doimk.
  - Người dùng tự đổi qua changePassword.
- **View users**: mở cho users hoặc doimk. Giám đốc bấm một dòng để đặt lại mật khẩu.
- **Kiểm thử**: t10 (thêm 11 kiểm tra), smoke18.

## 5o. Tin nhắn kiểu Zalo, mục nhóm, thông báo (v2.7)

- **Sheet mới:**
  - ChatMuc: id, hoiThoai, loai (binhchon | nhac | baocao), tieuDe, noiDung, cauHinh JSON, taoBoi, taoLuc, trangThai (mo | dong | xong | xoa), ketTiep 'yyyy-mm-dd HH:mm', lanCuoi, capNhat, nhacHan.
  - ChatTraLoi: id, mucId, ky (ngày báo cáo; trống với bình chọn), username, giaTri JSON, thoiGian.
- **Vai trò nhóm:** ThanhVien.vaiTro = truong | pho | thanhvien.
  - laQL = trưởng, phó hoặc admin: gupdate, gremove (phó không mời trưởng / phó khác), tạo báo cáo.
  - chat.grole chỉ trưởng nhóm.
- **Handlers:**
  - chat.binhchon {ht, tieuDe, luaChon[], nhieu, themLC, han}; chat.bau {mucId, chon[]} (thay phiếu cũ); chat.themlc.
  - chat.nhac {ht, tieuDe, luc, lap: mot | ngay | tuan | thang}.
  - chat.baocao {ht, tieuDe, moTa, mau[{ten, kieu: text | dai | so | chon | cokhong | ngay, luaChon, batBuoc}], gio, han, thu[1..7], nguoi[], ngayNay} (QL).
  - chat.nopbc {mucId, ky, giaTri}: kiểm tra người được giao, mục bắt buộc, kiểu; upsert; ghi tin 'bcnop'.
  - chat.mucdong {mucId, xoa}; chat.lich (admin).
- **Tin nhắn mới:**
  - Loại: binhchon, nhactao, nhac, bctao, baocao, bcnop.
  - traLoi = mucId; tep = {ky} hoặc {luc}, {han, con}.
  - chat.init và chat.poll (khi đổi danh sách) trả thêm muc = mucDs(s).
- **chayLich()** (api.chayLich):
  - Nhắc hẹn: đến ketTiep thì ghi tin 'nhac', rồi lùi theo chu kỳ lặp (mot → xong).
  - Báo cáo: đến ketTiep thì phát 'baocao' cho ngày gần nhất (bỏ ngày đã lỡ), ketTiep = bcKeTiep(thứ hợp lệ, giờ). Tới han thì nhắc tên người chưa nộp 1 lần/ngày (nhacHan); phát trễ sau hạn thì không nhắc ngày đó.
  - Kích hoạt: GAS hàm nhacTuDong + trình kích hoạt 5 phút (caiDatNhacTuDong); bản demo đặt env.lichTrongPoll để chạy trong poll.
- **Giao diện:**
  - Danh sách: .zl-top xanh #0091FF→#0068FF, lọc chip, ảnh đại diện nhóm ghép.
  - Khung chat: header xanh "Bấm để xem thông tin", nền #E2E9F1, bong bóng của mình #E5EFFF; tên người gửi và giờ trong bong bóng; ♥.
  - Thẻ: .zc nh / al / bc / bv. Dải nhắc .zl-pin. Sheet "…" (.zl-sheet).
  - Thông tin nhóm: tab Thành viên, Nhắc hẹn · Bình chọn, Báo cáo, Ảnh tệp; chọn vai trò.
  - BC_MAU (4 mẫu + mẫu trống).
- **Thông báo:**
  - Cài đặt trong localStorage tphr_tb_<user> {bat, xemTruoc, am, rung}.
  - App chạy nền: registration.showNotification (tag theo hội thoại, renotify, vibrate, data.ht). Đang mở: banner #chToast trượt từ trên xuống + âm WebAudio.
  - sw.js notificationclick → postMessage openChat hoặc openWindow('#chat=ID'). navigator.setAppBadge theo số chưa đọc.
- **Kiểm thử:** test/t12.js (29), test/smoke19.js.

## 5p. Lưu trữ nhóm chat & phân tích AI (v2.8)

- **Drive:** thư mục gốc "TAPACO – Lưu trữ nhóm chat (KHÔNG CHIA SẺ)" tạo trong thư mục cha của file Google Sheet (thuMucLuuTru_, Script Property LT_ROOT). Mỗi nhóm: thư mục chatTenThuMuc(h) gồm `0. Thông tin nhóm.md`, `1. Tệp & ảnh` (= h.thuMuc; thư mục tệp cũ được moveTo và đổi tên), `2. Nhật ký từng ngày` (YYYY-MM-DD.md), `3. Báo cáo định kỳ` (<tiêu đề> [id4].csv, BOM UTF-8), `4. Phân tích AI` (Google Docs + PhanTich_<tu>_<den>_<yyyyMMdd-HHmm>.json). HoiThoai thêm cột luuTru JSON {goc, tep, nk, bc, ai, luc, den, dang, batDau}. Tin riêng 1–1 không lưu trữ.
- **env.luuTru** (GAS luuTruEnv_ / demo demoDrive): thuMuc(cha, ten) get-or-create, doiTen, chuyen, ghi(cha, ten, nd, mime) tạo hoặc setContent, ds(cha), doc(id), anh(id, max) (ảnh lớn → getThumbnail), taoDoc(cha, tieuDe, khoi) (DocumentApp: h1/h2/h3/p/li/nho/bang/anh), url, canhBao (sharing ≠ PRIVATE hoặc có editor/viewer khác chủ).
- **Lưu trữ:** ltTin (toàn bộ tin nhóm; thu hồi → không lưu nội dung; gộp cảm xúc bật/tắt; su = mốc thay đổi + ngày bị ảnh hưởng), ltCtx (muc, bc, phat theo tin 'baocao', tl), ltNhatKy (markdown), ltCSV, ltThongTin. ltChay(h, gioiHan): ngày "bẩn" = mọi thay đổi có thoiGian ≥ luc (≥ để không sót tin cùng giây) + ChatTraLoi nộp sau luc; ghi tối đa gioiHan ngày, còn thì lưu dang để lần sau làm tiếp; xong thì luc = batDau. chat.luutru (QL nhóm / giamdoc / admin, có khóa ghi). luuTruTuDong(ms) → {soNgay, xong}: nhacTuDong gọi mỗi ngày đến khi xong (Script Property LT_NGAY), mỗi lượt ≤ 50 giây.
- **Phân tích AI** (READ_ONLY, không giữ khóa ghi; trạng thái nhiều bước trong env.tam = CacheService chia khúc 30.000 ký tự): ai.phantich {buoc: batdau → {job, dot}; phan i → AI đọc 1 đợt (aiDot: ≤ 7 ngày, ≤ 60.000 ký tự); tonghop → thống kê tất định ltThongKe + AI đánh giá → kq, tạo Google Docs + JSON}. ai.nhom, ai.ds, ai.xem, ai.tep (ảnh chứng từ cho người xem không ở trong nhóm), ai.tepluu, ai.doctep, chat.nhatky (dữ liệu Excel nhật ký), ai.cauhinh (admin), ai.thu (admin).
- **Quyền:** ltQuyen = admin, giamdoc mọi nhóm; trưởng/phó nhóm của mình.
- **AI_HE / AI_P_DOT / AI_P_TONG:** chỉ dùng dữ liệu, không bịa; mỗi số vật tư/khối lượng kèm ngày, tài khoản, trích dẫn; JSON duy nhất. aiSachDot/aiSachTong làm sạch: ngày ngoài đợt, tài khoản lạ, mã tệp không thuộc nhóm bị bỏ; mức đánh giá, xếp loại theo danh sách cố định; aiSo đọc số kiểu Việt (2,5 / 1.200). aiGop cộng theo khóa (tên + đvt + hoạt động).
- **env.ai (GAS aiEnv_):** Script Properties AI_NCC, AI_KEY_GEMINI / AI_KEY_CLAUDE, AI_MODEL_GEMINI / AI_MODEL_CLAUDE (mặc định gemini-flash-latest; v3.13.3: lỗi model bị ngừng → `chonModelGemini_` theo ListModels + model gợi ý trong lỗi, gọi lại, lưu AI_MODEL_GEMINI; Gemini 3.x thinkingLevel 'low', 2.5-flash thinkingBudget 0; kiểm thử t37_gas.js; claude-haiku-4-5-20251001). Gemini: v1beta generateContent, header x-goog-api-key, systemInstruction, responseMimeType application/json, thinkingBudget 0 cho 2.5-flash (bỏ nếu lỗi). Claude: /v1/messages, x-api-key, anthropic-version 2023-06-01, max_tokens ≤ 4096. Thử lại 1 lần khi 429/5xx/529; lỗi 401/404/429/hết 60 giây báo tiếng Việt. Khóa không bao giờ trả về client.
- **Giao diện:** trang aibc "Phân tích nhóm (AI)" (danh sách nhóm, lưu trữ, xem tệp đã lưu, chọn khoảng ngày, tiến trình từng bước, báo cáo đã lập), xem báo cáo 6 tab (Tổng quan, Nhật ký ngày, Cá nhân, Vật tư, Khối lượng, Chứng từ), Xuất Excel (C.xlAiBaoCao, 7 trang), In/PDF (C.aiKhoi → HTML), Xuất nhật ký (C.xlNhatKy). Bản dùng thử: Drive mô phỏng, AI mô phỏng ghi rõ "không phải AI thật", 10 ngày lịch sử nhóm "BCH công trường mẫu".
- **Kiểm thử:** test/t13.js (52), test/t14_gas.js (Drive/Docs/UrlFetch giả lập, Gemini + Claude), test/smoke20.js.

## 5q. Đánh giá tổng thể, nhật ký thi công, Word/PDF (v2.9)

- **Phạm vi:** ltQuyen cho admin, giamdoc phân tích cả trò chuyện riêng (ltCay đặt dưới "_Trò chuyện riêng"; không tự lưu mỗi đêm); trưởng/phó chỉ nhóm của mình. ai.nhom trả thêm loai = rieng (tên "A & B").
- **ltThongKe:**
  - soNgayLV = ngày làm việc (laNgayNghiLe: CN + cfg.ngayLe); mỗi người tính từ max(tu, min(thamGia, ngày hoạt động đầu)) đến min(den, roi).
  - soNgayCoTT = ngày LV có tin text/anh/tep/bcnop/binhchon/nhactao/bctao, nộp báo cáo, bình chọn.
  - soNgayKhongTT, ngayKhongTT (≤ 40) = ngày LV không tương tác và không nghỉ.
  - soNgayNghi, nghiLoai từ NghiPhep 'Đã duyệt' (½ ngày = 0,5); nghiTheoNgay cho nhật ký.
  - diemTG = 50·coTT/(LV − nghỉ) + 30·đúng hạn/phải nộp (không có báo cáo thì 80·coTT/…) + 20·đóng góp/đóng góp lớn nhất.
- **Các bước ai.phantich:**
  - **batdau:** đợt ≤ 3 ngày (Claude 2), ≤ 45.000 ký tự.
  - **phan:** AI_P_DOT trả mảng ngay[] gồm thoiTiet{sang, chieu, anhHuong}, nhanLuc{tong, banChiHuy, congNhan, thauPhu[]}, thietBi[], thiCong[{hangMuc, viTri, khoiLuong, dvt, keHoach, soVoiKH}], vatTuNhap, vatTuXuat, yeuCauVatTu, nghiemThu, kienNghi[{username}], anToan{tinhHinh, baoHo, viPham}, anNinh, suCo, tonTai, keHoachMai; kèm doiTac[], suKien[{loai, mucDo}], chungTu, caNhan. aiSachNgay / aiSachDot làm sạch.
  - **tong1:** AI_P_TONG1 trả tongThe, mucRuiRo, lyDoRuiRo, danhGia đủ 18 AI_KHIA theo thứ tự (aiSachT1 điền "Chưa đủ dữ liệu").
  - **tong2:** AI_P_TONG2 trả caNhan{xepLoai, nhanXet, deXuat}, doiTac{danhGia}, diemManh, tonTai, giaiPhap{vanDe, giaiPhap, phuTrach, thoiHan, uuTien} (sắp theo ưu tiên), kienNghiBGD.
  - **lap:** aiGopDot gộp nhatKy (+ số đếm, nghiPhep), khoiLuong / khoiLuongTong (tổng + kế hoạch), vatTu / vatTuTong (nhập / xuất / đề nghị), thietBi (không phải "hoạt động"), kienNghi theo người, doiTac gộp theo tên; tạo Docs + PDF (taoDoc opt.pdf → getAs PDF) + JSON.
  - **lapnk:** nkKhoi → Docs + PDF; ghi lại JSON (nkDocId, nkPdfId).
  - ai.taifile tải PDF / tệp (không tải Google Docs).
- **Khối nội dung chung** (aiKhoi, nkKhoi): dau, h1, h2, h3, giua, p, nho, li, bang{rows, rong %}, anh, ngat, ky → Google Docs (taoDocBaoCao_), Word (client: thư viện docx 9.8.1 iife tải từ jsDelivr khi bấm, khoiSangDocx, A4, Times New Roman 12), HTML in (khoiSangHTML). nkO(n, truong, kq) dựng chữ từng ô nhật ký (dùng chung giao diện / Excel / Docs).
- **Excel xlAiBaoCao (9 trang):** Tổng quan, Giải pháp, Nhân sự, Nhật ký thi công, Khối lượng, Vật tư – thiết bị, Thầu phụ – NCC, Chứng từ, Sự kiện.
- **Giao diện:**
  - Danh sách lọc Nhóm / Trò chuyện riêng; tiến trình theo khóa bước (lt, bd, p0…, tong1, tong2, lap, lapnk), chạy lại từ bước lỗi.
  - Báo cáo 9 tab; aiChuan chuẩn hóa báo cáo bản 2.8.
  - Thông tin nhóm có dòng thông báo kênh công việc.
  - Bản dùng thử: demoPhan / demoTong1 / demoTong2 (bộ phân tích mẫu, ghi rõ không phải AI); seed có mục "Ban chỉ huy có mặt" và 1 đơn nghỉ đã duyệt.
- **Kiểm thử:** t13 (76), t14_gas (Docs + PDF + tải PDF, Gemini 8000 / Claude 6000 token), smoke20 (Word mở bằng LibreOffice ra PDF để soát).

## 5r. Chấm công GPS + khuôn mặt (v3.0)

- **Sheet:**
  - ChamCongGPS (id, maNV, username, hoTen, ngay, gio, loai vao/ra, diemId, tenDiem, boPhan, lat, lng, doChinhXac, khoangCach, banKinh, anh, khuonMat, diemKM, thietBi, trangThai, lyDo, nguoiDuyet, duyetLuc, ghiChu, thoiGian, tre);
  - KhuonMat (maNV, username, hoTen, mau = JSON các vector 128 chiều, anh, thietBi, dangKyLuc, trangThai, xoaBoi, xoaLuc).
- **Cấu hình:**
  - cfg.diemCC {maBoPhan | _chung: [{id, ten, lat, lng, banKinh 20–5000, gioVao?, gioRa?}]} lưu riêng, nên bp.save không làm mất;
  - cfg.gps {bat, doChinhXac 50, hieuChuanGiay 25, nguongKM 0,5, gioVao 07:30, gioRa 17:00, phutTre 10, gioDuCong 6};
  - load trả cả gps, diemCC cho người không có view_all.
- **Hàm chung:** khoangCachM (haversine) và kmKhoangCach (Euclid vector).
- **API:**
  - gps.cuatoi: điểm được phép (bộ phận + chung), mẫu KM, lượt tháng.
  - gps.dangky: ≥ 2 mẫu, chỉ 1 lần.
  - gps.cham:
    - từ chối nếu độ chính xác > ngưỡng, không có mặt hoặc thiếu ảnh;
    - ngoài vùng: ghi 'Sai vị trí', ghiNhan = false;
    - trong vùng: so KM, thiết bị, nghiGia, rồi trả 'Hợp lệ' hoặc 'Chờ duyệt';
    - tính tre theo giờ điểm hoặc giờ chung;
    - ảnh lưu `_CHẤM CÔNG GPS/yyyy-mm`.
  - gps.ds: view_all, hoặc ROLE_QL cho bộ phận mình.
  - gps.anh.
  - gps.duyet: không duyệt lượt Sai vị trí hay lượt của chính mình; Từ chối cần lý do.
  - gps.xoakm (edit_nv).
  - gps.diem: edit_nv, giamdoc, hoặc QL của bộ phận.
  - gps.caidat.
  - gps.tonghop (view_all): X nếu vào + ra ≥ gioDuCong giờ, X2 nếu thiếu.
- **Client:**
  - Thư viện @vladmandic/face-api 1.7.15 (jsDelivr): tinyFaceDetector, faceLandmark68TinyNet, faceRecognitionNet. Nạp sau khi camera chạy; **tf.setBackend webgl → cpu** vì gói không kèm tệp wasm.
  - Hiệu chuẩn bằng watchPosition (enableHighAccuracy): ≥ 3 mẫu đạt ngưỡng và ổn định, lấy trung bình có trọng số.
  - Nghi giả khi tọa độ trùng tuyệt đối và độ chính xác ≤ 3 m.
  - Mã thiết bị lưu ở localStorage tphr_thietbi.
  - Ảnh lật gương, in họ tên, giờ, tọa độ.
  - Màn hình: ccgps (Cá nhân, thanh dưới "Chấm công"), gpsql (Quản lý, 6 tab).
  - dgTuGPS chỉ điền ô trống.
  - Bản dùng thử có mô phỏng tat / trong / ngoai.
- **Kiểm thử:** t15 (GPS core), smoke21 (mô phỏng, quản lý, định vị + camera giả lập Chromium, chặn quyền).

## 5s. Giao diện điện thoại (v3.0.1)

- Cột lưới trên điện thoại dùng minmax(0,1fr) (bt-wrap, bt-main, bt-side, lcal, lgr, ctl-wrap). Không dùng 1fr vì cột sẽ nở theo nội dung dài và bị cắt mép phải.
- Khối CSS @media (max-width:600px) cuối thẻ style:
  - input, select, textarea 16 px; select tối đa 100% chiều rộng;
  - nhãn KPI 12 px; .lh 12 px; ô ngày 27 px;
  - danh sách ngày nghỉ .le-ls thay bảng .le-tb;
  - tiêu đề bảng tin xuống dòng được, nút Đăng bài không co.
- Màn hình ≤ 360 px: lịch 1 cột.
- Kiểm thử test/mob_survey.js (390 px, isMobile): mọi màn hình phải có scrollWidth = 390; chỉ dải tin nổi bật, thanh tab và thanh bước được cuộn ngang.

## 5t. Thông báo đẩy, bản đồ, công tác (v3.1)

- **Sheet mới:**
  - ThietBiPush (id, username, token, thietBi, tenMay, tuyChon JSON {tn, bt, yk, duyet, cc, xemTruoc}, dangKy, capNhat);
  - CongTac (id, maNV, username, hoTen, boPhan, tuNgay, denNgay, cheDo diem|linhhoat, diem JSON [{ngay?, ten, lat, lng, banKinh, diaChi}], noiDen, mucDich, phuongTien, trangThai Chờ duyệt|Đã duyệt|Từ chối|Đã hủy, nguoiTao, nguoiDuyet, duyetLuc, lyDo, taoLuc, capNhat);
  - ChamCongGPS thêm cột congTac.
- **Cấu hình:** cfg.push {bat, firebase (cấu hình web, công khai, trả trong load), vapid}. Khóa tài khoản dịch vụ chỉ lưu ở Script Properties FCM_TAI_KHOAN_DICH_VU (env.push.luuKhoa / khoa).
- **Hàng đợi đẩy (core):**
  - bao(users, {loai, t, b, tag, mo}) xếp vào PUSH.
  - handle() trả _push = pushDs() (đổi theo thiết bị, bỏ loại đã tắt, ẩn nội dung nếu xemTruoc = false).
  - Apps Script handle_ xóa _push khỏi phản hồi và gửi **sau khi nhả khóa** (pushSau_).
  - nhacTuDong lấy api.layPush().
- **Điểm phát sự kiện:**
  - chatGhi → chatDay: text, anh, tep, binhchon, nhac, baocao, bcnop, nhactao, bctao; bỏ người gửi (trừ nhac / baocao) và người có tatTB.
  - bt.save: chỉ lần đăng đầu, theo btThay; bt.comment; yk.create / yk.reply.
  - np.create / np.decide: npBao cho cấp duyệt hiện tại; kết quả báo người làm đơn.
  - ky.move: người của bước mới; Đã duyệt → toàn bộ người có phiếu.
  - gps.ct.tao / duyet.
- **Gửi FCM (Apps Script):**
  - JWT RS256 (Utilities.computeRsaSha256Signature) → oauth2 token, cache 3300 s.
  - FCM HTTP v1 messages:send, data-only (mọi giá trị là chuỗi), webpush Urgency high, TTL 86400.
  - fetchAll theo lô 50.
  - Cùng token + tag trong 20 s thì data.im = '1' (không đổ chuông lại).
  - 404 / UNREGISTERED → xoaTokenPush; 401 → xóa cache mã truy cập.
- **API:** push.dangky / huy / tuychon nhận mã thiết bị qua **maMay** (không dùng tên token, trùng phiên đăng nhập); push.cuatoi, push.thu, push.trangthai, push.caidat (config).
- **sw.js:**
  - push: đọc p.data || p.
  - App đang hiển thị và focus (không phải iOS) → postMessage {type: 'push'} → client hiện banner + âm báo (tin nhắn thì kích vòng chat.poll).
  - Ngược lại showNotification(t, {body, tag, renotify: im !== '1', vibrate [220,90,220,90,220], requireInteraction cho loai = duyet, data {mo}}).
  - notificationclick → postMessage {type: 'mo'} hoặc openWindow('./index.html' + mo).
  - pushsubscriptionchange.
- **Client:**
  - firebase@10.14.1 compat (jsDelivr): messaging.getToken({vapidKey, serviceWorkerRegistration}).
  - pushKhoiDong làm mới token sau đăng nhập; pushMoi mời bật (3 ngày một lần); đăng xuất → push.huy.
  - moTheo(#chat=, #bt=, #v=view).
  - Đăng xuất không xóa phiên mới; lỗi phiên của yêu cầu cũ không đăng xuất người mới.
- **Bản đồ:**
  - Leaflet 1.9.4 (cdnjs) + OSM (cfg.gps.banDo thay được); tìm địa chỉ bằng Nominatim (countrycodes = vn).
  - Bộ chọn dùng chung BDC (bấm / kéo ghim divIcon, vòng bán kính).
  - Lớp khung bản đồ là .bdo (không dùng .bd, trùng huy hiệu menu).
- **GPS core:**
  - gpsDiemCua(nv, ngay) gộp điểm bộ phận + _chung + 'nv:MANV' (caNhan) + điểm công tác trong ngày (congTac, choDuyet).
  - Công tác linh hoạt: ngoài mọi vùng vẫn ghi nhận tại chỗ (linhHoat, banKinh 0).
  - Công tác chưa duyệt → lý do 'Công tác chưa được duyệt'; ctCapNhatLuot khi duyệt / từ chối / hủy.
  - Không tính trễ cho lượt công tác.
  - gps.tonghop: ngày có lượt công tác hợp lệ → 'CT'; ngày công tác không chấm → ghiChu.
  - gps.diem nhận 'nv:MANV'; gps.caidat thêm banKinhCT (50–5000) và banDo.
- **Client quản lý:**
  - Thêm tab Bản đồ, Công tác.
  - gqNVs gộp APP.D.nv và APP.D.bpNV (trưởng bộ phận / CHT thấy đủ người của mình).
- **Kiểm thử:** t16 (43 mục, đẩy + công tác), t17_gas (FCM: JWT thật, lô, im, xóa token, nhắc hẹn qua trình kích hoạt), smoke22 (bản đồ, công tác, sự kiện push thật qua CDP ServiceWorker.deliverPushMessage với Chromium headless mới).

## 5u. Kiểm tra chấm công GPS bằng AI (v3.2)

- **Quyền:** CCKT_ROLES = admin, giamdoc, tchc, ktluong, ktcongno chạy kiểm tra, nhắc, xử lý. ROLE_QL (truongphong, cht) chỉ xem cảnh báo **đã nhắc** của bộ phận mình và giải trình.
- **Sheet:**
  - KiemTraCC (id, tu, den, pham, taoBoi, taoLuc, soNguoi, soLuot, soDo, soVang, thongKe JSON, aiNhanXet JSON, aiLuc);
  - CanhBaoCC (id, khoa = maNV|ngay|ma[|thiết bị], dotId, maNV, hoTen, boPhan, ngay, nhom pv|ng|tg, ma, mucDo do|vang, moTa, luotIds JSON, trangThai Mới|Đã nhắc|Đã giải trình|Đã xử lý|Bỏ qua, nhacLuc, nguoiNhac, hanGT, giaiTrinh, ketQua, nguoiGT, gtLuc, xuLy, nguoiXL, xlLuc, taoLuc, capNhat).
- **ccktQuet(tu, den, boPhan)** – quy tắc, không dùng AI:
  - **theo lượt:** PV_BIEN (≥ 0,9 bán kính); PV_ACC (> 0,8 ngưỡng); PV_GIA (≤ 2 m); NG_KM; NG_CSK; NG_NGUONG (≥ ngưỡng − 0,08; đọc số thập phân bằng soThuc vì num() coi dấu chấm là phân cách nghìn); NG_TB; NG_GIA; TG_TRE / TG_SOM (> 60′ đỏ, bỏ qua lượt công tác); TG_CHO (> 2 ngày).
  - **theo người – ngày:** PV_SAI (chỉ sai vị trí → đỏ, ≥ 3 → vàng); TG_NHIEU (≥ 5); TG_THIEURA / TG_THIEUVAO; TG_NGAN; TG_NGHI (đơn cả ngày đã duyệt); TG_NGHILE (CN / lễ, trừ công tác).
  - **chéo:** PV_TOC (> 2 km và > 150 km/h); NG_TBCHUNG (1 thiết bị ≥ 2 người, mỗi người 1 cảnh báo, ≤ 3 lượt); TG_VANG (người có khuôn mặt, ngày làm việc từ ngày đăng ký, không đơn nghỉ / công tác).
- **API:**
  - **cckt.chay:** khử trùng theo khoa; không báo lại cảnh báo đã kết luận; tự đóng cảnh báo mở không còn xuất hiện; ≤ 92 ngày.
  - **cckt.ds:** READ_ONLY, kèm luot chi tiết và dots.
  - **cckt.nhac:** bao(qlBoPhan) loai duyet, mo #v=gpsql&t=cl; tùy chọn guiNV.
  - **cckt.giaitrinh:** 3 kết quả; báo người nhắc.
  - **cckt.xuly:** Đã xử lý / Bỏ qua / Mở lại (lý do, nhắc lại).
  - **cckt.ai:** READ_ONLY; gửi duLieu tổng hợp + 40 mẫu mô tả, không gửi ảnh; lọc boPhan / maNV không có trong số liệu.
  - **cckt.luuai.**
- **Liên kết với chức năng khác:**
  - load / login trả **soCB** cho huy hiệu menu (QL: Đã nhắc; kiểm tra: Đã giải trình);
  - gps.ds trả **cb** {luotId: {mucDo, ds}} để đánh dấu đỏ.
- **Client:**
  - tab 'kt' (ckLaKT), tab 'cl' (ROLE_QL);
  - CK state, thẻ cảnh báo .ck-the.do, bộ lọc, chọn hàng loạt, xem lượt (ảnh + bảng);
  - Excel 2 sheet;
  - bản dùng thử demoCCKT (ghi rõ không phải AI);
  - seed một đợt kiểm tra đã nhắc BCH công trình.
- **Kiểm thử:** t18 (quy tắc, AI, nhắc, giải trình, chạy lại), smoke24 (giao diện, phân quyền).


## 5y. v3.5 – Trợ lý quy định công ty (bot chính sách trong Tin nhắn)

**Mục tiêu:** CBCNV hỏi về chế độ, quy định, biểu mẫu, chính sách ngay trong màn hình Tin nhắn; trợ lý trả lời **chỉ dựa trên nội dung TCHC đưa vào**, tuyệt đối không tự sinh số liệu.

**Sheet mới** (thêm vào `SCHEMA`, `upgradeSchema()` tự tạo, không xóa dữ liệu cũ):
- `BotQuyDinh`: id, nhom, tieuDe, noiDung, lienKet, vanBan, hieuLuc, trangThai (hieuluc|het), nguoiSua, capNhat
- `BotVanBan`: id, ten, fileId, mime, nhom, soHieu, ngayBanHanh, trangThai, soChu, textLuc, loi, nguoiSua, capNhat
- `BotHoi`: id, thoiGian, username, hoTen, cauHoi, traLoi, nguon, nhom, daTraLoi, dungAI, thieu, huuIch, phanHoi

**Action:** `bot.ds`, `bot.hoi`, `bot.nhatky` (chỉ đọc — thêm vào `READ_ONLY`); `bot.ghi`, `bot.danhgia`, `bot.luu`, `bot.xoa`, `bot.vb.luu`, `bot.vb.xoa` (có khóa ghi). Quyền sửa nội dung: `BOT_ROLES_QD = ['admin','giamdoc','tchc']`. Mọi tài khoản đăng nhập đều hỏi được.

**Hằng số:** `BOT_NHOM` (9 nhóm chủ đề), `BOT_HE` (prompt hệ thống), `BOT_P` (yêu cầu trả JSON), `BOT_TMUC = '_QUY ĐỊNH CÔNG TY (KHÔNG CHIA SẺ)'`, `BOT_DUNG` (âm tiết không mang chủ đề), `BOT_PHU_MIN = 0.6`, `BOT_KHOP_MIN = 2`.

**So khớp (`botKhop`)**: bỏ dấu (`botKhongDau`), tách âm tiết, bỏ `BOT_DUNG`; khớp **âm tiết trọn vẹn** (+2) và **cụm 2 âm tiết** (+5); trả `{d, phu}` với `phu` = tỉ lệ âm tiết chủ đề của câu hỏi tìm thấy. **Không được so chuỗi con** — đó là lỗi đã gặp: "trợ" khớp trong "trong", "công" khớp trong "công ty", khiến trợ lý trả lời sai chủ đề.

**Đọc văn bản Drive (`botTrichChu` / `botNapChu`)**: Google Tài liệu → `DocumentApp.openById().getBody().getText()`; `text/*` → `env.luuTru.doc()`; PDF/Word → `UrlFetchApp` POST `drive/v3/files/{id}/copy` với `mimeType: application/vnd.google-apps.document` và `Authorization: Bearer ScriptApp.getOAuthToken()`, đọc xong `setTrashed(true)`. Chữ trích được lưu đệm thành `vb-<id>.txt` trong `<BOT_TMUC>/_trich-xuat`.

**Chống bịa (bắt buộc giữ nguyên khi làm lại app):**
1. `BOT_HE` khóa phạm vi vào phần NGUỒN, cấm mọi con số không có trong NGUỒN, cấm suy từ luật chung.
2. Lọc `nguon` AI trả về theo danh sách nguồn thực tế đã cấp (`dungQD` + `dsVBDung`), so sau khi bỏ dấu, chấp nhận khớp một phần ≥ 6 ký tự.
3. `daTraLoi = AI báo được && còn căn cứ && có chữ`.
4. `daTraLoi = false` ⇒ **thay toàn bộ** `traLoi` bằng câu chuẩn, `nguon = []`, giữ `thieu` để báo TCHC.

**Chưa có khóa AI:** tra cứu trực tiếp — khớp câu hỏi với `tieuDe + nhom`, yêu cầu `phu ≥ BOT_PHU_MIN` và `d ≥ 2 × BOT_KHOP_MIN`, trích nguyên văn `noiDung`, `dungAI = false`.

**Gom nội dung còn thiếu (`bot.nhatky`)**: Jaccard tập âm tiết ≥ 0,6, giữ mẫu dài nhất, tối đa 200 nhóm, trả 30 nhóm nhiều lần hỏi nhất.

**Giao diện (`client.html`)**: `BOT` (trạng thái hỏi đáp), `botItemHTML()` ghim đầu `chListHTML()`, `botHTML()` thay khung phải của `vChat` khi `CH.bot`, `botGui/botVe/botDanhGia/botHoiNhanh`; màn hình `botqd` (`vBotQD`, trạng thái `BQ`) với 3 thẻ: Mục quy định, Văn bản công ty, Câu hỏi của CBCNV. Màu TAPACO navy/cam, chữ ≥ 14,5 px, vùng chạm ≥ 40 px. Bản dùng thử: `demoBot(o)` chỉ trích lại nguồn trong prompt, cùng ngưỡng 0,6 — không tự sinh nội dung.

**Dữ liệu mẫu bản dùng thử:** 6 mục quy định mang nhãn "(mẫu)" và mở đầu "(NỘI DUNG MẪU – phòng TCHC thay bằng quy chế thật của Công ty)", chỉ mô tả **thủ tục và chỗ tra cứu**, không đặt ra con số chế độ nào.

**Kiểm thử:** `test/t23_bot_gas.js` (40 phép thử trên Code.gs thật, gồm hồi quy chống bịa: AI trả lời không căn cứ phải bị loại hẳn, câu khác chủ đề không được trả lời bằng quy định không liên quan) và `test/smoke27.js` (giao diện 390×844).


## 5z. v3.5.1 – Bảng tin tự giãn theo màn hình, phân tích AI đọc báo cáo ngày

**Bảng tin**: `.bt-wrap{grid-template-columns:minmax(0,1fr) clamp(300px,21vw,380px);max-width:1800px;margin:0 auto;width:100%}`; danh sách bài viết bọc trong `.bt-feed`, `@media (min-width:1560px)` chia 2 cột. Bỏ `max-width:1180px` của bản cũ.

**Phân tích nhóm đọc báo cáo ngày** (bắt buộc giữ khi làm lại app):
- `ltNhatKy` đặt khối báo cáo định kỳ TRƯỚC phần diễn biến; tiêu đề khối ghi rõ “nguồn ưu tiên”.
- `ltBCNgay` ghi dòng người nộp dạng `**Họ tên** (tài_khoản, nộp lúc hh:mm)` để quy việc về đúng username.
- `ltRutGon(t, max)`: cắt ở mốc `\n## Diễn biến trong ngày`, giữ nguyên phần đầu (thống kê + báo cáo), chỉ lược phần trò chuyện, cắt theo ranh giới dòng. `aiPhanDot` dùng hàm này thay cho `t.slice(0, 45000)`.
- `AI_P_DOT` bổ sung: bắt buộc đọc khối báo cáo; bảng ánh xạ tên trường mẫu báo cáo sang `nhanLuc`, `thiCong`, `thietBi`, `vatTuNhap/Xuat`, `yeuCauVatTu`, `kienNghi`, `tonTai`, `keHoachMai`, `thoiTiet`, `anToan`, `suCo`; ưu tiên số liệu báo cáo hơn tin nhắn; cộng nhân lực khi nhiều người cùng báo cáo một ngày; quy việc về đúng tài khoản.
- Kiểm thử: `test/smoke28.js` — bơm 340 tin nhắn cho nhật ký vượt 45.000 ký tự rồi khẳng định khối báo cáo vẫn còn nguyên (nhân lực 18, khối lượng 42 m³, thiết bị), và bảng tin phủ hết bề ngang ở 1920/1600/1440/390/2560.


## 5aa. v3.5.2 – Khối lượng từ báo cáo cá nhân, vật tư tách cột nhập/xuất

**Bóc số liệu xác định từ báo cáo** (`src/core.js`, đặt trước `ltBCNgay`):
- `BC_DVT` danh sách đơn vị xây dựng; `BC_RE_CUOI` (số + đơn vị ở cuối câu), `BC_RE_GIUA` (số + đơn vị giữa câu).
- `bcBocHM(txt)` tách ô văn bản ở `;` và xuống dòng (KHÔNG tách ở dấu phẩy — câu tiếng Việt dùng phẩy nhiều), mỗi phần cho `{hangMuc, khoiLuong, dvt, trich}`. **Chỉ nhận số khi có đơn vị hợp lệ** — nếu không “đài móng T4” thành khối lượng 4, “dầm D3 60%” thành 3.
- `bcSoLieu(mau, gv)` khớp tên trường (`BC_TR`) sang nhân lực, ban chỉ huy, khối lượng, thiết bị, vật tư nhập/xuất, đề nghị cấp, kế hoạch.
- `bcNgaySo(ctx, ngay)` gom số liệu máy đọc của cả ngày theo người nộp.
- `ltBCNgay` in thêm dòng `[số liệu máy đọc được, dùng làm chuẩn]` dưới mỗi lần nộp.
- `AI_P_DOT` thêm ví dụ tách khối lượng và yêu cầu dùng đúng dòng số liệu máy đọc.

**Lưới an toàn `aiBuBC(kq, ctx, d, tv)`** gọi trong `aiPhanDot` sau `aiSachDot`: với mỗi ngày, nếu `thiCong` / `thietBi` / `vatTuNhap` / `vatTuXuat` / `yeuCauVatTu` **rỗng** mà báo cáo có dữ liệu thì điền từ số liệu máy bóc (`ghiChu: 'lấy từ báo cáo định kỳ'`); `nhanLuc.tong` null thì cộng từ báo cáo; `caNhan.dongGop` rỗng thì điền việc của chính người nộp. Chỉ bù khi AI để trống hoàn toàn — không trộn vào kết quả AI đã có, tránh nhân đôi số liệu.

**Vật tư tách cột**: `aiGopVT(ds)` gom theo `(tên, ĐVT)` không phân biệt hoa thường, trả `{ten, dvt, nhap, xuat, deNghi, coNhap, coXuat, coDeNghi, chenh, lan, ngay[]}`; `chenh` chỉ tính khi `coNhap && coXuat`. Thay cho `aiGop(... '|' + hoatDong ...)` cũ. Cập nhật đồng bộ 4 nơi: bảng Word/Docs, 2 bảng Excel (8 cột, `cols` phải khớp), bảng tóm tắt chữ `aiTongHopTxt`, và bảng trên màn hình trong `client.html`.

**Kiểm thử**: `test/smoke29.js` — 18 phép thử, gồm trường hợp giả lập AI trả về rỗng hoàn toàn mà khối lượng 42 m³, nhân lực 18 và việc của baopq vẫn ra đủ; và trường hợp “Lắp dựng cốt thép đài móng T4” không bị bịa thành khối lượng 4.


## 5ab. v3.5.3 – Tự lưu file kết quả phân tích lên Drive

- `src/gas.js` – `luuTruEnv_.ghiNhiPhan(chaId, ten, data, mime)`: `Utilities.newBlob(Utilities.base64Decode(data), mime, ten)` + `createFile`; trùng tên thì `setTrashed(true)` bản cũ rồi tạo bản mới (DriveApp không thay được nội dung nhị phân tại chỗ).
- `src/core.js` – `H['ai.luutep']` (có khóa ghi): `ltQuyen` + `ltCay` như các action `ai.*`; chặn đuôi ngoài `.xlsx/.docx/.pdf`, nội dung rỗng và base64 quá 24 triệu ký tự (~18 MB); ghi vào `c.ai`; nếu có `jsonId` thì gắn `xlsxId` / `docxId` / `pdfTaiId` vào JSON của đợt (bọc try/catch, lỗi thì bỏ qua).
- `src/client.html` – thêm bước `luutep` cuối `aiChay`: `XLSX.write(C.sangXLSX(C.xlAiBaoCao(aiChuan(R.kq)), XLSX), {type:'base64', bookType:'xlsx'})` rồi gọi `ai.luutep`. Bọc try/catch — hỏng thì chỉ ghi chú ở dòng tiến trình. Bản dùng thử: `demoDrive.ghiNhiPhan` mô phỏng. Đầu báo cáo có dòng “Đã lưu trên Drive: …”.
- Kiểm thử: `test/t24_gas.js` (12 phép trên Code.gs thật — tệp vào đúng thư mục, ghi đè không sinh rác, chặn đuôi lạ / tệp lớn / rỗng / người ngoài nhóm) và `test/smoke30.js` (12 phép giao diện — có bước lưu, gắn mã tệp vào JSON, mở lại báo cáo cũ vẫn thấy, và trường hợp thiếu thư viện Excel vẫn không hỏng lượt phân tích).

### 5ac. Công văn đến – đi và Đề xuất nội bộ (v3.6)

**Sheet:** `CongVan` (sổ đến/đi: `huong` den|di, `nam`, `so`, `soHieu`, `ngayVB`, `ngayVaoSo`, `noiGui`, `noiNhan`, `loai`, `doKhan`, `trichYeu`, `nguoiKy`, `soBan`, `fileId`/`tenFile`/`mime`/`kichThuoc`, `butPhe`, `nguoiXuLy`, `phoiHop` (JSON mã người), `hanXuLy`, `trangThai`, `ketQua`, `hoanThanhLuc`, `lienKet`, `ghiChu`, `nguoiNhap`, `taoLuc`, `capNhat`); `CongVanXL` (vết xử lý: `cvId`, `hanhDong`, `nguoi`, `tenNguoi`, `noiDung`, `hanXuLy`, `fileId`, `tenTep`, `thoiGian`); `DeXuat` (`ma` DX-<năm>-<nnn>, `loai`/`tenLoai`, `tieuDe`, `noiDung`, `soTien`, `donVi`, `maNV`/`hoTen`/`boPhan`, `doKhan`, `hanCan`, `trangThai`, `capDuyet`, `luongDuyet` (ảnh chụp các cấp duyệt), `theoDoi`, `thucHien`, `lichSu`, `noiDungId`, `ketQua`, `hoanThanhLuc`, `nguoiDuyet`, `ngayDuyet`, `ghiChuDuyet`, `nguoiTao`, `taoLuc`, `capNhat`); `DeXuatTep` (`dxId`, `loai` anh|tep, `ten`, `mime`, `kichThuoc`, `fileId`, `nguoiTai`, `taiLuc`).

**Quyền mới:** `congvan` (admin, giamdoc, tchc – xem toàn bộ sổ), `vanthu` (admin, tchc – đăng ký vào sổ, tải bản quét, lưu hồ sơ, xóa).

**Thư mục Drive:** `env.files.put` nhận thêm `duong: [...]` để ghi vào thư mục nhiều cấp dưới thư mục gốc riêng tư; thêm `env.files.ghiVB(duong, ten, nd)` (tệp văn bản, trùng tên thì ghi đè) và `env.files.duongUrl(duong)`. Đường dùng: `['_CÔNG VĂN', 'Đến'|'Đi', nam]` và `['_ĐỀ XUẤT NỘI BỘ', nam, ma]`.

**API công văn:** `cv.ds` (chỉ đọc – trả danh sách theo quyền, vết xử lý, số đến/đi kế tiếp, cờ `vanThu`/`giaoDuoc`/`toanBo`, link thư mục Drive) · `cv.save` (đăng ký / sửa; số tự tăng theo `huong`+`nam`, chặn trùng số) · `cv.upload` (bản quét; công văn đi chuyển "Đã ban hành") · `cv.get` (chỉ đọc; `xlId` để lấy tệp kèm một lần báo cáo) · `cv.giao` (bút phê + người xử lý + phối hợp + hạn, chặn hạn đã qua) · `cv.baocao` (`xong` → "Đã xử lý", kèm tệp) · `cv.chot` (Lưu hồ sơ – khóa) · `cv.delete` (chỉ khi còn mới đăng ký).

**API đề xuất:** `dx.ds` (chỉ đọc – đề xuất mình liên quan + danh mục loại) · `dx.create` (`nhap: true` để lưu nháp) · `dx.update` (`trinh: true` để trình duyệt; chỉ người lập, khi chưa cấp nào duyệt) · `dx.nguoi` (đổi người theo dõi / thực hiện) · `dx.upload` / `dx.tepget` (chỉ đọc) / `dx.tepdel` · `dx.decide` (duyệt từng cấp, Giám đốc & Quản trị vượt cấp, từ chối bắt buộc lý do) · `dx.ketqua` (người thực hiện cập nhật / báo hoàn thành) · `dx.huy`.

**Nhóm, loại, bộ phận:** `cfg.nhomDeXuat` = `[{id, ten}]` (Quản trị tạo). `cfg.loaiDeXuat` = `[{id, nhom, tienTo, ten, tien, cach: 'trinhtu'|'motnguoi', gdCuoi: true, hoatDong, buoc: [...], theoDoi: [], thucHien: []}]`. **Bộ phận đề xuất** là trường người lập chọn, kiểm tra ngược lại danh mục `cfg.boPhan` + `cfg.phongBan` + phòng ban/đơn vị đang có trong hồ sơ CBCNV; mặc định là bộ phận của người lập. `DeXuat` lưu `nhom`, `tenNhom`, `boPhan`, `tienTo`, `so`.

**Mã đề xuất:** `<tienTo><số 4 chữ số>` chạy riêng theo tiền tố (`TT0147`, `TUL0032`); không có tiền tố thì chỉ là số thứ tự. Tiền tố 2–8 chữ in hoa không dấu, không trùng giữa các loại — `kiemTraLoaiDX()` chặn.

**Giám đốc là cấp duyệt cuối (bất biến):** `dxCapDuyet()` **luôn nối thêm** cấp `{id:'gd', ten:'Giám đốc phê duyệt', nguoi:['r:giamdoc']}` vào cuối chuỗi; cấu hình không khai và không bỏ được cấp này, trình sửa hiện nó ở cuối dưới dạng chỉ đọc. `cach: 'motnguoi'` gộp toàn bộ người duyệt của loại (kể cả Giám đốc) thành **một** cấp.

**Phân quyền cấu hình:** nhóm và loại đề xuất **chỉ Quản trị** sửa (`quytrinh.save` bỏ qua `nhomDeXuat`/`loaiDeXuat` nếu người gửi không phải admin); Giám đốc vẫn vào màn hình để xem.

**Trình tự duyệt theo loại:** `cfg.loaiDeXuat` = `[{id, ten, tien, buoc: [{id, ten, nguoi: ['r:vaiTro'|'u:tenDN'|'tbp'], tren}], theoDoi: [], thucHien: []}]`. `dxCapDuyet()` chọn các cấp thỏa `!tren || soTien > tren`, quy `tbp` về trưởng bộ phận của người đề nghị, **bỏ cấp mà người duyệt chính là người nộp**, **gộp hai cấp liền nhau cùng người**, rồi chụp vào `DeXuat.luongDuyet` — giống hệt cách `capDuyetCho()` làm với đơn nghỉ, nên đề xuất đã trình không bị lệch khi sửa cấu hình. Lưu cấu hình qua `quytrinh.save {loaiDeXuat}`; kiểm tra bằng `kiemTraLoaiDX()` + `kiemNguoi()`.

**Ghi nội dung đề xuất ra Drive:** `dxLuuDrive(o)` dựng lại tệp `NoiDung_<mã>.txt` (thông tin chung, nội dung, trình tự duyệt, người theo dõi/thực hiện, diễn biến, kết quả) sau **mỗi** lần tạo / sửa / duyệt / báo kết quả / hủy. Bọc `try/catch`: lỗi Drive không được làm hỏng thao tác nghiệp vụ.

**Dữ liệu khởi động vẫn nhẹ:** `loadAll` chỉ trả `soCV`, `soDX` (số việc chờ, dùng cho chấm đỏ) và `coCV` (có liên quan công văn nào không, dùng để hiện mục trong menu); danh sách đầy đủ tải riêng bằng `cv.ds` / `dx.ds`.

**Giao diện:** hai mục `cv`, `dx` trong `VIEWS`; trạng thái `CVS`, `DXS` tải chậm (lazy) và **phải xóa trong `forceLogout()`** — nếu không, người đăng nhập sau nhìn thấy sổ của người trước. Màn hình "Quy trình & phê duyệt" có mục 4 cấu hình loại đề xuất (sửa lồng: loại → cấp duyệt → người duyệt).

**Nguyên tắc dữ liệu mẫu:** loại đề xuất dựng sẵn chỉ đặt **tên loại và trình tự**, mọi ngưỡng tiền để 0 — phần mềm không tự đặt ra mức tiền hay thẩm quyền nào của Công ty.

### 5ad. AI đọc công văn, Giám đốc duyệt chia sẻ, trao đổi dưới văn bản (v3.7)

**Sheet:** `CongVan` thêm cột `hanTraLoi, tomTat, aiLuc, aiTinCay, csTrangThai, csDeXuat, nguoiDoc, nguoiDuyetCS, csLuc, csGhiChu, soTN` (soTN là cột số). Sheet mới `CongVanTN` (id, cvId, username, hoTen, noiDung, fileId, tenTep, mime, thoiGian). `upgradeSchema()` tự thêm.

**Hằng số:** `CV_CS = ['Chưa trình','Chờ duyệt','Đã chia sẻ','Không chia sẻ']`, `CV_CS_OK = 'Đã chia sẻ'`.

**AI đọc văn bản:** `cv.ai` (quyền `vanthu`, READ_ONLY vì gọi AI chậm) nhận `{huong, file}`; kiểm MIME + dung lượng; ảnh gửi thẳng `aiGoi({anh})`, PDF/Word ghi tạm vào `_CÔNG VĂN/_tam-doc` → `botTrichChu` → xóa tạm. Prompt `CV_HE`: **chỉ trích những gì có trong văn bản, không thấy thì chuỗi rỗng**, ngày `yyyy-MM-dd`, trả JSON `{soHieu, ngayVB, noiGui, noiNhan, loai, doKhan, trichYeu, nguoiKy, hanTraLoi, tomTat, doTinCay}`. Máy chủ **lọc lại** mọi trường: ngày qua `cvNgay`, `loai`/`doKhan` chỉ nhận giá trị trong `CV_LOAI`/`CV_KHAN`, cắt độ dài; ghi nhật ký. Client `cvAIDoc()` chạy khi chọn/thả tệp, **chỉ điền ô trống** (ô chọn còn mặc định thì điền; ô văn thư đã chọn đánh dấu `data-doi`), tô `.aifill`, hiện độ tin cậy. `cv.save` lưu `hanTraLoi, tomTat, aiTinCay` và khởi tạo `csTrangThai:'Chưa trình', csDeXuat:'[]', nguoiDoc:'[]', soTN:0`.

**Duyệt chia sẻ:** `cv.trinhcs` (văn thư/GĐ/QT; `nguoiDoc` mã người dạng `r:|u:|tbp`, `note`) → `Chờ duyệt`, báo giamdoc. `cv.duyetcs` (chỉ giamdoc/admin): `duyet:true` (được sửa danh sách) → `Đã chia sẻ`, ghi `nguoiDoc, nguoiDuyetCS, csLuc`, báo người được đọc; `duyet:false` bắt buộc `note` → `Không chia sẻ`. `cvXem(s,o)` thêm điều kiện `laNguoiXuLy(s, cvNguoiDoc(o))` với `cvNguoiDoc` **chỉ trả danh sách khi `csTrangThai === CV_CS_OK`**. `demCV` đếm thêm văn bản chờ GĐ duyệt chia sẻ.

**Trao đổi:** `cv.tn` (READ_ONLY, cần `cvXem`) trả lời trao đổi của văn bản; `cv.gui` (`noiDung`, `file` tùy chọn lưu qua `cvTepVao(o, file, 'TD')`) – khóa sau `Lưu hồ sơ` trừ admin, tăng `soTN`, báo người liên quan. `cv.get` nhận `tnId` để mở tệp kèm trao đổi. `cv.delete` xóa luôn dòng `CongVanTN` và tệp.

**Giao diện:** form vào sổ đặt ô tệp + `#cvAIst` lên đầu, thêm `f_hanTraLoi`, `#cvTomTat`; chi tiết hiện hạn trả lời, nhãn chia sẻ, tóm tắt AI, mục "Trao đổi xử lý văn bản" (`CVT`, `cvTNTai/cvTNVe/cvTNGui/cvTNTep`); nút "Trình chia sẻ / Chia sẻ / Duyệt chia sẻ" (`cvCSForm`, `cvCSTuChoi`). Bản dùng thử: `DEMO_AI.goi` với `loai:'congvan'` trả dữ liệu **ghi rõ "NỘI DUNG MẪU"**.

**Kiểm thử:** `test/t26_gas.js` (AI không bịa trên bản mờ, quyền, MIME, xóa tệp tạm; chia sẻ trước/sau duyệt, sửa danh sách, từ chối có lý do; trao đổi: quyền, tệp vào `_CÔNG VĂN`, đếm, khóa, dọn khi xóa) và `test/smoke32.js` (giao diện).

### 5ae. Thu hẹp phần quản lý: Phân tích nhóm (AI) & Chấm công GPS (v3.7.1)

**Quyền mới** `PERM.giamsat = ['admin','giamdoc','tchc','ktluong','ktcongno']`. Trưởng/phó nhóm chat, Trưởng phòng, Chỉ huy trưởng **không** còn quyền quản lý ở hai phần này.
- `ltQuyen` (mọi `ai.*`, `chat.luutru`, `chat.nhatky`): cần `giamsat`; trò chuyện riêng chỉ admin/giamdoc. `ai.nhom` cần `giamsat`, trả mọi nhóm (trò chuyện riêng chỉ admin/giamdoc). `ai.cauhinh`/`ai.thu` vẫn chỉ admin.
- `gpsQuanLy` (gps.ds, gps.anh của người khác, gps.duyet, gps.ct.ds/duyet/huy, khai báo công tác thay): chỉ `giamsat`, phạm vi toàn công ty. `gps.diem`: edit_nv + giamdoc (bỏ trưởng bộ phận).
- Đăng ký công tác báo `nguoiGiamSatCC()` = các tài khoản tchc, ktluong, ktcongno (không báo TP/CHT).
- `cckt.ds`, `cckt.giaitrinh` chỉ `CCKT_ROLES`; `soCanhBao` chỉ cho CCKT_ROLES. `cckt.nhac` báo `nguoiGiamSatCC` (+ người lao động nếu `guiNV`), không báo TP/CHT; `Mở lại` không còn báo bộ phận. `cckt.giaitrinh` nhận `dong:true` → ghi `ketQua`, `giaiTrinh` và chuyển thẳng `Đã xử lý` (nút "✍ Ghi kết quả & đóng" ở tab Kiểm tra AI).
- Giao diện: `allowed('gpsql')` và `aiCo()` = `can('giamsat')`; bỏ tab "Cần kiểm tra lại" (liên kết cũ `t=cl` chuyển sang `kt`); nút "Báo cáo AI" trong Thông tin nhóm chỉ hiện khi `aiCo()`.
- Kiểm thử: `test/t27.js` (ma trận quyền 9 vai trò, thông báo công tác/nhắc, kết luận cảnh báo); cập nhật t13, t14, t15, t16, t18, smoke20–24, smoke28.

### 5af. Bảng phân quyền tích chọn được (v3.8)

- `PERM` giữ nguyên đối tượng, thêm quyền `aiNhom`, `ccQuanLy` (thay `giamsat`), `danhgia` (thay DGT_ROLES), `botqd` (thay BOT_ROLES_QD); `CCKT_ROLES`/`DGT_ROLES`/`BOT_ROLES_QD` thành đối tượng `{indexOf}` gọi `can()`. `PERM_MD` = bản sao mặc định.
- `PQ_NHOM` (4 nhóm, 18 quyền có nhãn), `PQ_SUA`, `PQ_KHOA = {users:[admin], config:[admin], congthuc:[admin], quytrinh:[admin,giamdoc], view_all:[admin,giamdoc], approve:[giamdoc]}`, `PQ_NGUOI_SUA = [admin, giamdoc]`.
- `apDungPQ(pq)` dựng lại `PERM[k]` từ `cfg.phanQuyen` (chỉ khóa trong PQ_SUA, vai trò hợp lệ) + ô khóa; có bộ nhớ đệm theo JSON. `pqLamSach(pq)` chuẩn hóa để lưu / so sánh. `getCfg()` luôn gọi `apDungPQ`; `handle()`, `chayLich()`, `luuTruTuDong()` gọi `getCfg()` trước khi kiểm tra quyền.
- `pq.save {pq}` hoặc `{macDinh:true}` – chỉ PQ_NGUOI_SUA; ghi `cfg.phanQuyen`, nhật ký liệt kê cấp/bỏ; trả `{pq, doi}`. `cfg.save` bỏ qua khóa `phanQuyen`. `DEFAULT_CFG.phanQuyen = {}`; cấu hình gửi cho mọi người dùng có `phanQuyen` để app ẩn/hiện menu.
- Client: `can()` gọi `pqDongBo()` (so tham chiếu `APP.D.cfg.phanQuyen`, gọi `C.apDungPQ`). `pqHTML()` ở trang Tài khoản: ô tích (`data-pq="k|r"`), ô khóa disabled 🔒, ô đổi `.pq-doi`, thanh `#pqBar`/`#pqBar2` (Về mặc định, Hoàn tác, Lưu), `pqLuuForm()` xác nhận + cảnh báo quyền nhạy cảm (`PQ_NHAY`) cho TP/CHT/NV; `PQ_TT` các dòng chỉ xem (theo quy trình / cố định).
- Kiểm thử: `test/t28.js` (quyền sửa, áp dụng ngay, lưu bền khi khởi động lại, ô khóa, lọc vai trò/quyền lạ, cfg.save không ghi đè, về mặc định), `test/smoke33.js` (giao diện).

### 5ag. Mục "Bảng số lượng" trong yêu cầu báo cáo (v3.9)

- Kiểu mục mới `bang` (server `BC_KIEU`, client `BC_KIEU.bang`). Cấu hình mục: `dong: [{ten, dv}]` (≤60, bỏ trùng tên không phân biệt hoa thường, bỏ ":" và xuống dòng), `them` (cho người báo cáo tự thêm dòng; mặc định true). Không có dòng và `them=false` → lỗi.
- Client `BC_GOIY` 3 nhóm (Nhân lực – người, Máy thi công – cái, Công việc thực hiện – đơn vị từng dòng); mẫu mặc định mới đứng đầu `BC_MAU`: "Báo cáo công trường (nhân lực, máy, công việc)" (3 mục bang + diễn giải, thời tiết, sự cố, vướng mắc, kế hoạch). Trình chọn trong `bcHangHTML`: chip + `.bc-mot` (chọn lẻ, ẩn mục đã chọn) + "Cả nhóm" + dòng tự gõ; `bcKieuDoi` ẩn/hiện.
- Nộp (`chat.nopbc`): giá trị gửi lên là mảng `[{ten, so, dv}]` hoặc chữ; `bcBangSach` chuẩn hóa số kiểu VN, chặn số âm/không hợp lệ, bỏ dòng trống/0, gộp trùng, lưu **chữ** "Tên: số đv" mỗi dòng + "Tổng: …" chỉ khi mọi dòng cùng đơn vị đếm được (`BC_DV_DEM`: người, cái, chiếc, xe, máy, bộ, tổ, đội, công). `HRCORE.bcBangDoc(txt)` đọc lại (bỏ dòng Tổng) – client dùng để mở sửa.
- `bcSoLieu`: mục bang → `nhanLuc` = tổng (+ `nhanLucCT` chi tiết), thiết bị / vật tư / đề nghị / khối lượng lấy thẳng từng dòng {hangMuc, khoiLuong, dvt, trich}; dòng nhật ký AI "nhân lực N người (NC … n, …)". Tổng hợp kỳ cộng dồn từng dòng: `so` thêm mục "<tên mục> – <dòng> (<đv>)".
- Giao diện điền `bcOHTML` kiểu bang: lưới tên | ô số | đơn vị, dòng tự thêm (`.bcb-them`), tổng tạm (`bcBangTong`); bảng tổng hợp đổi `\n` thành `<br>`.
- Kiểm thử: `test/t29.js`, `test/smoke34.js`.

### 5ah. Sửa lỗi chọn tệp trong menu Tin nhắn (v3.9.1)

- Lỗi: menu (…) của khung chat chứa luôn 2 `<input type=file>` ẩn; nút "Ảnh"/"Tệp" chạy `$('chPop').remove()` **trước** rồi mới `.click()` ô chọn tệp → ô đã bị xóa cùng menu, lỗi `null.click`, không mở hộp chọn tệp.
- Sửa: `chChonTep('anh'|'tep')` tạo một lần các ô `#chImgIn`, `#chFileIn` gắn thẳng vào `body` (ngoài mọi menu / hộp thoại), làm rỗng giá trị trước mỗi lần bấm để chọn lại được cùng tệp. **Quy tắc:** không đặt ô chọn tệp trong phần tử sẽ bị xóa trước khi mở hộp chọn tệp.
- Kiểm thử: `test/smoke35.js` (máy tính + điện thoại: Ảnh, Tệp mở hộp chọn tệp và gửi được, chọn lần 2, nút ảnh cạnh ô nhập, Kho tệp).

### 5ai. Ảnh CCCD trong lý lịch (v3.10)

- Không thêm sheet: dùng `TaiLieu` loại "Căn cước công dân", `ten` = "CCCD – mặt trước" | "CCCD – mặt sau" (`CCCD_TEN`).
- `tl.upload` nhận thêm `cccdMat: 'truoc'|'sau'`: chỉ nhận ảnh JPG/PNG/WEBP; tự đặt loai, ten, soHieu = `nv.cccd`, ngay = `nv.ngayCapCCCD`; quyền như tài liệu (edit_nv cho mọi người, CBCNV cho mình nếu loại nằm trong `loaiNVTuTai` → "Chờ xác nhận"). Sau khi lưu, xóa bản cũ cùng mặt mà người tải **được phép xóa** (edit_nv: mọi bản; CBCNV: bản của mình đang chờ xác nhận), xóa luôn file; trả `xoa: [id]`; nhật ký ghi "thay N ảnh cũ".
- Client: `cccdHTML(ma, self)` trong tab Thông tin (2 ô tỉ lệ thẻ 85,6×54), `cccdTai` (btNen → tl.upload), `cccdHydrate()` tải ảnh qua `tl.get` sau render (đệm `CCCD_ANH`, xóa khi `forceLogout`); bấm ảnh → `viewTL`.
- Kiểm thử: `test/t30.js`, `test/smoke36.js` (ảnh mẫu `test/cccd_truoc.png`, `cccd_sau.png` là ảnh thử nghiệm tự vẽ, không phải giấy tờ thật).

### 5aj. AI đọc ảnh CCCD → đề xuất điền lý lịch (v3.11)

- `nv.cccdai {maNV}` (edit_nv, READ_ONLY vì gọi AI chậm): lấy ảnh mới nhất mỗi mặt (`cccdAnhHS`), gửi `aiGoi({he: CCCD_HE, nd: '=== ĐỌC CCCD ===…', anh: [{nhan:'Ảnh MẶT TRƯỚC thẻ:'…}, {…MẶT SAU…}], loai:'cccd'})`. `CCCD_HE`: chỉ chép chữ in trên thẻ đọc rõ, không đọc rõ → "", không suy đoán kể cả tên cơ quan cấp thường gặp; JSON {hoTen, soCCCD, ngaySinh, gioiTinh, queQuan, noiThuongTru, ngayCap, noiCap, coGiaTriDen, doTinCay, ghiChu}.
- Làm sạch phía máy chủ: số chỉ nhận đúng 12 chữ số; ngày `yyyy-MM-dd` hợp lệ, ngày sinh / ngày cấp không ở tương lai; giới tính Nam|Nữ. `cccdKiem` đối chiếu chéo (chữ số 4 = giới tính + thế kỷ, 5–6 = năm sinh; ngày cấp ≤ ngày sinh; hết hạn; họ tên khác hồ sơ so không dấu) → `canhBao[]`, không tự sửa. Trả `{ai, hienTai, canhBao, soO, soAnh, doTinCay, ghiChu, demo}`; nhật ký chỉ ghi số ô / số cảnh báo.
- `nv.cccdluu {maNV, gt}` (edit_nv): chỉ nhận khóa trong `CCCD_TRUONG` (cccd, ngaySinh, gioiTinh, ngayCapCCCD, noiCapCCCD, hetHanCCCD, queQuan, diaChi); kiểm lại định dạng; `hetHanCCCD` ghi vào `TaiLieu.hetHan` các bản CCCD; cập nhật `soHieu`/`ngay` của ảnh CCCD theo số / ngày cấp mới; nhật ký ghi tên ô đã đổi.
- Client: nút trong `cccdHTML` (chỉ edit_nv, khi đã có ảnh); `cccdAI(ma)` mở hộp soát (ảnh + bảng `.cai`), tích sẵn chỉ ô lý lịch đang trống, ô khác lý lịch `.khac` không tích; gõ vào ô thì tự tích. `DEMO_AI.goi` loại `cccd`: trả lại dữ liệu hồ sơ, gắn `demo:true`, độ tin cậy thấp, giao diện ghi rõ "dữ liệu MẪU".
- Kiểm thử: `test/t31.js` (AI giả lập đọc chữ trong "ảnh"), `test/t32_gas.js` (trên Code.gs thật + Gemini giả lập), `test/smoke37.js`.

### 5ak. Thuế TNCN & BH với tiền làm thêm giờ (v3.12)

- Căn cứ: Luật Thuế TNCN 109/2025/QH15 khoản 8 Điều 4; NĐ 253/2026/NĐ-CP Điều 26, Điều 69 (áp dụng từ kỳ tính thuế 2026 với cá nhân cư trú); BLLĐ 2019 Điều 107; Luật BHXH 2024 Điều 31, NĐ 158/2025.
- Biến hệ thống mới trong `taoBien`: `OT_GIO`, `OT_GIO_LUYKE` (từ `o.otLuyKe` = tổng giờ OT các kỳ trước cùng năm, hàm `otLuyKe(maNV, ky, ccAll)`; `dauVaoTinhLuong` nhận tham số thứ 7 `ccAll`), `OT_GIO_MIEN = max(0, min(OT_GIO, otGioiHanThang, otGioiHanNam − OT_GIO_LUYKE))`, `OT_TL_MIEN` (kỳ ≥ 2026-01: `OT_GIO_MIEN / OT_GIO`; kỳ cũ: phần hệ số vượt 100%). Cấu hình `otGioiHanThang` (40), `otGioiHanNam` (200).
- Mẫu chuẩn: `TANGCA.thueGT = 'GT * (1 - OT_TL_MIEN)'`. Mẫu TAPACO: `CONG_TL` thử việc = `NGAY_CONG`; thêm trung gian `OT_QD`, `TONG_OT`; khoản `TANGCA = TONG_OT − TONG_TG` (phân loại tăng ca, thuế công thức như trên) → tổng không đổi.
- `otVaCT(ds)` (xuất ra client) nhận diện & vá bộ công thức đang lưu; API `congthuc.capnhatot` (`xem:true` để xem trước; `tinhLaiKyMo`) — kiểm `kiemTraCT`, ghi nhật ký. Client: khung `otCanhBaoHTML()` ở trang Công thức lương, Cài đặt thêm 2 ô giới hạn giờ.
- Kiểm thử: `test/t33.js` (tổng không đổi, thuế, giới hạn tháng / năm, kỳ 2025, lương đóng BH không đổi, nâng cấp cấu hình cũ, quyền), `test/smoke38.js`; `test/t4.js` cập nhật theo cấu trúc mới.

### 5al. Loại lương theo hợp đồng (v3.13)

- Hằng số `LOAI_LUONG` {bh, khongbh, khoan}, `LY_DO_KHONG_BH` ['Đóng BH nơi khác','Đã nghỉ hưu'], hàm `loaiLuongNV(nv)` (xuất ra client). Cột mới: `NhanVien.loaiLuong`, `NhanVien.lyDoKhongBH`, `ChamCong.tienKhoan` (FIELDS_LOAI.tinhluong, NUM_COLS).
- `tinhPhieuLuong`: khoan → dùng `CT_KHOAN` (KHOAN = IF(TIEN_KHOAN > 0, TIEN_KHOAN, LUONG_CHINH); THUONGKT; TRUYLINH) thay bộ công thức; `laVangLai` gồm khoan. khongbh → NLĐ 0; DN: nơi khác = lbh × tyLeDN.tnld (0,5) vào bhxhDN, nghỉ hưu = 0. ghiChu giải thích từng loại, cảnh báo khoán + HĐLĐ có thời hạn.
- Biến: `TIEN_KHOAN`, `LA_KHOAN`, `KHONG_BH`. DEFAULT_CFG `nguongVangLai` 5.000.000, `tyLeDN.tnld` 0,5.
- `nv.save`: nhận mã hoặc tên loại (nhập Excel), bắt buộc lý do khi khongbh, đồng bộ `thamGiaBH`.
- `xlBangTinh`: cột "Loại lương"; cột "Tiền khoán việc" khi có người khoán; dòng loại 2/3 ghi giá trị BH, khoản thu nhập dạng số (không công thức).
- Giao diện: hồ sơ (ô Loại lương 1/2/3, ô lý do ẩn/hiện `llDoi()`, nhãn "Mức khoán tháng", cho phép 0), lưới bước Tính lương cột "Tiền khoán việc" (chỉ hiện khi có người khoán, chỉ mở ô người khoán), phiếu lương, Cài đặt `dn_tnld`.
- Kiểm thử: `test/t34.js`, `test/t11_xl.js` (thêm 3 người, LibreOffice tính lại 0 lệch), `test/smoke39.js`.

### 5am. Quy chế ăn ca (v3.13.1)

- Biến `NGAY_DU_CA` trong `taoBien`: đếm ô `chiTietNgay` có `KY_HIEU[k].cong >= 1` (X, CT); không có bảng theo ngày → `floor(NGAY_CONG)`. Mẫu TAPACO `ANCA = IF(NGAY_AN > 0, NGAY_AN, NGAY_DU_CA) * TIEN_AN_NGAY`; mức theo `noiLamViec[].anCa` (VP 30.000, DA 70.000), mức riêng `nv.anCa` ưu tiên.
- `anCaVaCT(ds, bienKy)` vá công thức cũ (`IF(NGAY_AN>0,NGAY_AN,NGAY_CONG)*TIEN_AN_NGAY` hoặc `NGAY_CONG*TIEN_AN_NGAY`); API `congthuc.capnhatanca` (`xem`, `tinhLaiKyMo`); giao diện `anCaCanhBaoHTML()` trên trang Công thức lương.
- Kiểm thử: `test/t35.js`, `test/smoke40.js`.
- v3.13.2: `gps.gioLamDuCong` 8, `gps.nghiTruaTu/nghiTruaDen` '12:00'/'13:30'; `gioLamThuc(vao, ra, g)` trừ phần trùng nghỉ trưa, `nguongDuCong(g)` = gioLamDuCong − phutTre/60; dùng ở `gps.tonghop` (X / X2), kiểm tra AI `TG_NGAN`, bảng tổng hợp GPS. `gioDuCong` cũ bỏ không dùng. Kiểm thử `test/t36.js`.

### 5an. v3.15 – đăng nhập nhanh, AI 4 nhà cung cấp, phân tích song song

- `loadAll` (view_all): `cc`, `bl` chỉ gồm kỳ ≥ `kyCatCu()` (cùng tháng năm trước) hoặc kỳ chưa duyệt; `dlTu` = mốc nếu còn kỳ cũ. API `dl.cu` {truoc} (READ_ONLY). Client `nhanDL(d)` giữ kỳ cũ đã có (đã duyệt) hoặc hẹn `taiDuLieuCu()` tải ngầm; chỉ vẽ lại màn bc/dash/nvd/me/slip khi không đang nhập. `login` thuộc READ_ONLY (ghi lastLogin bằng khóa muộn).
- `aiEnv_`: DS gemini/openai/claude/deepseek; Script Properties AI_KEY_*, AI_MODEL_*, AI_MODEL2_* (model tổng hợp cho loai tong1/tong2/dgttong); OpenAI & DeepSeek theo Chat Completions (json_object; OpenAI max_completion_tokens + reasoning_effort low, tự bỏ tham số khi 400); DeepSeek không ảnh → nhờ nhà cung cấp khác có khóa, loai 'phan' thì bỏ ảnh; `dsModel` / API `ai.models`; `chonModelKhac_` thay model bị ngừng cùng dòng. `ai.cauhinh` nhận 4 ncc + model2.
- Phân tích nhóm: bước `lapnk` chỉ cần kết quả các đợt (chạy song song với tong1 → tong2 → lap); bước mới `chot` gộp J_kq + J_nk, ghi JSON. Client chạy các đợt 3 luồng.
- Kiểm thử: `test/t39_gas.js`, `test/t13.js` (song song), `test/smoke41.js`.

### 5ao. v3.16 – URL máy chủ từ webapp.txt

- Client: `docWebapp()` fetch `webapp.txt` (cùng thư mục, no-store, 4 giây) → `urlTuWebapp()` lấy dòng đầu khớp `https://script.google.com/(a/macros/<miền>|macros)/s/<id>/exec`, bỏ dòng `#` / `//`. `apWebapp()`: khác `APP.api` thì đổi, lưu `tphr_api`, đang đăng nhập thì `forceLogout()` + thông báo. Khởi động: chưa có URL → đọc tệp trước rồi mới hiện màn đăng nhập; đã có → vào app ngay, kiểm tra sau 0,8 giây; quay lại app sau 15 phút kiểm tra lại. Chạy từ file:// thì bỏ qua.
- Gói 1_UngDung có thêm `webapp.txt` (mẫu, chỉ có ghi chú). v3.17.1: mở từ file:// (trình duyệt chặn fetch tệp cục bộ) → màn đăng nhập có nút chọn webapp.txt (`napWebappTep`: đọc tệp, kiểm tra doGet, lưu tphr_api); hộp Cấu hình máy chủ có nút "Lấy URL từ tệp webapp.txt…"; công cụ `0_CongCu/GanURL_MayChu.bat` + `.ps1` (PowerShell 5+) ghi URL từ webapp.txt vào `const DEFAULT_API` trong index.html (UTF-8 không BOM, giữ .bak). Kiểm thử `test/smoke44.js`. Kiểm thử `test/smoke42.js`; chạy toàn bộ: `test/chay_tat_ca.sh`.

### 5ap. v3.17 – hồ sơ đính kèm trong lý lịch

- Core `TL_NHOM` (hd: HĐLĐ, phụ lục, quyết định → "1. Hợp đồng – Quyết định"; ll: Sơ yếu lý lịch → "2. Lý lịch"; cn: còn lại → "3. Hồ sơ cá nhân"), `tlNhom(loai)`; `tl.upload` ghi Drive theo `duong: [<mã – họ tên>, <thư mục nhóm>]`; API `tl.thumuc` {maNV, nhom?} → url thư mục; `tl.get`, `tl.thumuc` thuộc READ_ONLY.
- Client `hsDinhKemHTML(ma, self)` trong tab Thông tin (3 thẻ nhóm, Xem, + Tải lên theo nhóm), `openTLForm(ma, id, nhom)` lọc loại theo nhóm, chọn nhiều tệp `tlPickNhieu` (tải lần lượt), nhận .doc/.xls/.heic theo đuôi tệp.
- Kiểm thử: `test/t40_gas.js`, `test/smoke43.js`.

### 5aq. v3.18 – chấm dứt HĐLĐ

- `LOAI_QT` thêm 'Thanh lý hợp đồng'; `QT_CHAM_DUT` = Nghỉ việc, Thanh lý hợp đồng. `NhanVien.ngayNghi` = ngày chấm dứt (ngày đầu không còn làm việc). `ngoaiHD(nv, ngay)` = trước ngayVao hoặc ≥ ngayNghi; `dangLamTrongKy`: ngayNghi ≤ đầu kỳ → không có trong kỳ.
- `chamDutCua(maNV)` (bản ghi mới nhất trong Tiếp nhận / chấm dứt), `apChamDut(nv)` (gọi sau qt.save, qt.delete), `khoaTK(nv, lyDo)` (Users.khoaTu, khoaLyDo 'Tự động: …'; bỏ qua admin; khóa tay 'Khóa bởi …' giữ nguyên), `chamDutDenHan()` (login/load của view_all), `tkBiKhoa(u)` trong sess()/login. upsertUser: đổi active tay → ghi 'Khóa bởi …' hoặc xóa khoaTu.
- cc.save bỏ ô chiTietNgay ngoài thời hạn HĐ. Client: `dgOff = C.ngoaiHD`, `hdTag()` nhãn vào làm / chấm dứt giữa kỳ, ccFill dùng `soNgayLamHD()`; form QT, form hồ sơ (ngayNghi), trang Tài khoản hiện lý do khóa.
- Kiểm thử: `test/t41.js`, `test/smoke45.js`.

### 5ar. v3.19 – ngày dd/mm/yyyy, một thư mục dữ liệu

- Client: `dmyGan(input[type=date])` bọc ô ngày gốc (ẩn, giữ giá trị ISO và mọi xử lý cũ; ghi đè thuộc tính value để cập nhật ô hiển thị) bằng ô chữ `id + '__dmy'` dd/mm/yyyy (tự thêm /, `dmySangIso` kiểm ngày có thật, nút 📅 showPicker); `dmyChu` đổi chữ yyyy-mm-dd trong nút văn bản sang dd/mm/yyyy (trừ SCRIPT/STYLE/TEXTAREA/CODE); MutationObserver toàn trang. CSS `.dmy`.
- GAS: `thuMucMe_()` (Script Property DATA_ROOT, tên "TAPACO Nhân sự – Dữ liệu app (KHÔNG CHIA SẺ)", cùng thư mục cha với Google Sheet), `thuMucGoc_()` / `thuMucLuuTru_()` tạo trong thư mục mẹ; thư mục cũ (TL_ROOT, LT_ROOT) được moveTo thư mục mẹ 1 lần (cờ *_GOM); hàm chạy tay `gomDuLieuDrive()`; `caiDatLuuTru()` gọi gom.
- Kiểm thử: `test/smoke46.js`, `test/t42_gas.js`.

### 5as. v3.20 – công ty thành viên

- cfg `congTyTV: [{ ma, ten, tenTat, mst, diaChi, trangThai: 'Hoạt động'|'Ngừng' }]`, mặc định TPC (TAPACO) / TC (Thịnh Vượng Toàn Cầu) / ATP (An Thuận Phát), MST – địa chỉ để trống. `chuanCTV` chuẩn hóa trong mergeCfg (mã viết hoa A-Z0-9_, không trùng; rỗng → mặc định). Tài khoản không view_all cũng nhận `congTyTV`.
- Cột mới: `NhanVien.congTy`, `BangLuong.congTy` (upgradeSchema tự thêm). Hàm lõi: `congTyNV(nv,cfg)` (rỗng/không hợp lệ → công ty đầu danh mục), `congTyBL(b,nvMap,cfg)` (dòng lương cũ → theo hồ sơ), `nhanCTV(v,cfg)` (mã hoặc tên, không dấu), `tenCTV`, `dsCTV`, `tkCongTy(nvs, bls, kts, cfg, {kys, nam, ct})` → `{ congTy:[{dang, thuViec, nghi, nghiNam, vaoNam, khen, kyLuat, soPhieu, quy, bhNLD, bhDN, kpcd, thue, thucLinh, chiPhi}], tong, boPhan:[{ten, loai, dang, theoCT, …tiền}] }`.
- recalc (bl.calc) gán `p.congTy = congTyNV(nv, cfg)`. nv.save nhận mã/tên công ty, báo lỗi nếu không có trong danh mục. API `ctv.save {congTyTV}` (edit_nv): không xóa công ty còn hồ sơ NhanVien hoặc đã có BangLuong; đổi công ty đầu danh mục thì gán cứng công ty cũ cho hồ sơ đang để trống.
- xlChiPhi: cột Q "Công ty" ở DuLieu (lấy `b.tenCongTy`), sheet `TheoCongTy` (SUMIFS cột Q).
- Client: `ctvDs/ctvTen/ctNV/ctOpts/ctBLfn`; Tổng quan `dashCongTy()` (#dashCT thẻ `.ctvk-i` + bảng theo công ty; #dashBP bảng bộ phận gộp công ty, lọc 1 công ty; chọn kỳ tháng / cả năm, `APP.filt.dash`); form hồ sơ ô `congTy`; danh sách lọc `#nvCt`; vBP tab `ctv` (`ctvHTML`, `ctvForm`, `ctvGan` dùng nv.import); Báo cáo lương `BC.ct` + `bcLocCT/bcBlKy/bcTenCT/bcGanTenCT`. Dữ liệu mẫu: NV003, NV006 → ATP; NV007 → TC.
- Kiểm thử: `test/t43.js`, `test/smoke47.js`.

### 5at. v3.21 – bảng lương tháng tách theo công ty → bộ phận

- `BC.ct`: '' = tất cả, tách nhóm công ty (mặc định); '*' = gộp các công ty, chỉ theo bộ phận; mã = 1 công ty. `bcChonCT()`, `bcTachCT()`. bcThang: hàng `tr.bgc` (công ty), `tr.bgr` (bộ phận), `tr.bsub`, `tr.bsc` (cộng công ty), `tr.btot`.
- `xlBangLuongThang(ky, bls, ctl, nvMap, info)`: `info.nhomCT` + `info.thuTuCT` (tên tắt theo thứ tự danh mục), bls mang `tenCongTy`; style `grpc` / `totc`; cộng công ty = SUBTOTAL(9, đầu công ty : cuối công ty) (SUBTOTAL bỏ qua SUBTOTAL lồng); ChuyenKhoan thêm cột Công ty khi tách.
- Kiểm thử: `test/t44.js` (LibreOffice tính lại khớp), `test/smoke48.js`.

### 5au. v3.22 – chấm công ngày theo bộ phận, tự lưu nền

- `dgNVs(ky, tatCa)` sắp theo `dgBpThuTu()` (phòng ban rồi công trình theo cfg.boPhan, "Chưa xếp bộ phận" cuối), lọc `DG.bp`, `DG.ct`; dòng `tr.dgg[data-bp]` (bấm → `dgNhomTo(bp)`).
- Tự lưu: `dgUpd` → `dgHen(1500)`; `dgGui()` gửi DG.dirty (xóa ngay khỏi dirty, lỗi thì trả lại), api quiet, `applyKyData`, chỉ cập nhật DG.data của người không bị tô tiếp và khác bản máy chủ (`dgVeHang`), không render; `dgSave(silent)` lặp tới khi sạch, lỗi → toast 1 lần + hẹn 15 s; `DG.dang` chặn `dgLoad` ghi đè; beforeunload / visibilitychange đẩy lưu. Trạng thái `#dgTT` (`dgTTHTML`).
- Máy chủ cc.save: tra ChamCong theo mã kỳ (map), khi chưa tới bước tính lại (`canTinhLai` sai) trả `{ky, cc, pd}` – không đọc BangLuong.
- Kiểm thử: `test/smoke49.js`, `test/perf/perf_dg.js`, perf_gas (Lưu chấm công ngày).

### 5av. v3.23 – lưu nền hồ sơ / quá trình, tin nhắn hỏi nhanh

- Client: `nenLuu(ten, goi, xong, loi)` – hàng đợi Promise nối tiếp `NEN.chuoi`, đếm `NEN.n`, chỉ báo `#nenTT`; `napGoiNV({nv, qt, users})`; `veLaiNeuRanh()`; `moLaiForm(mo, v)`; `refreshNen()`. openNVForm (sửa) và openQTForm, delQT: cập nhật APP.D ngay, render, lưu nền; lỗi → hoàn lại + mở lại form. Thêm hồ sơ: `nv.save {kem:1}` → `{nv, qt, users}`, không gọi load. bp.save / bp.dieudong / ctv.save dùng dữ liệu trả về; mh.save, mh.decide, user.save tải lại ngầm.
- Lõi: `goiNV(ma)`; `nv.save` có `req.kem`; `qt.save` trả `{row, nv, qt, users}`; `qt.delete` trả `goiNV`. chat.init / chat.poll gọi `env.chat.ghiNho(u, idsHoiThoai, last)` (GAS: cache `myht_<u>`, `tn_last`).
- GAS: `pollNhanh_(req)` chạy trước SheetDB_: chat.poll không cờ online, phiên có trong cache, `cur == tn_last`, `hv == cv_<u>`, có `myht_<u>` → `{cur, hv, typing, nhanh:1}` không mở bảng tính. chDelay khi mở Tin nhắn 2000 ms.
- Kiểm thử: `test/t45_gas.js`, `test/smoke50.js`.

### 5aw. v3.24 – thêm hồ sơ / điều động lưu nền; OCR công văn

- Thêm hồ sơ: mã tạm `maNVGoiY()`, bản ghi `_tam`, `nenLuu(nv.save {isNew, kem, maTuDong})`; lõi: `isNew && cur && req.maTuDong` → `maNVMoi()`. Điều động: cập nhật APP.D ngay, `nenLuu(bp.dieudong)`.
- OCR: `ocrDrive(fileId)` = Drive v3 `copy?ocrLanguage=vi` (Google Docs) → `export?mimeType=text/plain` → trash; `env.ocr` (dùng thử). `cvChuTep`: JPG/PNG/GIF/BMP, PDF → OCR (≥ 40 ký tự) → `nguon:'ocr'`; khác → ảnh cho AI. `cv.ai {chu}` phân tích lại. `CongVan.noiDungOCR`; `cv.ds` trả `soChuOCR`; `cv.ocr {id}`.

### 5ax. v3.25 – quy định theo thời kỳ; bảng chấm công thêm giờ

- cfg `mocLuat: [{tu:'YYYY-MM', ten, luongCoSo, luongToiThieuVung, giamTruBanThan, giamTruNPT, bieu:'b7'|'b5', otCach:'chenh'|'toanbo', nguongVangLai}]` (mặc định `MOC_LUAT_MD`: 2024-07, 2026-01, 2026-07), `luat109Tu2026` (kỳ 01–06/2026 dùng b5 + toanbo), `ngayNghiTuan: 'cn'|'t7cn'`. `BIEU_THUE`, `OT_CACH`, `chuanMoc`, `mocCuaKy`, `cfgTheoKy(cfg, ky)` → cfg + tham số kỳ, `cheDoLuat`, `otCach`, `tenBieu`, `_ky`.
- `tinhPhieuLuong` gọi `cfgTheoKy(cfg, o.ky)` đầu hàm, trả `cheDoLuat` (cột BangLuong.cheDoLuat); `taoBien`: `OT_TL_MIEN` theo `cfg.otCach` (chenh: phần hệ số − 1; toanbo: giờ trong giới hạn / tổng giờ). `xlBangTinh` dùng cfg theo kỳ; recalc `ctCong.gtbt` theo kỳ.
- Làm thêm giờ: `ChamCong.chiTietOT` = `{ngày: giờ}`; `loaiOT(ky, d, cfg)` (lễ → ot300, nghỉ tuần → ot200, khác ot150), `cleanOT`, `tongHopOT` (vuotNgay: ngày thường > ½ giờ làm việc, > 12 giờ), `apGioOT(o, ky, cfg)` trong cc.save: OT150 = bảng thêm giờ; OT200, OT300 = bảng ngày (làm CN nếu chuNhatTinh='ot', làm lễ) + bảng thêm giờ. FIELDS_LOAI.chamcong thêm `chiTietOT`.
- Client: tab `ot` (`otTab`, `OTG`, `otHydrate` input/Enter, `otGui/otSave` lưu nền như DG, `otXuat`); Tổng hợp khóa OT khi có chiTietOT; Cài đặt `mocHTML/mocRow/mocDoc/mocMacDinh` (bỏ ô luongCoSo, LTTV, giảm trừ, biểu thuế, ngưỡng); phiếu lương dòng "Áp dụng", "tiền làm thêm giờ miễn thuế"; Báo cáo tab `sst` (`bcSSTinh`, `bcXuatSST`).
- Kiểm thử: `test/t48.js`, `test/t46.js`, `test/smoke52.js`, `test/smoke_all.js`.

### 5ay. v3.26 – 5 trường hợp tính lương; trạng thái tính thuế theo thời điểm; quyết toán năm

- `TRUONG_HOP` (bh, khoan, noikhac, huu, thuviec) ↔ `loaiLuong` + `lyDoKhongBH` (thêm 'Hợp đồng thử việc'); `TH_KHAC` cho hồ sơ cũ khongbh chưa ghi lý do; `truongHopNV(nv)`. nv.save nhận mã trường hợp. Form hồ sơ: 1 ô chọn trường hợp + mô tả (`llDoi`).
- `NhanVien.thueTT` = [{tu, cach: tudong|luytien|khautru10, giamTru, camKet, uyQuyen: tudong|co|khong, ghiChu}] (`chuanTT`); `ttThue(nv, ky)` → cách tính của kỳ (mặc định theo trường hợp: HĐLĐ → lũy tiến; noikhac → không giảm trừ, không ủy quyền; khoán / thử việc → khấu trừ 10%; camKet mặc định theo `camKet08`). `laVangLai(nv, ky)` = khautru10. API `nv.thue {maNV, thueTT}` (edit_nv hoặc payroll).
- `tinhPhieuLuong`: thuế theo `ttThue`; giảm trừ chỉ khi giamTru = co; trả `cachThue` (BangLuong.cachThue). BH: người hưu – DN đóng `tyLeDN.tnldHuu` (mặc định 0); thử việc – không BH. Biến `TRA_THAY_BH` (cfg.traThayBH='co', TH huu/thuviec) = MIN(LUONG_BH, trần) × (DN bhxh+bhyt+bhtn)%; cfg.save tự thêm khoản TRATHAYBH. Excel bảng tính: cột `gtc` (Giảm trừ gia cảnh tại Công ty 1/0).
- `quyetToanNam(nam, bls, nvs, cfg, {gtDuNam})` → từng người: soKy, tnct, bh, daKT, giamTruNam, tntt, phaiNop (biểu tháng 12 × 12), chenh, duUyQuyen, goiY[]. Báo cáo tab `qtt` (`bcQuyetToan`, `bcXuatQTT`, `bcGtDu`). Trạng thái thuế: `ttForm/ttRow`.
- Kiểm thử: `test/t49.js`, `test/smoke53.js`.

### 5az. v3.27 – TAPACO Payroll & Tax 2026
- Lõi: QUY_DINH_MD (LegalRules: ma, ten, vanBan, dieuKhoan, hieuLuc, hetHieuLuc, noiDung, url, trangThai xacminh|chinhthuc|chuaxacminh|hethieuluc), BIEU_MAU_MD (8 mẫu; trangThai hienhanh|dangxacminh|hethieuluc|noibo); cfg.quyDinh / cfg.bieuMau sửa qua `qd.save` (config hoặc duyetThue), chuanDsMa giữ mục sửa + bổ sung mục mặc định mới.
- NhanVien.hsThue (JSON, hsThueNV chuẩn hóa); bảng HoSoThue (id, maNV, nam, loai camket|uyquyen|npt|kekhaiGT|xacnhanBH, mauSo, phienBan, ngayKy, trangThai Nháp|Chờ phụ trách KT duyệt|Đã duyệt|Từ chối|Chờ xác minh mẫu, duLieu, ketQua, nguoiLap, ngayLap, nguoiKiemTra, ngayKiemTra, ghiChuDuyet, taiLieuId, tenTep, ghiChu, capNhat). Handler nv.hsthue, hst.save (bản ký → TaiLieu loại "Hồ sơ thuế TNCN"), hst.duyet (chặn mẫu đang xác minh nếu không xacNhanMau; đồng bộ thueTT), hst.delete.
- Vai trò kttruong (Phụ trách kế toán), quyền duyetThue. logDoi: nhật ký trước → sau.
- Hàm: kiemTraThuViec (TV_VITRI 180/60/30/6 ngày làm việc, ≥ 85%), dauHieuQHLD (≥ 2 dấu hiệu → đỏ), uocTinhNam, giamTruNam (y tế ≤ 23 tr, giáo dục ≤ 24 tr, hưu trí ≤ 3 tr/tháng), evaluateTaxCommitmentEligibility, evaluateTaxAuthorizationEligibility (vãng lai ≤ 15 tr), toiUuThue (A/B/C, chuaKhauTru, thueGiamTheoLuat, thuaHoan), canhBaoThueBH, PayrollEngine (calculateGrossSalary … calculateAnnualTaxSettlement, tinh, netToGross).
- ttThue: trường hợp 3 giảm trừ theo hsThue.noiKhac.gtTai; khoán kinh doanh → cach 'kinhdoanh' (thuế 0). mocLuat.anCaMien → cfg.anCaMienThue (ANCA 730.000 → 1.200.000 từ 07/2026). nv.thue chặn giảm trừ trùng.
- Giao diện: tab hồ sơ 'tt' Tối ưu thuế TNCN, hsThueForm, hstForm, hstDuyet, bmHTML/bmMo (In/PDF, Word .doc), tbhPills/tbhChiTiet, cột bảng lương mới + lọc, Báo cáo tab 'uq', 'bm', qdSua. Kiểm thử test/t50.js (74 kiểm tra, 12 bộ dữ liệu + ranh giới, tính độc lập), test/smoke54.js.

### 5ba. v3.28 – Module hợp đồng khoán việc theo tháng (CONTRACT-02)
- Bảng HopDongKhoan, ThanhToanKhoan; handler hdk.save, hdk.hoso (tải tệp vào Drive TAPACO/PAYROLL/NAM_yyyy/HOP_DONG_KHOAN/maNV/0x_…, trạng thái HS_TT), ttk.save, ttk.move (QT_KHOAN_MD 6 bước: BCH, KTKH, KT, PTKT, GD, TT), ttk.delete; quyền: edit_nv | payroll | duyetThue | BCH công trình mình.
- Hàm: classifyContractRelationship (CH_BAN_CHAT q1–q11, CH_TO_DOI, KL_PL), evaluateEmploymentInsuranceRisk, classifyContractorTaxRegime (CHE_DO_K tiencong10 | kinhdoanh | khongcutru | hdld | thieu), calculateMonthlyContractPayment, calculateContractorWithholdingTax (ngưỡng theo ngày trả), khoanNetToGross, tamUngTon, estimateAnnualContractorIncome (A/B/C/D), calculateAnnualPersonalDeductions (NPT theo tháng – hsThue.nptDs), validateTaxCommitmentDocuments, evaluateTaxCommitment (CK_TT), estimateAnnualPersonalIncomeTax, estimatePotentialTaxRefund, evaluateTaxSettlementAuthorization, quyenLoiThueKhoan, generateContractorTaxNotes, generateContractorPaymentStatement; biểu mẫu bm01–bm07 (bmKhoanHTML).
- Bảng lương: cfg.ckBatBuocDuyet → bl.calc truyền o.ckHopLe; quyetToanNam(opt.ttk). QUY_DINH: BLLD_13, ND68, TNCN_KCT, L109/ND253/TT89_GOC/ETAX (urlKiemTra 'chua').
- Kiểm thử: test/t51.js (32 ca gồm 18 tình huống bắt buộc, ghi test/t51_ketqua.json), t52.js, smoke55.js.

### 5bb. v3.29 – Nhân sự hưởng lương hưu
- hsThue.huu: tinhTrang (HUU_TT), ngayHuong, soQD, donViChiTra, luongHuu, theBHYT, trangThaiXM/ngayXM/nguoiXM (HUU_XM, chỉ qua nv.huuxm), phuongAnBH (HUU_PA A|B), canCuBH, congViecNguyCo; tnKhac.tuQTBatBuoc. NhanVien.hoSoHuu (HS_HUU 10 mục, nv.hshuu → Drive NhanSu/mã/HoSoHuuTri/năm). BangLuong.traThayBHCT.
- tinhPhieuLuong = lớp hưu trí (huuTaiKy, traThemBH) bọc tinhPhieuLuongGoc; o.traThayBH ghi đè biến TRA_THAY_BH; công thức chưa có khoản TRA_THAY_BH thì tự thêm dòng TRATHAYBH chịu thuế.
- calculateRetireePIT, chiTietLuyTien, evaluateRetireeTaxAuthorization, phanTichHuuTri, ghiChuHuuTri, huuChuaXacMinh (chặn ky.move → Đã duyệt).
- QUY_DINH: BLLD_149, TNCN_LUONGHUU, TNCN_UQ_TUQT, BLLD_GOC, BHXH_GOC, TT87 (urlKiemTra 'chua'); BIEU_MAU bmh01, bmh02. Kiểm thử test/t53.js (22 ca), smoke56.js.

### 5bc. v3.30 – Phiếu nhập / xuất kho trong Tin nhắn
- Bảng PhieuKho (id, hoiThoai, soPhieu PN|PX-yyyymm-nnn, loai nhap|xuat, ngay, kho, doiTac, hangMuc, nguoiGiaoNhan, soCT, dong JSON [{ten, quyCach, dv, sl, donGia, thanhTien, ghiChu}], anh JSON [{fileId, ten, mime, w, h, moTa}], ghiChu, nguoiLap, hoTenLap, taoLuc, msgId, trangThai, lyDoHuy, nguoiHuy, capNhat).
- Handler chat.kho (ảnh tải trước bằng chat.upload, kiểm tra cùng nhóm, đúng người tải, chưa gắn tin), chat.khods, chat.khoxem (READ_ONLY), chat.khohuy; tin loại 'kho' (tep: pk, so, loai, ngay, 4 dòng đầu, 4 ảnh đầu); chat.recall chặn loại 'kho'.
- Giao diện: chKhoForm (Chụp ảnh capture=environment, Tải ảnh multiple, btNen), chKhoCard, chKhoXem, khoLb, khoIn, tab thông tin nhóm 'kho' (khoTab, khoTong, khoXuat). Kiểm thử test/t54.js, smoke57.js.

### 5bd. v3.31 – Tiền làm thêm giờ cho HĐLĐ chính thức
- CT_LUONG_TAPACO.OT_QD = OT_QD_MOI: IF(LA_THU_VIEC, ROUND((OT200*2 + OT300*4 + OT150*1.5)/GIO_CONG_NGAY, 2), ROUND((OT150*HS_OT150 + OT200*HS_OT200 + OT300*HS_OT300)/GIO_CONG_NGAY, 2)); mergeCfg gọi vaOTQD vá cấu hình còn công thức cũ (OT_QD_CU2). Kiểm thử test/t55.js.

### 5be. v3.32 – Chấm ăn ca, công nợ 4 khoản, 2 bảng lương, tăng tốc
- ChamCong + chiTietAn {ngày: 0–3}, soBua, tienAnThu, thuUng; FIELDS_LOAI.chamcong + chiTietAn; congno = tamUng, tienAnThu, thuUng, khauTruKhac. noiLamViec + anCaBua, giaBua. Biến SO_BUA, TIEN_AN_BUA, GIA_BUA; ANCA = ANCA_BUA (vaANCA vá cấu hình cũ).
- tinhPhieuLuong: tongNo = tamUng + tienAnThu (bảng chấm ăn: SO_BUA × GIA_BUA) + thuUng + khauTruKhac; ăn ca miễn thuế mienAnCa = MIN(mức kỳ, R(NGAY_CONG × mức ÷ anCaMienChia 22)) khi anCaMienTheoCong; trả anCaMien, otThue (thuế cả tháng − thuế khi bỏ phần chịu thuế của dòng phanLoai tangca), otThucLinh.
- cfg.cheDoBangLuong 'gop' | 'tach'; xlBangLuongThang(info.tach) → BangLuong (không cột tăng ca), TangCa, ChuyenKhoan, ChuyenKhoan_TangCa.
- ghiCC(ky, b, rows) dùng chung cc.save / ky.move (req.rows: lưu + chuyển 1 lần); recalc cập nhật tại chỗ dòng BangLuong.
- Tin nhắn: chat.upload {gui, file.thumb, w, h} → {fileId, thumbId, msg}; TepChat.thumbId; chat.file {thumb}; chat.anh {ids ≤ 8, thumb} (READ_ONLY); phản hồi kèm _tf / _fd / _tfXoa để gas.js ghi bộ nhớ đệm (tf_<fileId>, chfd_<ht>) rồi xóa trước khi trả. gas.js: tepNhanh_ (đọc ảnh không mở Sheet), taiTruoc_ (ghi Drive trước khi lấy khóa; req._pre chỉ máy chủ đặt, handle_ xóa _pre từ máy khách), boTaiTruoc_. Client: chNen (1600 px 82 % + 480 px), chNhieu (3 song song), chFetch + Cache Storage 'tapaco-chat-anh-v1', IntersectionObserver; sw.js giữ cache này. Kiểm thử t56, t57_gas, smoke58.

### 5bf. v3.33 – Lịch sử hợp đồng lao động
- NhanVien.lichSuHD (JSON) [{id, soHD, loaiHD, tuNgay, denNgay, luongHD, luongBH, hinhThuc bh|khongbh|khoan, lyDo, tlId, tenFile, ghiChu, trangThai Hiệu lực|Đã hủy, lyDoHuy, nguoi, luc, daBao}]; LOAI_HD + 'Cộng tác viên (CTV)'; LY_DO_KHONG_BH + 'Khoán việc, thời vụ, CTV' (truongHopNV → khoan).
- Handlers hd.save (thêm dòng + TaiLieu "Hợp đồng lao động" + QuaTrinh "Thay đổi HĐLĐ"; trả kyDaDuyet), hd.huy (lý do), hd.tep; nv.save bỏ lichSuHD và các trường HĐ khi đã có lịch sử; hdDongBo đồng bộ trường hiện tại theo HĐ hôm nay.
- mhNguon thêm sự kiện 'hd' (luongChinh = luongHD, luongBH) tại tuNgay. apHD: loại HĐ → loaiLuong / lyDo (HD_KHONG_BH: thử việc, thời vụ – khoán, CTV không BH). hdDuoi3Thang (XĐTH dưới 3 tháng) → ttThue khấu trừ cố định.
- dauVaoTinhLuong trả segs = hdDoanKy(nv, ky); recalc: 1 giai đoạn → tinhPhieuLuong với hồ sơ theo HĐ; ≥ 2 → tinhNhieuDoan (o.khongBH cho giai đoạn trước, thuế gộp theo cách tính); BangLuong.ctHD.
- MOC_LUAT_MD + nldXH, nldYT, nldTN, dnXH, dnYT, dnTN, kpcd, tnld, thueVL; cfgTheoKy áp tyLeNLD / tyLeDN / thueVangLai theo mốc; mergeCfg điền mốc cũ. cfg.canhBaoHD (30); chayLich → hdNhac mỗi ngày (cfg.hdNhacNgay), báo r:tchc. Kiểm thử t58, smoke59.

### 5bg. v3.34 – ChamCong.anThuTay: cc.save tienAnThu '' → theo bảng chấm ăn; khác số tính → anThuTay=1 (giữ khi chấm ăn lại); engine dùng tienAnThu khi anThuTay.

### 5bh. v3.35 – In lý lịch: inLyLich(ma) (chọn mục) → llHTML(ma, {tt, qt, hd, kt, luong, nh}) vào #print, CSS .llp (@page A4 dọc). Kiểm thử smoke60.

### 5bi. v3.36 – Lý lịch: llMuc(ma, o) → các mục {t, kv | cols/rows/rong}; llHTML (in / PDF), llWord → khoiSangDocx (thêm khối 'kv', ô link {x, url} → ExternalHyperlink). LL_MUC: tt, qt, hd, bh, thue, huu, tl, kt + luong, nh. moTheo '#tl=<id>' → viewTL. Kiểm thử smoke60.

### 5bj. v3.37 – OCR hồ sơ: TaiLieu + ocrId, ocrLuc, ocrNguon, ocrTom. Core ocrTach(text, loai) → {nv, tl, nhan}; OCR_TRUONG, OCR_TL; b64Utf8. Handlers tl.ocr (req._ocr do gas.js đọc trước khóa – chỉ dùng khi _ocrFid trùng fileId; req.text từ máy khách; env.files.ocr; lỗi → 'OCR_CHUA_BAT'), tl.ocrxem (READ_ONLY), tl.ocrluu (edit_nv, kiểm tra CCCD 12 số / không trùng, MST 10 hoặc 10-3 số); tl.delete xóa cả ocrId. gas.js ocrDrive_ (Drive API v3 create → Google Docs, ocrLanguage vi; v2 insert ocr:true), ocrTruoc_. Client: ocrMay (pdf.js 3.11.174 + tesseract.js 5.1.1), tlOcr / ocrSoat / tlOcrHang / tlOcrXem; modal onClose; cfg.ocrTuDong. Kiểm thử t59, smoke61.

### 5bk. v3.38 – Bảng phân quyền: PERM thêm xem_bp (MD truongphong, cht), tu_xem (mọi vai trò trừ admin, PQ_KHOA toàn bộ), sua_pq (MD + khóa admin, giamdoc); PQ_CAM {tu_xem:[admin], sua_pq:[nhanvien], xem_bp:[admin]} bị lọc trong apDungPQ / pqLamSach. PQ_NHOM thêm nhóm "Theo chức danh". pq.save kiểm tra can(role,'sua_pq'). loadAll: bpXem(s) (ROLE_QL → s.boPhan; vai trò khác có xem_bp → phòng ban của mình) → trả bpXem, bpNV, bpQT, đơn nghỉ bộ phận. Client: bpXemTen(); allowed('ql') = !view_all && bpXemTen(). Bảng phân quyền: dòng Lương = từng bước cfg.quyTrinhLuong, Nghỉ phép = từng cấp cfg.quyTrinhNghi; ô = 'r:<vai trò>' trong nguoi (PQ.qt nháp, qtDoi, pqQtTick chặn bỏ người cuối); Lưu gọi quytrinh.save (quyTrinhLuong + apDungKyDangXuLy / quyTrinhNghi) và pq.save. Test t60, smoke62.

### 5bl. v3.39 – Ăn ca miễn thuế: cfg.anCaMienCT (trống → theo anCaMienTheoCong / anCaMienChia cũ). ctMienAnCa(cfg), CT_ANCA_MD = 'MIN(MUC_MIEN, NGAY_CONG * MUC_MIEN / 22)'. mienAnCa(c, cfg, bien, gt) chỉ cho ANCA: tinhCT với bien + MUC_MIEN (mucMienAnCa theo kỳ) + GT; lỗi → CT_ANCA_MD; kẹp [0, MUC_MIEN]. kiemCTAnCa kiểm biến (BIEN_HT, PC_, NLV, bienKy, mã khoản lương trừ ANCA, MUC_MIEN, GT); cfg.save từ chối nếu lỗi. Excel: MAX(0,MIN(mt,ROUND(ctExcel…))). Client: ô text + xem thử (fieldHTML f.inp, f.xem). Test t61, smoke63.

### 5bm. v3.40 – Phiếu lương tách: slipHTML(b) = slipTach(b) (blTach() && otGTr(b) > 0) ? slip1HTML(b, true) + slipOTHTML(b) : slip1HTML(b, false). slip1HTML p1: bỏ phanLoai 'tangca', tổng = tongThuNhap − otGTr, thuế phiếu = thueTNCN − otThue, còn nhận = thucLinh − otTLr. slipOTHTML: khoản tangca trong chiTiet (hoặc tienOT), otThue, otTLr. In: mỗi .slip 1 trang. Test smoke64.

### 5bn. v3.41 – Mẫu phiếu / bảng lương: core MAU_PHIEU_MD {tieuDe, tieuDe1, tieuDe2, khoi[15] {k, ten, khoa}, the[10], ky[]}, MAU_BL_MD {cot[27] {k, ten, khoa, an}, ky[]}; mauDs (giữ thứ tự người dùng, chèn mục thiếu cạnh mục kề, khóa luôn hiện; an = ẩn khi in, anX = ẩn Excel), mauPhieu(cfg), mauBL(cfg), mauGon. cfg.mauPhieu / cfg.mauBangLuong làm sạch trong cfg.save; mauPhieu có trong cfg của CBCNV. xlBangLuongThang(info.mau): cột mk, nhóm khấu trừ sắp theo mẫu, tkt = SUM(đầu:cuối nhóm), cột anX → {wch, hidden}, đổi tên, chữ ký. sangXLSX nhận cols dạng object. Client: slip1HTML dựng theo khối K[k] (mỗi mục 1 table, số La Mã theo thứ tự hiện), slipKy; printBL theo MB.cot; Cài đặt mauCardHTML (draggable + ▲▼, MAU nháp, mauTam xem thử, mauLuu). Test t62, smoke65.

### 5bo. v3.42 – Core: boDau, diemTim(q, t) (chuỗi con không dấu → 88–100; từng từ: đầu từ / chứa / Levenshtein ≤1 (≥4 ký tự) ≤2 (≥7) → 45–85; chữ đầu → 70; chuỗi có số chỉ khớp đúng; ký tự theo thứ tự → ≤35), locGan(q, ds, layChu, toiDa). cv.ds trả ocrTim (1.500 ký tự đầu OCR). Client: timNVNang (MutationObserver) bọc mọi <select> có ≥2 mã CBCNV bằng ô .tim-ip + danh sách #timDs (position fixed), giữ select gốc ẩn (value / change không đổi); cvTimHTML / cvTimGo (CVS.q, CVS.tatCaNam) chỉ vẽ lại #cvBody; ccGrid(ky, cur, 'cong'|'no'), tab 'no' = Công nợ. Test t63, smoke66.

### 5bp. v3.43 – App Android: thư mục `android/` (src: AndroidManifest.xml, apktool.yml, res/values/strings.xml, smali vn/tapaco/nhansu: MainActivity (WebView, xin quyền vị trí / camera, back, file chooser kết quả), Khach (WebViewClient: phục vụ assets/www tại https://appassets.androidplatform.net, link ngoài → Intent VIEW), Chrome (WebChromeClient: cấp định vị, camera getUserMedia, chọn tệp / chụp ảnh qua MediaStore), CauNoi (@JavascriptInterface TapacoApp.luuTep(ten, mime, base64) → Download/TAPACO + mở; inAn(ten) → PrintManager; phienBan()), In (Runnable in)); tools/apktool.jar 2.9.3 + apksigner.jar; build_apk.sh <1_UngDung> <x.yy> <out.apk> (icon từ icon-512, ký v1–v3, tạo khóa nếu chưa có ở android/khoa). Client: window.TapacoApp → chặn a.click()/a[download] blob → appLuuTep; window.print → inAn; xlGhi thay XLSX.writeFile; pdfKhung (iframe PDF → nút mở); không đăng ký SW; pushLayToken báo chưa hỗ trợ; caiAppDT (menu), cfg.linkAPK. Test smoke67.

### 5bq. v3.44 – Mức hưởng: mhSuaHL(s) = admin || can approve (sửa / xóa khai báo Hiệu lực, mh.goc). H['mh.goc'] {maNV, giaTri}: mỗi khoản → mốc đầu khai khoản (mhNguon): mh → cu[k]; ll → luongCu / luongBHCu; hd → báo sửa ở HĐLĐ; không có mốc → hồ sơ (nv[k] / phuCap); rồi mhSauThayDoi. Client: mhMatrix ô .mh-o → mhOClick(ma, mốc, k); cột ✎ / ✕; mhGocForm; openMHForm(ma, id, tiêu điểm, ngày HL) + nút Xóa khai báo; mhSuaO ở Mức đang hưởng. Test t64, smoke68.

### 5br. v3.45 – Ảnh nhanh: bt.upload nhận file.thumb → thumbId (btTep giữ thumbId, btIds xóa cả thumb); H['bt.anh'] {ids ≤ 12, thumb} kiểm btThay; chat.anh ≤ 12. gas.js: fcLay_/fcCat_/fcXoa_ (CacheService 'fc_'+id, ≤ 95.000 ký tự base64, 6 giờ) trong env.files.get / put / thumb / del, docTepChat_, taiTruoc_. Client: btNen tạo thumb 640px q0.72; btFetch(fid, thumb) (BT_IMG, Cache Storage khóa 'bt:'+fid[#t], hàng đợi BTQ gộp 12); btNap/btHydrate IntersectionObserver 800px, bài 1 ảnh nâng cấp ảnh nét, .lb ảnh gốc; chThumbGui gộp 10 (thiếu kết quả → chat.file). Test smoke69, t57_gas.

### 5bs. v3.46 – Cắt ảnh CCCD (client): cccdTai → cccdCatAnh(file, mặt) (modal: ảnh + SVG đa giác + 4 .cat-h kéo bằng pointer events, xem trước catNan 320px, Xoay 90° = xoay thứ tự góc, Tự nhận lại, Dùng ảnh gốc = btNen, Hủy → không tải). catTimGoc(canvas ≤ 400px): khác màu nền (trung vị dải viền 4%, L/A/B) + 0,35·gradient → Otsu → đóng 2 lần → vùng liền lớn nhất (ưu tiên chứa tâm) → bao lồi → 4 điểm cực (nếu ≥ 90% diện tích bao lồi) hoặc hình chữ nhật xoay nhỏ nhất; catXep (thuận chiều kim, góc trên-trái, cạnh dài ở trên). catNan: phép biến đổi phối cảnh 8 ẩn (khử Gauss) + nội suy song tuyến → 1280 × 807 (tỷ lệ 85,6/53,98), JPEG 0,9. Test smoke70.

### 5bt. v3.47 – Ảnh thẻ: tl.upload {anhThe: 1 | 'cccd'} → TaiLieu loại 'Ảnh thẻ', tên 'Ảnh thẻ 3×4' (không cần có trong danh mục; thay bản cũ như cccdMat). Client: anhTheBan, atAvatar (.av-anh), anhTheHTML, atTai / atTuCCCD → atKhung (khung 3:4 kéo / đổi cỡ, AT_KHUNG {x .03, y .29, w .25} trên thẻ đã nắn), atCatCV 600×800; llAnh(ma) (ảnh thẻ → cắt CCCD mặt trước) cho llIn (async) / llWord (khối kv + anhData → bảng 2 cột có ImageRun 90×120). Hệ số năng suất: cfg.hsns {bien, tieuChi [{ma, ten, tyTrong, tuDong:'chuyencan'}], xepLoai [{ma, ten, tu, hs}]}, hsnsCfg, chuyenCan, tinhNS(cfg, ccRow, ncc); ChamCong.chamNS (JSON {d, hsTay, ghiChu, nguoi, luc}), FIELDS_LOAI.tinhluong + 'chamNS'; ghiCC ghi chamNS và bienKy[bien] = hệ số. Client: tab 'ns' (nsTab, nsDong, nsLuu, nsXuat, nsIn), Cài đặt nsCfgHTML. Test t65, smoke71.

### 5bu. v3.48 – Cột biến tính lương: MAU_BL_MD thêm cột (an, anX mặc định ẩn) + BL_THEM; mauBLMd(cfg) thêm 'ct:MÃ' (khoản Công thức lương) và 'bk:MÃ' (biến kỳ) trước 'nganHang'; blGiaTri(b, k, {tach, cc}) dùng chung cho in và Excel (bhNLD, tongKhauTru, anCaNhan = chiTiet ANCA, tnKhongThue, dnBH, ct:, bk: từ ChamCong.bienKy, tách 2 bảng). xlBangLuongThang: info.cc; cột thêm (không anX) đặt trước "Ký nhận". cfg.save / mauGoi làm sạch mẫu với cfg đầy đủ (giữ cột ct: / bk:). Client: printBL val → C.blGiaTri; ô lọc #mauLoc. Test t66, smoke72.

### 5bv. v3.49 – Cài app điện thoại: build_apk.sh gọi android/zipalign.py (resources.arsc STORED, căn 4 byte, kiểm `--kiem`) trước apksigner (v1+v2+v3) – thiếu bước này Android 11+ báo "Ứng dụng chưa được cài đặt". cfg.linkWeb (https, cfg.save kiểm cả linkAPK); caiAppDT: link web + Sao chép, thứ tự theo iOS/Android, cảnh báo file://, nút cài PWA qua beforeinstallprompt. apple-touch-icon ./icon-180.png. Thư mục gói 6_DangWeb_iPhone. Test t67 (+ đường dẫn APK), smoke73.

### 5bw. v3.50 – login/load trả data.svVer; client mcCu/mcBao cảnh báo admin khi máy chủ cũ hơn app; lưu ảnh thẻ gặp lỗi "Chọn loại tài liệu" (máy chủ < 3.47) → báo cần cập nhật Code.gs và triển khai phiên bản mới. Test smoke74.

### 5bx. v3.51 – docWebapp đọc WEBAPP_NGUON (webapp.txt trên GitHub) trước, dự phòng ./webapp.txt; áp dụng cả file:// và APK; chỉ nhận script.google.com/…/exec; APP.webappNguon hiện trong Cấu hình máy chủ. Test smoke75.

### 5by. v3.52 – Lý lịch tải lên → ocrQT tách "Tóm tắt quá trình công tác" (Word tách chữ trên máy bằng docxVanBan; PDF/ảnh qua OCR) → ocrSoat hiện bảng QT tích chọn/sửa → tl.ocrluu {qt[]} lưu QuaTrinh loại "Theo lý lịch" (QT_PHU: không khép/không đổi vị trí), bỏ trùng; dòng trùng thời gian tại Công ty không tích sẵn. Test t68, smoke76.

### 5bz. v3.53 – cfg.gps.ggKey → Google Maps JS API qua lớp tương thích Leaflet (GWMap/GWLop/GL, lx(m)); lỗi khóa → OpenStreetMap; tab Điểm chấm công bấm bản đồ → gqBamThem thêm điểm. Test smoke77.

### 5ca. v3.54 – Đề xuất lương nhiều cấp: cá nhân / TP-CHT / TCHC lập → TP/CHT → TCHC → Giám đốc (cfg.quyTrinhLenLuong, QT_LL_MD); LenLuong.luongDuyet/capDuyet/lichSu; bỏ cấp trùng người lập; Giám đốc vượt cấp; view myll. Test t69, smoke78.

### 5cb. v3.55 – Trợ lý quy định: env.dem (CacheService) đệm câu trả lời 6 giờ theo phiên bản dữ liệu + đệm chữ văn bản; prompt 30.000 ký tự; AI lỗi → tra cứu từ khóa; bot.kiem đo từng bước; client đếm giây, hỏi lại 1 lần khi mạng rớt. Test t70, smoke79.

### 5cc. v3.56 – APK onActivityResult đọc ClipData (chọn nhiều ảnh); client anhMo() mở ảnh có dự phòng (createImageBitmap → <img>), báo HEIC, báo khi không nhận tệp. Test smoke80.

### 5cd. v3.57 – Form mục quy định: tải tệp (docx/pdf/ảnh/txt) → trích chữ trên máy → điền Nội dung (+ Tiêu đề, Căn cứ, Hiệu lực); noiDung ≤ 40.000 ký tự giữ xuống dòng; bot.hoi chỉ gửi đoạn liên quan (botDoanLQ). Test t71, smoke81.

## 6. API (action)

`ping, login, logout, load, changePassword, nv.save, nv.import, nv.delete, qt.save, qt.delete, ll.save, ll.decide, ll.delete, kt.save (nhiều maNVs), kt.decide (nhiều ids), kt.delete, cc.save, bl.calc, ky.move, yk.create, yk.reply, np.create, np.decide, np.cancel, quytrinh.save, congthuc.save, congthuc.test, tl.upload, tl.update, tl.get, tl.delete, user.save, user.bulk, user.delete, cfg.save, log.list`.

Quy tắc chính:
- Đăng nhập: SHA-256(salt|password); sai 5 lần khóa 10 phút; trả token + dữ liệu theo vai trò. Tài khoản mới `mustChange = Có` → buộc đổi mật khẩu.
- `nv.save` hồ sơ mới có ngày vào → tự tạo quá trình "Tiếp nhận". Mã NV viết hoa. Cập nhật một phần (chỉ trường được gửi).
- `qt.save`: bản ghi mới tự khép quá trình đang mở trước đó (denNgay = tuNgay − 1); nếu là quá trình mới nhất đã hiệu lực → cập nhật phòng ban/chức vụ/đơn vị; "Nghỉ việc" → trạng thái Đã nghỉ việc + ngày nghỉ.
- `ll.save` lưu lương cũ tại thời điểm đề xuất; `ll.decide` duyệt → `applyDueRaises()` (cũng chạy mỗi lần login/load của vai trò quản lý) áp quyết định có hiệu lực ≤ hôm nay vào hồ sơ, đặt ngayHuongBac khi là nâng bậc.
- `kt.decide` không cho duyệt thưởng gắn kỳ đã ở bước Chờ duyệt / Đã duyệt; nếu kỳ đang ở bước 2–4 thì tính lại bảng lương.
- `cc.save`: chỉ nhận các cột của bước hiện tại (CC_FIELDS) và chỉ người phụ trách bước; sau bước Chấm công không thêm dòng mới; ở bước 2–4 lưu xong tự tính lại toàn kỳ. Trả về {ky, cc, bl, pd, thieu}.
- `yk.create {ky, chuDe, noiDung}`: chỉ chính CBCNV, chỉ kỳ Đã duyệt có phiếu của mình; 10–2.000 ký tự; tối đa 5 ý kiến đang mở/kỳ. `yk.reply {id, noiDung, trangThai?, boPhan?}`: chủ ý kiến chỉ nhắn (nhắn khi đã đóng → mở lại 'Đang xử lý'); người xử lý = role trùng boPhan, giamdoc, admin — đổi trạng thái (lần đầu tự 'Đang xử lý'), chuyển bộ phận, đóng ý kiến bắt buộc có nội dung; thay đổi trạng thái/bộ phận ghi vào ghiChu của tin. loadAll: nhanvien nhận yk của mình, quản lý nhận tất cả.
- Nghỉ: `soNgayDon` đếm ngày trong [tuNgay, denNgay] trừ Chủ nhật + ngayLe, nửa ngày = 0,5; `ngayNghiTheoKy` cắt theo tháng; `quyPhepNam(nv, năm)` = ngayPhepNam ghi đè hoặc 12 + floor(số năm (năm − năm vào) / 5), năm vào/nghỉ việc tỷ lệ theo tháng (vào sau ngày 15 không tính tháng đó), làm tròn xuống 0,5. `np.create`: CBCNV cho mình, TCHC (edit_nv) cho người khác; bắt buộc lý do trừ phép năm; chặn trùng đơn mở/đã duyệt, chặn kỳ lương đã qua bước Chấm công, phép năm chặn vượt quỹ (quỹ − đã duyệt − đang chờ) và không vắt năm. `np.decide`: Chờ duyệt → hr (soNgay > nghiCanGDDuyet → Chờ GĐ duyệt, ngược lại Đã duyệt) hoặc giamdoc/admin (Đã duyệt); Chờ GĐ duyệt → giamdoc/admin; không tự duyệt đơn của mình; từ chối cần lý do; kiểm lại quỹ khi duyệt. `np.cancel`: chủ đơn khi đang chờ hoặc đã duyệt mà chưa tới ngày; TCHC/GĐ mọi lúc; không hủy đơn đã duyệt rơi vào kỳ đã khóa.
- `bl.calc`: chỉ ở bước Tính lương, vai trò ketoan/admin.
- `ky.move {ky, to, note}`: tiến đúng 1 bước (người phụ trách bước hiện tại; rời bước 2–4 thì tính lại và chặn nếu còn người thiếu chấm công; sang Đã duyệt thì lưu soNguoi, tổng thu nhập, tổng thực lĩnh, maKiemTra = 12 ký tự đầu SHA-256 của [maNV, tongThuNhap, bhxh, thueTNCN, thucLinh] đã sắp xếp); lùi về bước bất kỳ trước đó cần ghi lý do (người phụ trách bước hiện tại; từ Đã duyệt chỉ admin = Mở khóa, xóa kết quả duyệt). Mỗi lượt ghi PheDuyet.
- `nv.delete` chỉ admin và chỉ khi chưa có bảng lương.

## 7. Màn hình

- **Đăng nhập**: logo, form, link "Cấu hình máy chủ" (kiểm tra URL bằng GET trước khi lưu), khối tài khoản dùng thử khi chưa có máy chủ.
- **Tổng quan**: 4 thẻ số (CBCNV đang làm, quỹ lương kỳ gần nhất có bảng lương, tổng chi phí lương DN, việc chờ duyệt); danh sách đến hạn nâng bậc; HĐLĐ hết hạn trong 30 ngày; cột quỹ lương 6 kỳ; cơ cấu phòng ban; sinh nhật trong tháng.
- **Hồ sơ CBCNV**: tìm kiếm, lọc phòng ban/trạng thái, bảng (mã, tên, phòng, chức vụ, ngày vào, thâm niên, loại HĐ, lương, trạng thái, đã cấp TK); thêm/sửa (form 4 nhóm: cá nhân, công việc & HĐ, lương & phụ cấp, thuế–BH–ngân hàng), nhập/xuất Excel theo nhãn cột tiếng Việt.
- **Chi tiết hồ sơ** (dùng chung cho "Hồ sơ của tôi" ở chế độ chỉ xem): thẻ đầu (avatar chữ tắt, trạng thái, loại HĐ, thâm niên, nút thao tác theo quyền) + tab Thông tin / Quá trình công tác (dòng thời gian gộp quá trình, lên lương, KT-KL) / Lương & phụ cấp / Khen thưởng – KL / Phiếu lương.
- **Quá trình công tác**: bảng toàn công ty, lọc loại/năm, thêm/sửa/xóa.
- **Lương & nâng bậc**: tab Đề xuất & quyết định (duyệt/từ chối có số QĐ), Đến hạn nâng bậc (nút Đề xuất), Thang bảng lương. Chọn ngạch/bậc tự điền mức.
- **Khen thưởng – Kỷ luật**: thêm cho nhiều người một lúc (danh sách tích chọn có lọc), duyệt hàng loạt, lọc loại/trạng thái/kỳ.
- **Chấm công & Tính lương**: chọn kỳ (mặc định tháng trước nếu trước ngày 20); thanh 6 bước (xong / hiện tại / chờ, tên và giờ người xác nhận); nút "chuyển bước tiếp" và "Trả lại" chỉ hiện cho người phụ trách; hiện lý do trả lại gần nhất. Tab: Chấm công & công nợ (một lưới, tiêu đề nhóm TCHC / KT lương / KT công nợ, chỉ mở ô của bước hiện tại; Điền đủ công; mẫu/nhập Excel chỉ đổ vào cột đang mở), Vi phạm – quy chế (ghi theo danh mục, bảng số thực khấu trừ), Bảng lương (xuất Excel 2 sheet có dòng trạng thái và chữ ký; in bảng lương A4 ngang có 4 ô ký Người lập / KT công nợ / TCHC / Giám đốc lấy từ PheDuyet, mã kiểm tra, chữ "BẢN NHÁP – CHƯA PHÊ DUYỆT" khi chưa duyệt; in phiếu lương), Tổng hợp chi phí, Lịch sử phê duyệt.
- **Phiếu lương**: thu nhập chi tiết từng phụ cấp, khấu trừ, thực lĩnh, **số tiền bằng chữ**, ô ký; lũy kế năm.
- **Ý kiến CBCNV** (quản lý): lọc trạng thái (mặc định chưa xử lý xong), bộ phận (mặc định bộ phận mình), kỳ; mở dạng hội thoại (CBCNV bên trái, bộ phận bên phải), nút Xem phiếu lương, ô trả lời + trạng thái + chuyển bộ phận; số trên menu = ý kiến mở của bộ phận. **Phiếu lương của tôi**: khung hội thoại dưới phiếu, nút gửi ý kiến (chọn chủ đề hiện bộ phận nhận), ô bổ sung / mở lại; số trên menu = phản hồi mới chưa đọc (theo dõi trong localStorage theo từng tài khoản).
- **Nghỉ phép của tôi** (4 thẻ: phép năm còn lại / việc riêng / ốm / không lương trong năm; tạo đơn có xem trước số ngày và phép còn; danh sách đơn có nút Hủy). **Nghỉ phép CBCNV** (quản lý; số trên menu = đơn mình được duyệt): tab Chờ duyệt, Tất cả đơn (lọc tháng/loại/trạng thái), Lịch nghỉ tháng (ma trận người × ngày, ký hiệu P/R/Ô/K màu, chờ duyệt viền đứt, Chủ nhật/lễ tô xám), Quỹ phép năm; TCHC tạo đơn thay. Hồ sơ có tab Nghỉ phép. Chấm công có cột Nghỉ ốm và nút "Lấy công từ đơn nghỉ". Tổng quan hiện người nghỉ hôm nay.
- **Tài khoản**: danh sách, tạo/sửa/khóa/đặt lại mật khẩu/xóa, cấp hàng loạt (username = mã NV viết thường), bảng ma trận phân quyền.
- **Cài đặt**: toàn bộ tham số mục 3, bảng sửa danh mục phụ cấp và thang lương.
- **Nhật ký**: 300 thao tác gần nhất.

Hộp thoại xác nhận tự dựng (không dùng `confirm()`), toast thông báo, thanh tiến trình khi gọi API, phiên và dữ liệu lưu `localStorage` để mở lại nhanh rồi đồng bộ nền.

## 8. Kiểm tra trước khi giao

**Kiểm thử không phụ thuộc giờ chạy:** mọi kịch bản giao diện đụng tới báo cáo định kỳ, nhắc hẹn, hạn, chấm công phải cố định đồng hồ trình duyệt (`page.clock.install({ time: <hôm nay 10:00> })`, cho đổi bằng biến `GIO=hh:mm` để thử giờ khác) — báo cáo ngày mẫu mở lúc 07:00 nên chạy trước 7 giờ sáng sẽ "lỗi" giả. Ngày "hôm nay" trong kịch bản lấy theo giờ máy (`getFullYear/getMonth/getDate`), **không** dùng `toISOString()` (giờ UTC, lệch ngày từ 0h–7h giờ Việt Nam). Kịch bản phải báo lỗi rõ khi không tìm thấy dữ liệu cần kiểm, không được âm thầm bỏ qua. Bộ máy chủ (t*.js) dùng `NOW` cố định trong môi trường giả lập.

1. `node --check` cho từng khối script và `Code.gs`.
2. Đối chiếu mọi `onclick/onchange/oninput` với hàm đã định nghĩa.
3. Kiểm thử lõi bằng Node: ví dụ chuẩn ở mục 4, áp trần BH lương cao, OT, thử việc 10%, quy trình đầy đủ (tạo NV → quá trình → lên lương → thưởng → chấm công → tính → chốt → duyệt), các phép thử phân quyền (HR không duyệt, KT không duyệt bảng lương, NV chỉ thấy dữ liệu của mình, kỳ khóa không sửa).
4. Smoke test Playwright/Puppeteer: đăng nhập dùng thử, đi qua mọi màn hình, chạy đủ 6 bước bằng 4 tài khoản khác nhau, ghi vi phạm, phê duyệt, kiểm tra CBCNV chỉ thấy phiếu đã duyệt, mở phiếu, chụp màn hình desktop 1366 px và mobile 390 px, không lỗi console, không tràn ngang.
5. Gói giao: `index.html, Code.gs, manifest.webmanifest, sw.js, icon-192.png, icon-512.png, HUONG-DAN.md, PROMPT-TAO-LAI-APP-NHAN-SU-TAPACO.md`.

## 5v. v3.2.1 – Chống lỗi camera, định vị, bản đồ
- Nhúng CSS Leaflet 1.9.4 vào app (chỉ tải leaflet.min.js từ cdnjs). Nền bản đồ lỗi (`tileerror` ≥ 2, chưa ô nào tải được) → lớp `.bd-khongnen` (lưới) + ghi chú; ghim, vòng bán kính, tọa độ vẫn hoạt động. Dòng "Đang đặt: tên · lat, lng · bán kính" dưới bản đồ khai báo điểm. Nút "Xem trên Google Maps" luôn lấy tọa độ mới nhất.
- `quyenKhoa(f)` dùng `document.permissionsPolicy / featurePolicy.allowsFeature`; `trinhDuyetTrongApp()` nhận Zalo, Facebook/Messenger, Instagram, WeChat, LINE, TikTok; `loiMoiTruong(loai)` trả câu giải thích tiếng Việt.
- Camera lỗi: thay khung video bằng hướng dẫn theo lỗi (NotAllowed, NotReadable, NotFound, không https) + nút Thử lại; video ẩn tới khi phát được. Bản dùng thử: tự chọn mô phỏng khi khung nhúng chặn quyền; camera lỗi thì dùng ảnh mẫu.

## 5w. v3.3 – Kiểm tra vị trí khi nhắn tin
- Sau mỗi `chat.send` thành công, client gọi `vtSauTin(htId)`: chỉ khi `cfg.viTriTN.bat`, có maNV, chức danh không thuộc `boQuaVaiTro`, trong giờ làm (nếu `chiGioLam`), quá `chuKyPhut` từ lần đo trước (localStorage). Lần đầu: hộp "Kiểm tra vị trí khi nhắn tin" → "Đã hiểu" (`vitri.xacnhan`). Lấy vị trí → `vitri.bao {lat,lng,doChinhXac,tuoi,ht}`; ngoài phạm vi thì toast cho chính người dùng.
- Máy chủ `vitri.bao`: kiểm tra bỏ qua (tắt, vai trò, ngày nghỉ / lễ / ngoài giờ, đơn nghỉ đã duyệt, công tác linh hoạt, sai số > 100 m, vị trí cũ > 30′, chu kỳ qua `env.tam`), so với `gpsDiemCua`; ngoài (bán kính + max(dung sai, sai số)) → ghi `KiemTraCong` (id, ngayGhi, gioGhi, hoTen, maNV, viTriChamCong, viTriNgoaiPhamVi, khoangCach, gioDo, …), `bao()` tới r:tchc, r:giamdoc, r:ktluong, r:ktcongno, gộp đợt 60′ / 500 m; `log('vitri.ngoaivung')`.
- `vitri.ds` (READ_ONLY, chỉ CCKT_ROLES), `vitri.caidat` (TCHC/GĐ/QT; chuKyPhut 5–240, doChinhXacToiDa 20–500, dungSai 0–1000, baoLaiPhut 10–1440). Tab gpsql 'vt' "📍 Ngoài phạm vi": KPI, bản đồ (vòng điểm, chấm đỏ, nét đứt), danh sách, Xuất Excel, link sheet (`env.sheetUrl(ten)` = URL bảng tính + `#gid=`).

## 5x. v3.4 – Đánh giá tổng hợp CBCNV (AI)
- Sheet `DanhGiaTH` (id, tieuDe, tu, den, pham, taoBoi, taoLuc, soNguoi, thongKe, aiTong, aiLuc, docId, docUrl, pdfId, trangThai) và `DanhGiaCN` (id, dotId, maNV, hoTen, boPhan, chucDanh, username, chiSo, diemCC, diemTT, diemTN, diemTong, xepLoaiHT, aiXepLoai, aiTongQuan, aiTichCuc, aiHanChe, aiThieuSot, aiDeXuat, aiLuc, xepLoai, ketLuan, nguoiChot, chotLuc, congBo, nvPhanHoi, phLuc, capNhat).
- Quyền DGT_ROLES = admin, giamdoc, tchc. Lệnh: dgt.chay, dgt.ds (READ_ONLY), dgt.xem (READ_ONLY), dgt.ai (READ_ONLY, ≤ 6 người / lần hoặc tong), dgt.luuai, dgt.chot, dgt.cuatoi (READ_ONLY, nhân viên), dgt.phanhoi, dgt.xoa, dgt.baocao (khối taoDoc: dau, h1, giua, nho, h2, bang, h3, p, li, ky).
- Tính điểm như HUONG-DAN 4w; chấm công chỉ tính từ ngày đầu tiên có chấm GPS của từng người; bỏ tài khoản GĐ / QT; điểm tổng cần ≥ 2 tiêu chí. AI (DGT_HE) chỉ nhận chỉ số JSON, không nhận nội dung tin nhắn.
- Giao diện view `dgt`: lập đợt, KPI, phân bố xếp loại (lọc), nhận xét chung AI, bảng bộ phận, thẻ từng người (3 thanh điểm, tổng, xếp loại chốt / AI / theo điểm), modal chi tiết + chốt; AI chạy theo lô 5 người có tiến độ; Xuất Excel 26 cột; Báo cáo Word/PDF. "Hồ sơ của tôi": mục "Đánh giá tổng hợp của tôi" + phản hồi.
