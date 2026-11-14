# Báo cáo kiểm thử – Module Nhân sự hưởng lương hưu (v3.29)

Sinh tự động từ `test/t53.js`: 22 ca, đạt 22.

| # | Tình huống | Đầu vào | Quy tắc pháp lý | Công thức | Kết quả kỳ vọng | Kết quả thực tế | Sai lệch | Trạng thái |
|---|---|---|---|---|---|---|---|---|
| 01 | Đang hưởng lương hưu, HĐLĐ 12 tháng | Lương hưu 12 tr; lương HĐLĐ 20 tr; căn cứ trả thêm 20 tr; không NPT; cư trú | BLLĐ Đ168.3; Luật BHXH 2024 Đ2; biểu 5 bậc | trả thêm 20 tr × 21,5%; TNTT = 24,3 − 15,5 = 8,8 tr × 5% | [20000000, 4300000, 24300000, 0, 15500000, 8800000, 440000, 23860000, 12000000, 35860000, 24300000] | [20000000, 4300000, 24300000, 0, 15500000, 8800000, 440000, 23860000, 12000000, 35860000, 24300000] | 0 | ĐẠT |
| 01b | Chi tiết bậc thuế calculateRetireePIT | như Test 01 | Lũy tiến từng phần | bậc 1: 8.800.000 × 5% | [[1, 8800000, 5, 440000]] | [[1, 8800000, 5, 440000]] | 1 | ĐẠT |
| 01c | Khoản trả thêm có căn cứ, tỷ lệ, ngày hiệu lực, căn cứ pháp lý | như Test 01 | Lưu căn cứ tính | lưu trong BangLuong.traThayBHCT | [20000000, 17.5, 3, 1, 4300000, "2026-07", "BLLD_168"] | [20000000, 17.5, 3, 1, 4300000, "2026-07", "BLLD_168"] | 0 | ĐẠT |
| 02 | Hưởng lương hưu + 1 người phụ thuộc | Test 01 + 1 NPT | Giảm trừ 15,5 + 6,2 tr | 24,3 − 21,7 = 2,6 tr × 5% | [21700000, 2600000, 130000, 24170000] | [21700000, 2600000, 130000, 24170000] | 0 | ĐẠT |
| 03 | Lương TAPACO thấp, không phát sinh thuế | Lương 10 tr + trả thêm 2,15 tr; HĐLĐ 12 tháng | Lũy tiến; HĐLĐ ≥ 3 tháng không thuộc diện cam kết 08 | 12,15 tr < 15,5 tr → thuế 0 | [2150000, 12150000, 0, 0, "luytien"] | [2150000, 12150000, 0, 0, "luytien"] | 0 | ĐẠT |
| 04 | Đủ tuổi nghỉ hưu nhưng chưa hưởng lương hưu | tình trạng "Đã đủ tuổi nhưng chưa hưởng" | Không dùng tuổi làm căn cứ miễn đóng BH | BH theo HĐLĐ: 20 tr × 10,5%; không trả thêm | [1600000, 300000, 200000, 0, 3500000] | [1600000, 300000, 200000, 0, 3500000] | 0 | ĐẠT |
| 04b | Đã nhận BHXH một lần, không hưởng lương hưu | tình trạng "Hưởng trợ cấp khác / BHXH 1 lần" | Đánh giá BHXH theo HĐLĐ thực tế | BH NLĐ 10,5% | [2100000, 0] | [2100000, 0] | 0 | ĐẠT |
| 05 | HĐLĐ ngắn hạn dưới 3 tháng | HĐ dưới 3 tháng, lương 12 tr + trả thêm 2,58 tr | NĐ 253/2026 Đ50.2: khấu trừ 10% khi trả ≥ 5 tr; cam kết chỉ khi đủ điều kiện, có hồ sơ | 10% × 14.580.000; chưa có bản cam kết duyệt → không "Đủ điều kiện" | [2580000, 1458000, true] | [2580000, 1458000, true] | 0 | ĐẠT |
| 05b | Ghi chú hợp đồng dưới 3 tháng |  | Ghi chú tự động |  | true | true | 0 | ĐẠT |
| 06 | Hưởng lương hưu, có thêm HĐLĐ ở công ty khác | giảm trừ bản thân đăng ký ở Công ty X; có lương HĐLĐ nơi khác | Không giảm trừ 2 nơi; tự quyết toán | 24,3 tr × biểu 5 bậc, không giảm trừ | [0, 1930000, "khongdu"] | [0, 1930000, "khongdu"] | 0 | ĐẠT |
| 07 | Bắt đầu hưởng lương hưu giữa năm (01/7/2026) | kỳ 06/2026 và 07/2026 | Tách giai đoạn trước / sau khi hưởng | kỳ 06: BH 10,5%, không trả thêm; kỳ 07: BH 0, trả thêm 4,3 tr | [2100000, 0, 0, 4300000] | [2100000, 0, 0, 4300000] | 0 | ĐẠT |
| 08 | Tổng chi trả 20 tr đã gồm tiền thay BH (phương án B) | tổng thỏa thuận 20 tr, không khai căn cứ riêng | Không cộng 21,5% lần hai; tách lương công việc + trả thêm | S + 21,5% × S = 20 tr | [20000000, true, true] | [20000000, true, true] | 0 | ĐẠT |
| 09 | Chưa rõ lương HĐ đã gồm tiền thay BH | phương án trống, Cài đặt trả thêm tắt | Cảnh báo yêu cầu xác nhận, chưa tính khoản trả thêm |  | [0, true] | [0, true] | 0 | ĐẠT |
| 10 | Khai đang hưởng nhưng chưa xác minh, chưa có QĐ | trạng thái xác minh "Chờ" | Chưa đủ căn cứ – kế toán kiểm tra trước khi khóa bảng lương |  | [0, true, 1] | [0, true, 1] | 0 | ĐẠT |
| 11 | Lương hưu không làm mất quyền ủy quyền; khoản 3 Điều 51 chưa xác định | HĐLĐ 12 tháng, không thu nhập nơi khác | NĐ 253 Đ51 | chưa trả lời câu hỏi khoản 3 Đ51 → cần bổ sung | ["bosung", true] | ["bosung", true] | 0 | ĐẠT |
| 12 | Đủ điều kiện ủy quyền khi khai đủ | như 11 + không thuộc khoản 3 Đ51 | NĐ 253 Đ51 |  | ["du", true] | ["du", true] | 0 | ĐẠT |
| 13 | Phân tích tối ưu thuế cả năm | 12 kỳ như Test 01 (01–06 biểu 7 bậc, 07–12 biểu 5 bậc) | A = đã khấu trừ cả năm; B = thuế năm; NĐ 253 Đ70 | B = biểu năm(24,3×12 − 186 tr); nộp thừa = max(0, A − B) | [5280000, 1140000, true] | [5280000, 1140000, true] | 0 | ĐẠT |
| 14 | Không giảm trừ BH giả định 10,5% cho người hưởng lương hưu đã xác minh | Test 01 |  |  | [0, 0] | [0, 0] | 0 | ĐẠT |
| 15 | Không tự xác minh qua form; xác minh cần hồ sơ chứng minh | form ghi "Đã xác minh"; chưa tải QĐ | Kiểm tra trước khi áp dụng |  | ["Chờ", false] | ["Chờ", false] | 0 | ĐẠT |
| 16 | Có quyết định hưởng lương hưu → xác minh; lưu người, ngày; thư mục NhanSu/mã/HoSoHuuTri/năm | tải 02_QD_huong_luong_huu |  |  | [true, "Đã xác minh", "QT", "NhanSu/NV0001/HoSoHuuTri/2026/02_QD_huong_luong_huu"] | [true, "Đã xác minh", "QT", "NhanSu/NV0001/HoSoHuuTri/2026/02_QD_huong_luong_huu"] | 0 | ĐẠT |
| 17 | Sửa số QĐ sau khi xác minh → về "Chờ" |  | Dữ liệu cốt lõi thay đổi phải xác minh lại |  | Chờ | Chờ | 0 | ĐẠT |
| 18 | Khóa bảng lương bị chặn khi còn người chưa xác minh | bảng lương kỳ 09/2026 có NV0001 chưa xác minh |  |  | 1 | 1 | 0 | ĐẠT |
