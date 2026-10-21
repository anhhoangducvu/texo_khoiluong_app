## Rà soát nhanh hồ sơ khối lượng – Phòng Kỹ thuật TEXO

App là bộ lọc trước khi trình ký: đọc hồ sơ khối lượng của nhà thầu, chỉ ra 5–20 điểm cần làm rõ, xếp theo mức rủi ro và tỷ trọng. App không đo bóc lại, không kiểm đơn giá, không soát hình thức biểu mẫu. Kết quả là gợi ý để người rà soát đối chiếu hồ sơ, chưa phải kết luận.

### Cách dùng

1. Tải hồ sơ lên: Excel (.xlsx, .xlsm, .xls) hoặc PDF xuất trực tiếp từ Excel. PDF scan chưa đọc được.
2. Nếu có, tải thêm hồ sơ đợt trước để đối chiếu lũy kế kỳ trước, khối lượng hợp đồng, đơn vị.
3. Bước 1: mở từng sheet, xem app đã nhận đúng cột chưa. Sai thì chọn lại cột, dòng tiêu đề, vai trò sheet (Tổng hợp / Chi tiết / Bỏ qua).
4. Bước 2: bấm Rà soát. App chạy toàn bộ quy tắc, gom các dòng cùng lỗi, gộp nhiều lỗi trên cùng một dòng.
5. Bước 3: sửa trực tiếp trong bảng: bỏ tick điểm không cần, tick thêm điểm ở cuối bảng, đổi mức, sửa câu hỏi, điền hồ sơ còn thiếu.
6. Bước 4 (tùy chọn): có API key thì bấm AI rà soát bổ sung. AI đọc bảng tổng hợp và các điểm đã có, bổ sung tối đa 8 điểm (ghi Nguồn = AI), đề nghị đổi mức, gợi ý quy tắc mới.
7. Bước 5: tải Excel mẫu TEXO (Tổng quan, Danh sách điểm, Thiếu hồ sơ, Giao việc, Kết quả quét), báo cáo Word theo mẫu trình bày TEXO nếu tick (A4 dọc: tiêu đề, thẻ số liệu, hộp kết luận, thông tin hồ sơ, đề nghị, khối ký; phụ lục A4 ngang: bảng điểm có icon mức độ), và JSON kết quả để gửi Claude khi cần rà sâu hơn.

### Quy tắc đang kiểm tra tự động

| Nhóm | Nội dung |
|---|---|
| Bảng tổng hợp | Lũy kế vượt hợp đồng (CH-01); kỳ trước + kỳ này ≠ lũy kế (CH-02); kỳ trước khác lũy kế đã ký đợt trước (CH-02); một đợt > 60% hợp đồng (CH-03); có khối lượng nhưng không có trong hợp đồng, khối lượng hợp đồng đổi so với đợt trước (CH-04); trùng tên trong cùng phần (CH-05); đơn vị khác đợt trước (CH-06); số tròn (CH-07); dòng ẩn, sheet ẩn, liên kết ngoài, ô lỗi, số nhập tay (CH-08); dòng cộng không khớp (CH-09); tổng hợp lệch bảng chi tiết cùng mã hiệu (CH-22); khối lượng âm (CH-23); nhiều mác bê tông (BTCT-03) |
| Từ khóa | Phát sinh, tạm tính (CH-04); hao hụt, hệ số (CH-12); ẩn khuất (CH-18); đào thủ công (DAO-07); đắp (DAP-05); vận chuyển đất (VCD-01); thuê cừ (CU-07); bê tông cọc nhồi, khoan đá, bentonite (CKN-08, CKN-07, COC-04); chống thấm đài, cọc (CTN-01, CTN-03); bơm nước, biện pháp, bơm bê tông (DAT-06, BTCT-12, BTCT-06); phá dỡ, hiện trạng, gia cố (CT-01, CT-02, CT-06); chi phí chung, trọn gói, thí nghiệm (HMC-01..03); cơ điện (ME-05, ME-06); vật liệu CĐT cấp (HT-08) |
| Đơn vị | Bê tông m³; ván khuôn m², 100 m²; cốt thép tấn, kg; hoàn thiện m²; đào đắp m³ (CH-06); cửa bộ, cái, m² (CUA-03) |
| Chỉ số | Hàm lượng thép sàn, dầm, cột, vách, móng, cọc nhồi (BTCT-08); ván khuôn sàn, dầm, cột / bê tông (CH-24); sơn / trát (HT-03). Lệch quá khoảng tham khảo cộng dung sai 30% mới báo |

### Thêm quy tắc dần dần

Bộ quy tắc nằm ở `rules/quy_tac_tu_dong.json`, dùng chung với skill `texo-khoiluong-radar` (skill để file này ở `scripts/qkl_auto/rules/`).

Cách 1 – ngay trong app, tab Bộ quy tắc: chọn loại (từ khóa / đơn vị / chỉ số), app tự gợi ý mã tiếp theo (không dùng lại mã đã cấp), bấm Thêm. App thử ngay trên hồ sơ đang mở để xem bao nhiêu dòng khớp. Sau đó tải JSON về:
- chép đè `rules/quy_tac_tu_dong.json` trên GitHub để app trên Cloud dùng bản mới;
- gửi file cho Claude kèm câu "cập nhật bộ quy tắc tự động vào skill khối lượng".

Cách 2 – nhắn Claude: chuyển tin nhắn của anh em chuyên môn, nói "thêm vào skill khối lượng". Claude ghi quy tắc vào file hạng mục của skill; quy tắc nào kiểm tra được bằng máy (từ khóa, đơn vị, chỉ số) thì Claude thêm luôn vào file JSON và gửi bản mới cho app.

Quy tắc kiểu từ khóa so khớp có dấu khi tên công việc có dấu (phân biệt cốt/cột, đài/dài), tự bỏ dấu khi hồ sơ gõ không dấu.

### Chạy và triển khai

```
pip install -r requirements.txt
streamlit run app.py
```

- Streamlit Community Cloud: đẩy thư mục lên GitHub (repo riêng tư), chọn `app.py`. App luôn hỏi mật khẩu khi vào; mật khẩu mặc định `texo2026`, đổi bằng cách dán `APP_PASSWORD = "..."` vào mục Secrets (Cloud) hoặc file `.streamlit/secrets.toml` (chạy ở máy).
- Không cần LibreOffice. Có LibreOffice thì app tự tính lại công thức cho file Excel chưa lưu sẵn kết quả.
- Kiểm thử: `python3 tests/tao_ho_so_mau.py` tạo hồ sơ mẫu có cài lỗi; `python3 tests/chay_thu.py` chạy toàn bộ và xuất file vào `tests/out/`.

### File lớn

App đọc theo luồng nên chạy được hồ sơ quyết toán nặng (thử với file 17 MB, 70 sheet: đọc khoảng 15–25 giây, quét dưới 2 giây):
- tự bỏ khối tên vùng (defined names) rác – có file chứa tới 80 MB tên vùng làm treo app bản cũ;
- mỗi sheet đọc tối đa 8.000 dòng, 60 cột, dừng khi gặp 300 dòng trống liên tiếp;
- sheet ẩn mặc định Bỏ qua (vẫn báo ở CH-08); dòng ẩn chỉ xét ở bảng tổng hợp;
- nhiều bảng tổng hợp lặp cùng số liệu (THKL, bảng giá trị…) thì cùng một lỗi chỉ báo một lần, ghi chú "trùng ở sheet …".

### Bảo mật

Không nhập API key: mọi xử lý chạy trên máy chạy app, hồ sơ không gửi đi đâu. Bấm AI rà soát bổ sung: bảng tổng hợp (tối đa 160 dòng tỷ trọng lớn nhất) và các điểm đã phát hiện được gửi tới nhà cung cấp AI đã chọn. Key chỉ giữ trong phiên, trừ khi chủ động bấm Lưu key (file `.secrets.json` cạnh app, đã loại khỏi git).
