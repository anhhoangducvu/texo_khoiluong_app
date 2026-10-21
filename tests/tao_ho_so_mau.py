# -*- coding: utf-8 -*-
"""Tạo bộ hồ sơ khối lượng mẫu (số liệu giả định) có cài sẵn lỗi để kiểm thử app.
Chạy: python3 tests/tao_ho_so_mau.py  → tests/mau/HS_KL_dot3.xlsx, HS_KL_dot2.xlsx
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

OUT = os.path.join(os.path.dirname(__file__), "mau")
os.makedirs(OUT, exist_ok=True)

# (stt, mã, tên, đvt, HĐ, kỳ trước, kỳ này, đơn giá, ẩn?, lũy kế ghi đè)
ROWS = [
    ("I", "", "PHẦN MÓNG", None, None, None, None, None, False, None),
    ("1", "AB.25112", "Đào đất móng bằng máy, đất cấp II", "m3", 1200, 800, 450, 85000, False, None),
    ("2", "AB.11312", "Đào đất móng thủ công, đất cấp II", "m3", 300, 150, 200, 320000, False, None),
    ("3", "AB.41432", "Vận chuyển đất thải đổ đi cự ly 10km", "m3", 1500, 900, 500, 95000, False, None),
    ("4", "AF.11111", "Bê tông lót móng đá 4x6 mác 100", "m3", 45, 45, 0, 1250000, False, None),
    ("5", "AF.11213", "Bê tông móng, đài móng đá 1x2 mác 300", "m3", 380, 380, 0, 1650000, False, None),
    ("6", "AF.61130", "Cốt thép móng, đài móng đường kính >18mm", "tấn", 52, 52, 0, 21500000, False, None),
    ("7", "AK.95211", "Chống thấm đài móng bằng màng khò nóng", "m2", 0, 0, 420, 185000, False, None),
    ("8", "AB.13112", "Đắp đất hố móng, độ chặt K95", "m3", 600, 200, 350, 110000, False, None),
    ("", "", "Cộng phần móng", None, None, None, None, None, False, None),
    ("II", "", "PHẦN THÂN", None, None, None, None, None, False, None),
    ("9", "AF.12313", "Bê tông cột tiết diện >0,1m2 mác 400", "m3", 160, 60, 50, 1850000, False, None),
    ("10", "AF.12513", "Bê tông xà dầm, giằng nhà mác 300", "m3", 210, 70, 60, 1720000, False, None),
    ("11", "AF.12613", "Bê tông sàn mái mác 300", "m3", 420, 150, 140, 1690000, False, None),
    ("12", "AF.61531", "Cốt thép cột, trụ đường kính >18mm", "tấn", 38, 14, 12, 21800000, False, None),
    ("13", "AF.61631", "Cốt thép xà dầm, giằng đường kính >18mm", "tấn", 42, 15, 13, 21800000, False, None),
    ("14", "AF.61711", "Cốt thép sàn mái đường kính ≤10mm", "tấn", 40, 20, 25, 22100000, False, None),
    ("15", "AF.86311", "Ván khuôn sàn mái", "100m2", 35, 12, 15, 9800000, False, None),
    ("16", "AF.82111", "Ván khuôn cột vuông, chữ nhật", "100m2", 18, 7, 6, 11200000, False, None),
    ("17", "AF.12613", "Bê tông sàn mái mác 300", "m3", 420, 0, 140, 1690000, False, None),
    ("18", "AF.12413", "Bê tông lanh tô, ô văng phát sinh", "m2", 0, 0, 12, 1700000, False, None),
    ("19", "AF.99999", "Thép biện pháp chống đỡ sàn", "tấn", 0, 0, 3.5, 18000000, True, None),
    ("", "", "Cộng phần thân", None, None, None, None, None, False, None),
    ("III", "", "PHẦN HOÀN THIỆN", None, None, None, None, None, False, None),
    ("20", "AE.22224", "Xây tường gạch 220 vữa M75", "m3", 520, 200, 200, 1450000, False, None),
    ("21", "AK.21224", "Trát tường trong dày 1,5cm vữa M75", "m2", 4200, 1000, 1500, 95000, False, 2600),
    ("22", "AK.84114", "Sơn tường trong nhà 1 nước lót 2 nước phủ", "m2", 5200, 900, 3500, 75000, False, None),
    ("23", "AK.51250", "Lát nền gạch Granite 600x600", "m2", 1800, 0, 1000, 420000, False, None),
    ("24", "AH.12100", "Lắp dựng cửa nhôm kính", "m", 260, 0, 100, 2500000, False, None),
    ("25", "AB.99999", "Tháo dỡ tường cũ hiện trạng", "m3", 0, 0, -5, 450000, False, None),
    ("", "", "Cộng phần hoàn thiện", None, None, None, None, None, False, None),
    ("", "", "TỔNG CỘNG", None, None, None, None, None, False, None),
]


def build(path, dot=3, prev=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    ws["A1"] = "BẢNG TỔNG HỢP KHỐI LƯỢNG HOÀN THÀNH ĐỀ NGHỊ THANH TOÁN ĐỢT %d" % dot
    ws["A2"] = "Công trình: Nhà làm việc A (số liệu giả định để kiểm thử)"
    h1 = ["STT", "Mã hiệu", "Nội dung công việc", "Đơn vị", "Khối lượng", None, None, None, "Đơn giá", "Thành tiền", None]
    h2 = [None, None, None, None, "Theo hợp đồng", "Lũy kế đến hết kỳ trước", "Thực hiện kỳ này",
          "Lũy kế đến hết kỳ này", None, "Kỳ này", "Lũy kế"]
    ws.append([])
    ws.append(h1)
    ws.append(h2)
    for rng in ("A4:A5", "B4:B5", "C4:C5", "D4:D5", "E4:H4", "I4:I5", "J4:K4"):
        ws.merge_cells(rng)
    ws.append([f"({i})" for i in range(1, 12)])
    r = 7
    sec_start = r
    sub_rows = []
    for stt, ma, ten, dvt, hd, tr, kn, dg, hidden, lk_override in ROWS:
        if prev and dvt:
            tr, kn = (tr or 0) * 0.6, (tr or 0) * 0.4      # đợt trước: lũy kế = kỳ trước của đợt 3
            if ma == "AF.12513":
                kn += 5                                    # lỗi cài sẵn: kỳ trước đợt 3 khác lũy kế đợt 2
        if dvt is None and ten.startswith(("PHẦN",)):
            ws.append([stt, ma, ten])
            ws.cell(r, 3).font = Font(bold=True)
            sec_start = r + 1
        elif ten.startswith("Cộng"):
            ws.append([None, None, ten, None, None, None, None, None, None,
                       f"=SUM(J{sec_start}:J{r - 1})", f"=SUM(K{sec_start}:K{r - 1})"])
            sub_rows.append(r)
            if "thân" in ten and not prev:           # lỗi cài sẵn: cộng thiếu 1 dòng
                ws.cell(r, 10).value = f"=SUM(J{sec_start}:J{r - 3})"
        elif ten == "TỔNG CỘNG":
            ws.append([None, None, ten, None, None, None, None, None, None,
                       "=" + "+".join(f"J{x}" for x in sub_rows), "=" + "+".join(f"K{x}" for x in sub_rows)])
        else:
            lk = lk_override if (lk_override and not prev) else f"=F{r}+G{r}"
            ws.append([stt, ma, ten, dvt, hd, tr, kn, lk, dg, f"=G{r}*I{r}", f"=H{r}*I{r}"])
            if hidden:
                ws.row_dimensions[r].hidden = True
        r += 1
    if not prev:
        ws["L12"] = "=#REF!+1"
    ws.column_dimensions["C"].width = 48
    thin = Side(style="thin")
    for row in ws.iter_rows(min_row=4, max_row=r - 1, max_col=11):
        for c in row:
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Sheet chi tiết
    ct = wb.create_sheet("Chi tiet KL")
    ct.append(["BẢNG DIỄN GIẢI KHỐI LƯỢNG KỲ NÀY"])
    ct.append(["STT", "Mã hiệu", "Tên công việc / Diễn giải", "Đơn vị", "Số CK", "Dài", "Rộng", "Cao", "Khối lượng"])
    ct.append([1, "AB.25112", "Đào đất móng bằng máy, đất cấp II", "m3", None, None, None, None, 450])
    ct.append([None, None, "Hố móng M1 trục A-B", None, 10, 5, 3, 1.5, 225])
    ct.append([None, None, "Hố móng M2 trục C-D", None, 10, 5, 3, 1.5, 225])
    ct.append([2, "AF.12313", "Bê tông cột tiết diện >0,1m2 mác 400", "m3", None, None, None, None, 46.2])
    ct.append([3, "AF.61711", "Cốt thép sàn mái đường kính ≤10mm", "tấn", None, None, None, None, 25])
    wb.save(path)


if __name__ == "__main__":
    build(os.path.join(OUT, "HS_KL_dot3.xlsx"), 3)
    build(os.path.join(OUT, "HS_KL_dot2.xlsx"), 2, prev=True)
    print("Đã tạo", os.listdir(OUT))
