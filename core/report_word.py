# -*- coding: utf-8 -*-
"""Báo cáo Word kết quả rà soát hồ sơ khối lượng – theo mẫu trình bày TEXO (skill texo-mau-trinh-bay).

Trang 1 (A4 dọc): tiêu đề, thẻ số liệu, hộp kết luận, thông tin hồ sơ, việc cần làm, khối ký.
Phụ lục (A4 ngang): bảng toàn bộ điểm cần làm rõ (mức có icon, màu), hạng mục thiếu hồ sơ.
"""
import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from . import texo_docx as K

FOOTER = "Tài liệu nội bộ · Rà soát hồ sơ khối lượng · Biên soạn/Kiểm tra: Hoàng Đức Vũ – Trưởng phòng Kỹ thuật"
DG_COLOR = {"Cao": "red", "Trung bình": "orange", "Thấp": "green"}
DG_TEXT = {"Cao": "Phải làm rõ các điểm mức Cao trước khi trình ký.",
           "Trung bình": "Không có điểm mức Cao; đề nghị người lập, người kiểm trả lời các điểm mức Trung bình trước khi ký.",
           "Thấp": "Chỉ còn các điểm lưu ý; có thể trình ký sau khi người lập xác nhận."}


def _short(s, n):
    s = (s or "").strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def build(data, out, so_hieu="……/BC-PKT"):
    try:
        d = datetime.date.fromisoformat(data.get("ngay", ""))
    except Exception:
        d = datetime.date.today()
    diem = data.get("diem", [])
    cnt = {k: sum(1 for x in diem if x.get("muc") == k) for k in ("Cao", "Trung bình", "Thấp")}
    dg = data.get("danh_gia_chung") or ("Cao" if cnt["Cao"] else "Trung bình" if cnt["Trung bình"] else "Thấp")
    du_an = data.get("du_an") or ""
    head_sub = _short((du_an + (" · " + data["goi_thau"] if data.get("goi_thau") else "")) or data.get("ho_so", ""), 48)

    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = K.FONT, Pt(12)
    content = K.setup_section(doc.sections[0], "A4P", "RÀ SOÁT HỒ SƠ KHỐI LƯỢNG", head_sub, FOOTER)

    # ---- Khối tiêu đề (không lặp tên công ty – header đã có)
    loai = (data.get("loai_ho_so") or "").strip()
    K.para(doc, "PHÒNG KỸ THUẬT · KẾT QUẢ RÀ SOÁT", 11, b=True, c=K.C["blue"], align=WD_ALIGN_PARAGRAPH.CENTER,
           before=4, after=3)
    K.para(doc, "RÀ SOÁT HỒ SƠ KHỐI LƯỢNG" + (f" {loai.upper()}" if loai else ""), 17, b=True, c=K.C["navy"],
           align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    if du_an or data.get("goi_thau"):
        K.para(doc, du_an + (f" – Gói thầu {data['goi_thau']}" if data.get("goi_thau") else ""), 12.5, b=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    K.para(doc, f"Hồ sơ: {data.get('ho_so', '')} · Ngày rà soát: {d:%d/%m/%Y}", 10.5, i=True, c=K.C["grey"],
           align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    rule = K.para(doc, "", after=8)
    K._pborder(rule, "bottom", K.C["blue"], 10, 1)

    # ---- Thẻ số liệu
    K.kpi_row(doc, [
        (str(len(diem)), "Điểm cần làm rõ", "đã chọn đưa vào báo cáo", "navy"),
        (str(cnt["Cao"]), "Mức Cao", "phải làm rõ trước khi ký", "red"),
        (str(cnt["Trung bình"]), "Mức Trung bình", "hỏi lại người lập, người kiểm", "orange"),
        (str(cnt["Thấp"]), "Mức Thấp", "lưu ý", "blue"),
    ], content)
    K.spacer(doc, 8)

    # ---- Kết luận
    body = [f"Đánh giá chung: mức {dg}. {DG_TEXT.get(dg, '')}"]
    if data.get("nhan_xet_chung"):
        body.append(data["nhan_xet_chung"])
    K.callout(doc, "KẾT LUẬN", body, content, color=DG_COLOR.get(dg, "blue"))

    # ---- Thông tin hồ sơ
    K.h1(doc, "I. THÔNG TIN HỒ SƠ")
    pairs = [("Dự án", du_an), ("Gói thầu", data.get("goi_thau", "")), ("Hồ sơ rà soát", data.get("ho_so", "")),
             ("Loại hồ sơ", loai), ("Hình thức hợp đồng", data.get("hop_dong", "")),
             ("Nguồn vốn", data.get("nguon_von", "")), ("Tính chất", data.get("tinh_chat", "")),
             ("Nhóm hạng mục chủ yếu", data.get("nhom_chu_yeu", "") if data.get("nhom_chu_yeu") != "–" else ""),
             ("Người rà soát", data.get("nguoi_lap", "")),
             ("Phạm vi, cách làm", "Quét bảng tổng hợp theo bộ quy tắc của Phòng Kỹ thuật, tập trung hạng mục tỷ trọng "
                                   "lớn và dấu hiệu bất thường; không đo bóc lại, không kiểm đơn giá, không soát hình thức.")]
    K.meta_table(doc, [(a, b) for a, b in pairs if b], content)

    # ---- Các điểm phải làm rõ (tóm tắt)
    K.h1(doc, "II. CÁC ĐIỂM CẦN LÀM RÕ TRƯỚC KHI KÝ")
    top = [x for x in diem if x.get("muc") == "Cao"][:6] or diem[:5]
    for x in top:
        p = K.para(doc, "", after=4, line=1.2)
        p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = K.Twips(360), K.Twips(-360)
        icon = K.SEV.get(x.get("muc"), ("•",))[0]
        K.run(p, icon + "  ", 10, emoji=True)
        K.run(p, f"{_short(x.get('hang_muc', ''), 70)} ", 11.5, b=True, c=K.C["navy"])
        K.run(p, f"({_short(x.get('vi_tri', ''), 60)}): ", 10.5, i=True, c=K.C["grey"])
        K.run(p, _short(x.get("van_de", ""), 260), 11.5)
    rest = len(diem) - len(top)
    if rest > 0:
        K.para(doc, f"Còn {rest} điểm khác – xem Phụ lục và file Excel kèm theo.", 10.5, i=True, c=K.C["grey"], after=4)

    # ---- Đề nghị
    K.h1(doc, "III. ĐỀ NGHỊ")
    if cnt["Cao"]:
        K.bullet(doc, f"Nhà thầu và đoàn tư vấn giám sát làm rõ {cnt['Cao']} điểm mức Cao, gửi lại hồ sơ điều chỉnh "
                      "hoặc văn bản giải trình trước khi trình ký.")
    if cnt["Trung bình"]:
        K.bullet(doc, f"Người lập, người kiểm trả lời {cnt['Trung bình']} điểm mức Trung bình theo phiếu giao việc "
                      "(sheet Giao việc trong file Excel).")
    for x in data.get("can_bo_sung") or []:
        K.bullet(doc, f"Bổ sung: {x}.")
    thieu = data.get("thieu_ho_so") or []
    if thieu:
        K.bullet(doc, f"{len(thieu)} hạng mục còn thiếu hồ sơ – danh sách tại Phụ lục 2.")
    if not (cnt["Cao"] or cnt["Trung bình"] or data.get("can_bo_sung") or thieu):
        K.bullet(doc, "Hồ sơ đủ điều kiện trình ký trong phạm vi rà soát.")
    K.spacer(doc, 8)

    ben = data.get("chu_ky_ben") or [["NGƯỜI RÀ SOÁT", data.get("nguoi_lap", "")],
                                     ["TRƯỞNG PHÒNG KỸ THUẬT", "Hoàng Đức Vũ"]]
    K.signature_block(doc, content, d, ben,
                      data.get("noi_nhan") or ["Ban Tổng Giám đốc (để b/c);", "Trung tâm, đoàn TVGS (để t/h);", "Lưu: P.KT."])

    # ---- Phụ lục 1: bảng điểm (A4 ngang)
    if diem:
        cw = K.add_section(doc, "A4L", "RÀ SOÁT HỒ SƠ KHỐI LƯỢNG", "Phụ lục – các điểm cần làm rõ", FOOTER)
        K.para(doc, "PHỤ LỤC 1 – DANH SÁCH CÁC ĐIỂM CẦN LÀM RÕ", 13.5, b=True, c=K.C["navy"],
               align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
        K.para(doc, "Xếp theo mức rủi ro và tỷ trọng; chi tiết và phiếu giao việc tại file Excel kèm theo.", 10.5,
               i=True, c=K.C["grey"], align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
        widths = [520, 1750, 1800, 2450, 3250, 3150, 1038]
        widths[-1] += cw - sum(widths)
        rows = []
        for i, x in enumerate(diem, 1):
            kl = x.get("khoi_luong") or ""
            tt = x.get("ty_trong")
            hm = [{"text": x.get("hang_muc", ""), "b": True}]
            if kl or tt:
                hm.append({"text": kl + (f" · {tt:.1f}%" if isinstance(tt, (int, float)) else ""), "i": True,
                           "c": K.C["grey"], "size": 9})
            rows.append([str(i), x.get("muc", ""), _short(x.get("vi_tri", ""), 90), hm, x.get("van_de", ""),
                         x.get("cau_hoi", ""), x.get("ma_quy_tac", "")])
        K.data_table(doc, widths, ["TT", "Mức", "Vị trí", "Hạng mục · khối lượng", "Vấn đề",
                                   "Câu hỏi cho người lập / người kiểm", "Quy tắc"],
                     rows, size=9.5, sev_col=1, center_cols=(0, 6))
        if thieu:
            K.para(doc, "", after=6)
            K.para(doc, "PHỤ LỤC 2 – HẠNG MỤC THIẾU HỒ SƠ", 13, b=True, c=K.C["navy"], align=WD_ALIGN_PARAGRAPH.CENTER,
                   after=6, keep_next=True)
            w2 = [520, 3200, 1600, 3800, 1800, 3038]
            w2[-1] += cw - sum(w2)
            K.data_table(doc, w2, ["TT", "Hạng mục", "Khối lượng", "Hồ sơ còn thiếu", "Đơn vị bổ sung", "Hướng xử lý"],
                         [[str(i), x.get("hang_muc", ""), x.get("khoi_luong", ""), x.get("thieu", ""),
                           x.get("don_vi_bo_sung", ""), x.get("huong_xu_ly", "")] for i, x in enumerate(thieu, 1)],
                         size=10, center_cols=(0,))
    doc.core_properties.author = "Hoàng Đức Vũ"
    doc.core_properties.title = "Rà soát hồ sơ khối lượng"
    doc.save(out)
    return out
