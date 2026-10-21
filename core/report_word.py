# -*- coding: utf-8 -*-
"""Báo cáo trình ký (Word) kết quả rà soát hồ sơ khối lượng – Phòng Kỹ thuật TEXO.
Khổ A4, Times New Roman 13, lề 20/20/30/20 mm, có tiêu đề công ty, Quốc hiệu, khối ký.
"""
import datetime

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

FONT = "Times New Roman"
NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x59, 0x59, 0x59)
ICON = {"Cao": "🔴", "Trung bình": "🟠", "Thấp": "🟢"}
CITY = "Hà Nội"


def _font(run, size=13, bold=False, italic=False, color=None, emoji=False):
    run.font.name = "Segoe UI Emoji" if emoji else FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Segoe UI Emoji" if emoji else FONT)


def _p(container, text="", size=13, bold=False, italic=False, align=None, before=0, after=6,
       indent_first=None, color=None, line=1.15):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if indent_first is not None:
        pf.first_line_indent = Mm(indent_first)
    if align is not None:
        p.alignment = align
    if text:
        _font(p.add_run(text), size, bold, italic, color)
    return p


def _no_border(table):
    tbl_pr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}")
        x.set(qn("w:val"), "nil")
        b.append(x)
    tbl_pr.append(b)


def _cell_text(cell, text, size=13, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, underline=False):
    p = cell.paragraphs[0] if not cell.paragraphs[0].text else cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    _font(r, size, bold, italic)
    r.font.underline = underline
    return p


def _shade(cell, hex_):
    tcpr = cell._tc.get_or_add_tcPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear")
    s.set(qn("w:color"), "auto")
    s.set(qn("w:fill"), hex_)
    tcpr.append(s)


def _fixed(table, widths_mm):
    """Cố định độ rộng cột (Word và LibreOffice đều tôn trọng)."""
    table.autofit = False
    tbl = table._tbl
    tblpr = tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblpr.append(lay)
    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_mm):
        gc.set(qn("w:w"), str(int(Mm(w).twips)))
    for row in table.rows:
        for c, w in zip(row.cells, widths_mm):
            c.width = Mm(w)


def _page_field(par):
    r = par.add_run()
    a = OxmlElement("w:fldChar"); a.set(qn("w:fldCharType"), "begin")
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = "PAGE"
    z = OxmlElement("w:fldChar"); z.set(qn("w:fldCharType"), "end")
    r._r.append(a); r._r.append(i); r._r.append(z)
    _font(r, 10, color=GREY)


def build(data, out, so_hieu="……/BC-PKT"):
    d = None
    try:
        d = datetime.date.fromisoformat(data.get("ngay", ""))
    except Exception:
        d = datetime.date.today()
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(20), Mm(20)
    sec.left_margin, sec.right_margin = Mm(30), Mm(20)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(13)

    # Footer: số trang
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _font(fp.add_run("Phòng Kỹ thuật – Kết quả rà soát hồ sơ khối lượng · Trang "), 10, color=GREY)
    _page_field(fp)

    # Đầu văn bản: công ty – Quốc hiệu
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _no_border(t)
    _fixed(t, [70, 90])
    L, R = t.rows[0].cells
    _cell_text(L, "CÔNG TY CỔ PHẦN TEXO", 12)
    _cell_text(L, "TƯ VẤN VÀ ĐẦU TƯ", 12)
    _cell_text(L, "PHÒNG KỸ THUẬT", 13, bold=True, underline=True, after=6)
    _cell_text(L, f"Số: {so_hieu}", 13)
    _cell_text(R, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 12, bold=True)
    _cell_text(R, "Độc lập - Tự do - Hạnh phúc", 13, bold=True, underline=True, after=6)
    _cell_text(R, f"{CITY}, ngày {d.day:02d} tháng {d.month:02d} năm {d.year}", 13, italic=True)

    _p(doc, "BÁO CÁO", 14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=18, after=0)
    loai = (data.get("loai_ho_so") or "").strip()
    sub = "Kết quả rà soát hồ sơ khối lượng" + (f" {loai.lower()}" if loai else "")
    sub += f" – {data.get('du_an', '')}" if data.get("du_an") else ""
    if data.get("goi_thau"):
        sub += f", gói thầu {data['goi_thau']}"
    _p(doc, sub, 14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=12)

    diem = data.get("diem", [])
    cao = [x for x in diem if x.get("muc") == "Cao"]
    tb = [x for x in diem if x.get("muc") == "Trung bình"]
    thap = [x for x in diem if x.get("muc") == "Thấp"]
    J = WD_ALIGN_PARAGRAPH.JUSTIFY

    extras = [x for x in (data.get("hop_dong"), data.get("nguon_von"), data.get("tinh_chat")) if x]
    mo = (f"Phòng Kỹ thuật đã rà soát hồ sơ khối lượng \"{data.get('ho_so', '')}\""
          + (f" ({', '.join(extras).lower()})" if extras else "")
          + ". Cách làm: quét bảng tổng hợp theo bộ quy tắc của Phòng, tập trung vào nhóm hạng mục chiếm tỷ trọng lớn "
            "và các dấu hiệu bất thường; không đo bóc lại toàn bộ, không soát hình thức biểu mẫu. Kết quả như sau:")
    _p(doc, mo, align=J, indent_first=10)

    _p(doc, "1. Kết quả chung.", bold=True, after=3)
    _p(doc, f"Đánh giá chung: mức {data.get('danh_gia_chung', '')}. Tổng số {len(diem)} điểm cần làm rõ, gồm "
            f"{len(cao)} điểm mức Cao, {len(tb)} điểm mức Trung bình, {len(thap)} điểm mức Thấp.",
       align=J, indent_first=10)
    if data.get("nhan_xet_chung"):
        _p(doc, data["nhan_xet_chung"], align=J, indent_first=10)
    if data.get("nhom_chu_yeu") and data["nhom_chu_yeu"] != "–" and "chủ yếu" not in (data.get("nhan_xet_chung") or ""):
        _p(doc, f"Nhóm hạng mục chủ yếu: {data['nhom_chu_yeu']} giá trị.", align=J, indent_first=10)

    _p(doc, "2. Các điểm cần làm rõ trước khi ký.", bold=True, after=3)
    top = (cao + tb)[:8] or diem[:5]
    for i, x in enumerate(top, 1):
        p = _p(doc, "", align=J, indent_first=10, after=3)
        _font(p.add_run(ICON.get(x.get("muc"), "") + " "), 11, emoji=True)
        _font(p.add_run(f"{x.get('hang_muc', '')} ({x.get('vi_tri', '')}): "), 13, bold=True)
        _font(p.add_run(f"{x.get('van_de', '')} {x.get('cau_hoi', '')}".strip()), 13)
    if len(cao + tb) > len(top):
        _p(doc, f"Các điểm còn lại xem tại Phụ lục và file Excel kèm theo.", italic=True, indent_first=10)

    _p(doc, "3. Hạng mục thiếu hồ sơ.", bold=True, after=3)
    thieu = data.get("thieu_ho_so") or []
    if thieu:
        for x in thieu:
            _p(doc, f"- {x.get('hang_muc', '')}: còn thiếu {x.get('thieu', '')}"
                    + (f"; đơn vị bổ sung: {x.get('don_vi_bo_sung')}" if x.get("don_vi_bo_sung") else "") + ".",
               align=J, indent_first=10, after=3)
    else:
        _p(doc, "Chưa ghi nhận hạng mục thiếu hồ sơ trong phạm vi rà soát.", align=J, indent_first=10)

    _p(doc, "4. Đề nghị.", bold=True, after=3)
    can = data.get("can_bo_sung") or []
    de_nghi = []
    if cao:
        de_nghi.append(f"Nhà thầu và đoàn tư vấn giám sát làm rõ {len(cao)} điểm mức Cao trước khi trình ký.")
    if tb:
        de_nghi.append(f"Người lập, người kiểm trả lời {len(tb)} điểm mức Trung bình theo phiếu giao việc kèm theo.")
    if can:
        de_nghi.append("Bổ sung tài liệu: " + "; ".join(can) + ".")
    if not de_nghi:
        de_nghi.append("Hồ sơ đủ điều kiện trình ký trong phạm vi rà soát.")
    for x in de_nghi:
        _p(doc, "- " + x, align=J, indent_first=10, after=3)
    _p(doc, "Phòng Kỹ thuật báo cáo để Lãnh đạo xem xét, chỉ đạo.", align=J, indent_first=10, before=3)

    # Khối ký: Nơi nhận | người ký
    ben = data.get("chu_ky_ben") or [["NGƯỜI RÀ SOÁT", data.get("nguoi_lap", "")], ["TRƯỞNG PHÒNG KỸ THUẬT", "Hoàng Đức Vũ"]]
    sig = doc.add_table(rows=1, cols=1 + len(ben))
    _no_border(sig)
    _fixed(sig, [60] + [100 // len(ben)] * len(ben))
    cells = sig.rows[0].cells
    _cell_text(cells[0], "Nơi nhận:", 12, bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    for nn in (data.get("noi_nhan") or ["Ban Tổng Giám đốc (để b/c);", "Trung tâm/Đoàn TVGS (để t/h);", "Lưu: P.KT."]):
        _cell_text(cells[0], "- " + nn, 11, align=WD_ALIGN_PARAGRAPH.LEFT)
    for (title, name), c in zip(ben, cells[1:]):
        _cell_text(c, title, 12, bold=True)
        _cell_text(c, "(Ký, ghi rõ họ tên)", 11, italic=True, after=54)
        _cell_text(c, name or "", 13, bold=True)

    # Phụ lục: bảng điểm Cao + Trung bình (trang ngang)
    rows = cao + tb
    if rows:
        new = doc.add_section()
        new.orientation = WD_ORIENT.LANDSCAPE
        new.page_width, new.page_height = Mm(297), Mm(210)
        new.top_margin, new.bottom_margin = Mm(20), Mm(15)
        new.left_margin, new.right_margin = Mm(20), Mm(15)
        _p(doc, "PHỤ LỤC – CÁC ĐIỂM CẦN LÀM RÕ (MỨC CAO VÀ TRUNG BÌNH)", 13, bold=True,
           align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, "(Chi tiết tại file Excel kèm theo)", 12, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
        heads = ["TT", "Mức", "Vị trí", "Hạng mục", "Vấn đề", "Đề nghị làm rõ", "Quy tắc"]
        widths = [11, 22, 32, 48, 72, 62, 15]
        tb_ = doc.add_table(rows=1, cols=len(heads))
        tb_.style = "Table Grid"
        tb_.alignment = WD_TABLE_ALIGNMENT.CENTER
        _fixed(tb_, widths)
        for j, h in enumerate(heads):
            c = tb_.rows[0].cells[j]
            c.width = Mm(widths[j])
            _cell_text(c, h, 11, bold=True)
            _shade(c, "1F3864")
            for r in c.paragraphs[0].runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        trpr = tb_.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trpr.append(th)
        for i, x in enumerate(rows, 1):
            cells = tb_.add_row().cells
            for c, w in zip(cells, widths):
                c.width = Mm(w)
            vals = [str(i), f"{ICON.get(x.get('muc'), '')} {x.get('muc', '')}", x.get("vi_tri", ""),
                    x.get("hang_muc", ""), x.get("van_de", ""), x.get("cau_hoi", ""), x.get("ma_quy_tac", "")]
            for j, v in enumerate(vals):
                cells[j].width = Mm(widths[j])
                p = cells[j].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in (0, 1, 6) else WD_ALIGN_PARAGRAPH.LEFT
                if j == 1:
                    _font(p.add_run(ICON.get(x.get("muc"), "") + " "), 9, emoji=True)
                    _font(p.add_run(x.get("muc", "")), 10)
                else:
                    _font(p.add_run(str(v)), 10)
            if x.get("muc") == "Cao":
                _shade(cells[1], "F4CCCC")
            elif x.get("muc") == "Trung bình":
                _shade(cells[1], "FFE699")
    doc.save(out)
    return out
