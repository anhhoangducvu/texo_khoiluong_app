# -*- coding: utf-8 -*-
"""Lớp AI tùy chọn: đưa kết quả quét theo quy tắc + bảng tổng hợp cho AI rà soát bổ sung."""
import json

from . import llm
from .text import fmt

SYSTEM = """Bạn là kỹ sư Phòng Kỹ thuật Công ty CP TEXO Tư vấn và Đầu tư, đang rà soát nhanh hồ sơ khối lượng do nhà thầu lập (đã qua tư vấn giám sát) trước khi lãnh đạo ký.
Nguyên tắc:
- Tư vấn giám sát xác nhận KHỐI LƯỢNG, không kiểm đơn giá. Thành tiền chỉ dùng để biết hạng mục nào chiếm tỷ trọng lớn.
- Không kết luận "sai" khi chưa có căn cứ; nêu dấu hiệu và câu hỏi để người lập, người kiểm trả lời.
- Không soát hình thức, biểu mẫu, chính tả. Không bịa số liệu, tên văn bản, tên người.
- Chỉ dùng số liệu có trong dữ liệu được cung cấp; dẫn đúng sheet và dòng.
- Văn phong cán bộ kỹ thuật: ngắn, cụ thể, có số liệu; dùng từ "đề nghị làm rõ", "chưa có căn cứ", "đối chiếu", "tạm chấp nhận", "bổ sung hồ sơ". Không dùng từ hoa mỹ, không emoji.
- Mức rủi ro chỉ dùng: Cao (phải làm rõ trước khi ký), Trung bình (hỏi lại người lập/người kiểm), Thấp (lưu ý). Xét cả tỷ trọng.
Trả về DUY NHẤT một JSON hợp lệ, không kèm giải thích, không rào ```."""

SCHEMA = """{
  "nhan_xet_chung": "2–3 câu: tình trạng chung của hồ sơ, điểm chính cần làm rõ",
  "danh_gia_lai": [{"stt": 1, "muc": "Cao|Trung bình|Thấp", "ly_do": "chỉ ghi khi đề nghị đổi mức của điểm do bộ quy tắc phát hiện"}],
  "diem_bo_sung": [{"muc": "Cao|Trung bình|Thấp", "vi_tri": "Sheet / dòng N", "phan": "Đất, hố móng|Cọc|Móng, tầng hầm|Bê tông cốt thép|Hoàn thiện|Cơ điện|Cải tạo|Hạ tầng|Chung",
                    "hang_muc": "...", "khoi_luong": "...", "van_de": "...", "cau_hoi": "...", "kiem_chung": "...", "ho_so_thieu": "", "ma_quy_tac": "mã trong danh mục nếu có, không có để trống"}],
  "goi_y_quy_tac_moi": [{"ten": "...", "dau_hieu": "...", "cau_hoi": "...", "muc": "Cao|Trung bình|Thấp", "phan": "..."}]
}"""


def build_prompt(info, points, analyses, rules, max_rows=160):
    L = ["THÔNG TIN HỒ SƠ:"]
    for k, lbl in (("du_an", "Dự án"), ("goi_thau", "Gói thầu"), ("ho_so", "File"), ("loai_ho_so", "Loại hồ sơ"),
                   ("hop_dong", "Hình thức hợp đồng"), ("nguon_von", "Nguồn vốn"), ("tinh_chat", "Tính chất")):
        if info.get(k):
            L.append(f"- {lbl}: {info[k]}")
    L.append("")
    L.append("CÁC ĐIỂM BỘ QUY TẮC ĐÃ PHÁT HIỆN (đánh số stt):")
    for i, p in enumerate(points, 1):
        L.append(f"{i}. [{p['muc']}] {p.get('ma_quy_tac', '')} | {p['vi_tri']} | {p['hang_muc']} | {p['khoi_luong']} | "
                 f"tỷ trọng {p.get('ty_trong') if p.get('ty_trong') is not None else '–'}% | {p['van_de']}")
    L.append("")
    L.append("BẢNG TỔNG HỢP (sheet | dòng | STT | mã | tên | đvt | HĐ | kỳ trước | kỳ này | lũy kế | tỷ trọng %):")
    rows = []
    summ = [a for a in analyses if a.role == "Tổng hợp" and a.items] or [a for a in analyses if a.items]
    for a in summ:
        for it in a.items:
            tt = (it.value / a.total_value * 100) if (a.total_value and it.value) else None
            rows.append((tt or 0, a.name, it))
    rows.sort(key=lambda x: -x[0])
    for tt, sh, it in rows[:max_rows]:
        g = it.get
        L.append(f"{sh} | {it.label} | {it.stt} | {it.ma} | {it.ten[:90]} | {it.dvt_raw} | {fmt(g('hd'))} | "
                 f"{fmt(g('truoc'))} | {fmt(g('kynay') if g('kynay') is not None else g('kl'))} | {fmt(g('luyke'))} | "
                 f"{fmt(tt, 1) if tt else '–'}")
    if len(rows) > max_rows:
        L.append(f"(còn {len(rows) - max_rows} dòng tỷ trọng nhỏ không gửi)")
    L.append("")
    L.append("DANH MỤC MÃ QUY TẮC (mã – tên):")
    seen = set()
    for r in rules.get("quy_tac", []):
        key = (r["ma"], r.get("ten"))
        if key not in seen:
            seen.add(key)
            L.append(f"- {r['ma']} – {r.get('ten', '')}")
    L.append("")
    L.append("YÊU CẦU: (1) nhận xét chung; (2) nếu thấy điểm nào của bộ quy tắc bị đánh giá sai mức thì nêu ở danh_gia_lai; "
             "(3) bổ sung tối đa 8 điểm rủi ro mà bộ quy tắc chưa bắt được, ưu tiên hạng mục tỷ trọng lớn "
             "(ví dụ: khối lượng không hợp lý giữa các công tác liên quan như đào–đắp–vận chuyển, bê tông–ván khuôn–cốt thép, "
             "xây–trát–sơn; hạng mục thiếu công tác đi kèm); (4) gợi ý tối đa 3 quy tắc mới có thể lập trình được. "
             "Không lặp lại điểm đã có.")
    L.append("Định dạng JSON:")
    L.append(SCHEMA)
    return "\n".join(L)


def review(provider, model, key, base_url, info, points, analyses, rules):
    prompt = build_prompt(info, points, analyses, rules)
    res = llm.ask_json(provider, model, key, prompt, base_url=base_url, system=SYSTEM)
    if res.get("ok"):
        d = res["data"] or {}
        for k in ("danh_gia_lai", "diem_bo_sung", "goi_y_quy_tac_moi"):
            if not isinstance(d.get(k), list):
                d[k] = []
        for p in d["diem_bo_sung"]:
            if p.get("muc") not in ("Cao", "Trung bình", "Thấp"):
                p["muc"] = "Trung bình"
            p.setdefault("ty_trong", None)
            p["nguon"] = "AI"
        res["data"] = d
    res["prompt_chars"] = len(prompt)
    return res
