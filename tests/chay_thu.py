# -*- coding: utf-8 -*-
"""Chạy thử toàn bộ (không cần Streamlit): đọc hồ sơ mẫu → quét → xuất Excel, Word, JSON.
Dùng: python3 tests/chay_thu.py [file_ho_so] [file_dot_truoc]
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from core import reader, engine, report_excel, report_word  # noqa: E402

MAU = os.path.join(ROOT, "tests", "mau")
OUT = os.path.join(ROOT, "tests", "out")
os.makedirs(OUT, exist_ok=True)

f = sys.argv[1] if len(sys.argv) > 1 else os.path.join(MAU, "HS_KL_dot3.xlsx")
fp = sys.argv[2] if len(sys.argv) > 2 else (os.path.join(MAU, "HS_KL_dot2.xlsx") if len(sys.argv) == 1 else None)

book = reader.load(f)
prev = reader.load(fp) if fp else None
rules = engine.load_rules()
errs = engine.validate_rules(rules)
assert not errs, errs
an, raw, grouped = engine.run_all(book, rules, prev_book=prev)
chosen, rest = engine.select(grouped, rules["cai_dat"])
for w in book.warnings:
    print("Lưu ý:", w)
for a in an:
    print(f"- {a.name}: {a.role}, {len(a.items)} dòng, cột {a.tmap.cols}")
print(f"\n{len(chosen)} điểm chọn, {len(rest)} phát hiện khác")
for x in chosen:
    print(f"  [{x.muc}] {x.ma:<14} {x.vi_tri:<28} {x.hang_muc[:40]:<40} {x.van_de[:90]}")

info = {"du_an": "Nhà làm việc A (mẫu)", "goi_thau": "XL01", "loai_ho_so": "Thanh toán",
        "hop_dong": "Đơn giá cố định", "nguon_von": "Vốn nhà nước", "tinh_chat": "Xây mới",
        "nguoi_lap": "Nguyễn Văn A", "kho_giay": "A4", "chu_ky": True}
data = engine.build_report(an, chosen, info, rules, book, rest)
stem = os.path.join(OUT, os.path.splitext(os.path.basename(f))[0])
with open(stem + "_ket_qua.json", "w", encoding="utf-8") as fh:
    json.dump(data, fh, ensure_ascii=False, indent=2)
report_excel.build(data, stem + "_ra_soat.xlsx")
report_word.build(data, stem + "_bao_cao.docx")
print("Đã xuất:", stem + "_ra_soat.xlsx,", stem + "_bao_cao.docx")
