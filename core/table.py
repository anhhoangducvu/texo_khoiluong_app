# -*- coding: utf-8 -*-
"""Nhận diện dòng tiêu đề, ghép cột và tách các dòng công việc từ một Sheet."""
import re
from dataclasses import dataclass, field

from .text import norm, contains_word, kw_in, num, unit, is_total_name, is_section_code

# Tên khóa cột hiển thị cho người dùng
COL_KEYS = {
    "stt": "STT",
    "ma": "Mã hiệu",
    "ten": "Tên công việc",
    "dvt": "Đơn vị",
    "hd": "KL hợp đồng/thiết kế",
    "truoc": "KL lũy kế kỳ trước",
    "kynay": "KL kỳ này (đề nghị)",
    "luyke": "KL lũy kế đến kỳ này",
    "kl": "Khối lượng (chung)",
    "dongia": "Đơn giá",
    "tt_hd": "Thành tiền hợp đồng",
    "tt_truoc": "Thành tiền kỳ trước",
    "tt_kynay": "Thành tiền kỳ này",
    "tt_luyke": "Thành tiền lũy kế",
    "thanhtien": "Thành tiền (chung)",
}
QTY_KEYS = ("hd", "truoc", "kynay", "luyke", "kl")
VAL_KEYS = ("tt_luyke", "tt_kynay", "thanhtien", "tt_hd", "tt_truoc")

P_TRUOC = ["ky truoc", "dot truoc", "lan truoc", "den het ky truoc", "den ky truoc", "luy ke truoc",
           "da thanh toan", "da nghiem thu", "giai doan truoc"]
P_LUYKE = ["luy ke", "den het ky nay", "den het dot nay", "den nay", "tong cong"]
P_KYNAY = ["ky nay", "dot nay", "lan nay", "de nghi", "thuc hien", "nghiem thu", "giai doan nay"]
P_HD = ["hop dong", "hd", "du toan", "thiet ke", "tk", "duoc duyet", "boq", "phe duyet", "goi thau", "theo hop"]
DIM_WORDS = ["dai", "rong", "cao", "so luong", "so ck", "so cau kien", "he so", "day", "dien tich", "the tich", "chu vi"]


def _has(t, words):
    return any(contains_word(t, w) for w in words)


def classify(text):
    """Trả khóa cột từ chữ tiêu đề (đã gộp nhiều dòng)."""
    t = norm(text)
    if not t:
        return None
    if re.fullmatch(r"\(?\d{1,2}\)?([=x\*\+\-/ \(\)\d]+)?", t):
        return None
    if t in ("stt", "tt", "so tt", "s tt", "no", "no."):
        return "stt"
    if _has(t, ["ghi chu"]):
        return "ghichu"
    if _has(t, ["don gia"]):
        return "dongia"
    if re.match(r"^ma\b", t) or _has(t, ["ma hieu", "ma cong viec", "ma dinh muc", "ma so", "ma dm", "ma cv"]):
        return "ma"
    if _has(t, ["don vi", "dvt", "dv tinh", "d.v.t"]) or t == "dv":
        return "dvt"
    is_val = _has(t, ["thanh tien", "gia tri", "vnd", "so tien"]) or re.search(r"\(\s*dong\s*\)", t) is not None
    period = None
    if _has(t, P_TRUOC):
        period = "truoc"
    elif _has(t, P_LUYKE):
        period = "luyke"
    elif _has(t, P_KYNAY):
        period = "kynay"
    elif _has(t, P_HD):
        period = "hd"
    if is_val:
        return "tt_" + period if period else "thanhtien"
    if _has(t, ["noi dung", "ten cong viec", "hang muc", "dien giai", "cong viec", "ten cong tac",
                "danh muc", "mo ta", "ten vat tu", "ten"]) and not period:
        return "ten"
    if period:
        return period
    if _has(t, ["khoi luong", "kl"]):
        return "kl"
    if _has(t, DIM_WORDS):
        return "dim"
    return None


@dataclass
class TableMap:
    header_row: int = None          # 0-based, dòng tiêu đề đầu tiên
    data_start: int = None          # 0-based, dòng số liệu đầu tiên
    cols: dict = field(default_factory=dict)   # khóa → chỉ số cột (0-based)
    dims: int = 0
    score: int = 0


def _col_texts(g, r, span, ncols):
    out = []
    for c in range(ncols):
        parts = []
        for rr in range(r, min(r + span, len(g))):
            v = g[rr][c] if c < len(g[rr]) else None
            if v not in (None, "") and (not parts or str(v).strip() != parts[-1]):
                parts.append(str(v).strip())
        out.append(" ".join(parts))
    return out


def detect(sheet, max_scan=40):
    """Tìm dòng tiêu đề tốt nhất (1–3 dòng gộp) và ghép cột."""
    g = sheet.head_grid
    if not g:
        return TableMap()
    ncols = max((len(r) for r in g[:max_scan + 3]), default=0)
    best = TableMap()
    filled = [sum(1 for v in row if v not in (None, "")) for row in g[:max_scan + 3]]
    for r in range(0, min(len(g), max_scan)):
        for span in (1, 2, 3):
            if r + span > len(filled) or any(filled[x] < 2 for x in range(r, r + span)):
                break
            texts = _col_texts(g, r, span, ncols)
            cols, dims, multi = {}, 0, {}
            for c, t in enumerate(texts):
                k = classify(t)
                if k == "dim":
                    dims += 1
                    continue
                if k:
                    multi.setdefault(k, []).append(c)
                if k and k not in cols:
                    cols[k] = c
            # Hai cột cùng "lũy kế" (đến kỳ trước / đến kỳ này, chữ bị cắt) → cột đầu là kỳ trước
            if len(multi.get("luyke", [])) >= 2 and "truoc" not in cols:
                cols["truoc"], cols["luyke"] = multi["luyke"][0], multi["luyke"][1]
            if "ten" not in cols:
                continue
            nq = sum(1 for k in QTY_KEYS if k in cols)
            score = len(cols) * 2 + 4 + nq * 2 + (2 if "dvt" in cols else 0) + min(dims, 3) - span
            if score > best.score:
                best = TableMap(header_row=r, data_start=r + span, cols=cols, dims=dims, score=score)
    return best


@dataclass
class Item:
    idx: int
    label: str
    ten: str
    dvt: str = ""
    dvt_raw: str = ""
    stt: str = ""
    ma: str = ""
    section: str = ""
    hidden: bool = False
    hard: bool = False
    is_total: bool = False
    vals: dict = field(default_factory=dict)
    phan: str = ""

    def get(self, k):
        return self.vals.get(k)

    @property
    def q(self):
        """Khối lượng đề nghị dùng để hiển thị: kỳ này → lũy kế → KL chung → hợp đồng."""
        for k in ("kynay", "luyke", "kl", "hd"):
            v = self.vals.get(k)
            if v not in (None, 0):
                return v
        return None

    @property
    def q_key(self):
        for k in ("kynay", "luyke", "kl", "hd"):
            if self.vals.get(k) not in (None, 0):
                return k
        return None

    @property
    def value(self):
        for k in VAL_KEYS:
            v = self.vals.get(k)
            if v:
                return v
        dg, q = self.vals.get("dongia"), (self.vals.get("luyke") or self.vals.get("kynay") or self.vals.get("kl"))
        if dg and q:
            return dg * q
        return None


PHAN_RULES = [
    ("Cải tạo", ["tháo dỡ", "phá dỡ", "đục phá", "phá bỏ", "cải tạo", "cạo bỏ"]),
    ("Đất, hố móng", ["đào", "đắp", "vận chuyển đất", "san nền", "san lấp", "bơm nước", "đổ thải", "phế thải",
                      "hạ mực nước", "đất cấp", "nạo vét"]),
    ("Cọc", ["cọc", "cừ", "larsen", "khoan nhồi", "barrette", "tường vây", "ép cừ", "nhổ cừ"]),
    ("Móng, tầng hầm", ["móng", "đài", "giằng móng", "dầm móng", "tầng hầm", "bê tông lót", "hầm", "bể nước ngầm"]),
    ("Cơ điện", ["điện", "dây cáp", "cáp điện", "ống nước", "cấp nước", "thoát nước", "pccc", "chữa cháy",
                 "điều hòa", "thông gió", "đèn", "ổ cắm", "công tắc", "thiết bị", "máy bơm", "ống ppr", "ống upvc",
                 "ống hdpe", "tủ điện", "báo cháy", "camera", "thang máy", "sprinkler", "ống gió"]),
    ("Hạ tầng", ["mặt đường", "nền đường", "đường giao thông", "đường nội bộ", "cống", "hố ga", "bê tông nhựa",
                 "cấp phối", "vỉa hè", "bó vỉa", "rãnh thoát"]),
    ("Hoàn thiện", ["xây", "trát", "sơn", "bả", "lát", "ốp", "trần", "cửa", "chống thấm", "mái", "len chân",
                    "lan can", "vách ngăn", "lưới chống nứt", "láng", "đá granite", "gạch", "vách kính", "tôn"]),
    ("Bê tông cốt thép", ["bê tông", "cốt thép", "ván khuôn", "cột", "dầm", "sàn", "vách", "cầu thang", "lanh tô",
                          "thép", "cốp pha", "xà"]),
]


def guess_phan(name, section=""):
    for txt in (name, section):
        if not txt:
            continue
        for phan, words in PHAN_RULES:
            if any(kw_in(txt, w) for w in words):
                return phan
    return "Chung"


def _cell(row, c):
    return row[c] if c is not None and c < len(row) else None


def parse(sheet, tmap):
    """Tách các dòng công việc. Trả (items, totals) – totals là các dòng Cộng/Tổng."""
    items, totals = [], []
    if tmap.data_start is None or "ten" not in tmap.cols:
        return items, totals
    cols = tmap.cols
    section = ""
    head_ten = norm(_cell(sheet.head_grid[tmap.header_row], cols["ten"])) if tmap.header_row is not None else ""
    for i in range(tmap.data_start, len(sheet.grid)):
        row = sheet.grid[i]
        name = _cell(row, cols["ten"])
        name_s = str(name).strip() if name is not None else ""
        dvt_raw = _cell(row, cols.get("dvt"))
        dvt_raw = str(dvt_raw).strip() if dvt_raw not in (None, "") else ""
        if not name_s:
            continue
        nn = norm(name_s)
        if nn == head_ten or norm(dvt_raw) in ("dvt", "don vi", "don vi tinh"):
            continue                                   # dòng tiêu đề lặp lại (PDF nhiều trang)
        if re.fullmatch(r"[\(\)\d\s\.,=+\-x*/]+", nn):
            continue                                   # dòng đánh số cột (1) (2) (3)…
        vals = {}
        for k in list(COL_KEYS):
            if k in ("stt", "ma", "ten", "dvt"):
                continue
            if k in cols:
                vals[k] = num(_cell(row, cols[k]))
        stt = _cell(row, cols.get("stt"))
        stt = "" if stt is None else str(stt).strip()
        ma = _cell(row, cols.get("ma"))
        ma = "" if ma is None else str(ma).strip()
        has_qty = any(vals.get(k) not in (None, 0) for k in QTY_KEYS)
        has_val = any(vals.get(k) not in (None, 0) for k in VAL_KEYS)
        it = Item(idx=i, label=sheet.label(i), ten=name_s[:200], dvt=unit(dvt_raw), dvt_raw=dvt_raw,
                  stt=stt, ma=ma, section=section, hidden=(i in sheet.hidden_rows), vals=vals)
        if is_total_name(name_s) and not dvt_raw:
            it.is_total = True
            totals.append(it)
            continue
        if not dvt_raw and not has_qty:
            if has_val or is_section_code(stt) or name_s.isupper() or len(name_s) < 120:
                section = name_s[:120]
            continue
        if sheet.formula is not None:
            for k in ("kynay", "luyke", "kl"):
                c = cols.get(k)
                if c is not None and vals.get(k) not in (None, 0):
                    fl = sheet.formula[i][c] if c < len(sheet.formula[i]) else False
                    if not fl:
                        it.hard = True
                    break
        it.phan = guess_phan(name_s, section)
        items.append(it)
    return items, totals


def guess_role(sheet, tmap):
    n = norm(sheet.name)
    if tmap.data_start is None:
        return "Bỏ qua"
    if _has(n, ["chi tiet", "dien giai", "do boc", "tinh toan", "ct", "dg"]) or tmap.dims >= 2:
        return "Chi tiết"
    return "Tổng hợp"
