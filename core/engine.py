# -*- coding: utf-8 -*-
"""Bộ máy rà soát nhanh theo bộ quy tắc JSON dùng chung với skill texo-khoiluong-radar.

Luồng: Book → phân tích từng sheet (nhận diện bảng, vai trò) → chạy quy tắc → danh sách phát hiện
→ chọn 5–20 điểm theo mức và tỷ trọng → JSON đúng cấu trúc qkl_report.py.
"""
import copy
import datetime
import json
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field

from .table import detect, parse, guess_role, QTY_KEYS
from .text import norm, contains_word, kw_in, kw_pos, fmt, pct, unit_label

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RULES = os.path.join(os.path.dirname(HERE), "rules", "quy_tac_tu_dong.json")
LEVELS = ("Cao", "Trung bình", "Thấp")


# --------------------------------------------------------------------------- quy tắc
def load_rules(path=DEFAULT_RULES):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def validate_rules(data):
    """Trả danh sách lỗi (rỗng nếu hợp lệ)."""
    errs = []
    if not isinstance(data, dict) or not isinstance(data.get("quy_tac"), list):
        return ["File phải có khóa 'quy_tac' là một danh sách."]
    for i, r in enumerate(data["quy_tac"], 1):
        tag = f"Quy tắc {i} ({r.get('ma', '?')})"
        if not r.get("ma"):
            errs.append(f"{tag}: thiếu 'ma'.")
        if r.get("loai") not in ("he_thong", "tu_khoa", "don_vi", "chi_so"):
            errs.append(f"{tag}: 'loai' phải là he_thong / tu_khoa / don_vi / chi_so.")
        if r.get("muc") not in LEVELS:
            errs.append(f"{tag}: 'muc' phải là Cao / Trung bình / Thấp.")
        if r.get("loai") in ("tu_khoa", "don_vi") and not r.get("tu_khoa"):
            errs.append(f"{tag}: thiếu danh sách 'tu_khoa'.")
        if r.get("loai") == "don_vi" and not r.get("don_vi_hop_le"):
            errs.append(f"{tag}: thiếu 'don_vi_hop_le'.")
        if r.get("loai") == "chi_so":
            for k in ("tu_so", "mau_so"):
                if not isinstance(r.get(k), dict) or not r[k].get("tu_khoa"):
                    errs.append(f"{tag}: thiếu '{k}.tu_khoa'.")
            if r.get("min") is None and r.get("max") is None:
                errs.append(f"{tag}: cần 'min' hoặc 'max'.")
    return errs


def next_code(rules, prefix):
    """Mã mới = số lớn nhất đã cấp của tiền tố + 1 (không dùng lại mã đã cấp)."""
    prefix = prefix.strip().upper()
    mx = int((rules.get("ma_lon_nhat") or {}).get(prefix, 0))
    for r in rules.get("quy_tac", []):
        m = re.match(rf"^{re.escape(prefix)}-(\d+)", r.get("ma", ""))
        if m:
            mx = max(mx, int(m.group(1)))
    return f"{prefix}-{mx + 1:02d}"


def register_code(rules, code):
    m = re.match(r"^([A-Z]+)-(\d+)", code or "")
    if m:
        reg = rules.setdefault("ma_lon_nhat", {})
        reg[m.group(1)] = max(int(reg.get(m.group(1), 0)), int(m.group(2)))


class _Safe(dict):
    def __missing__(self, k):
        return "{" + k + "}"


def _fill(t, **kw):
    return (t or "").format_map(_Safe(**kw))


def _match(name, spec):
    """spec: {tu_khoa, va_tu_khoa, va_tu_khoa_mot, tru_tu_khoa}. Trả từ khóa khớp hoặc None."""
    hit = None
    for k in spec.get("tu_khoa") or []:
        if kw_in(name, k):
            hit = k
            break
    if not hit:
        return None
    for k in spec.get("va_tu_khoa") or []:
        if not kw_in(name, k):
            return None
    one = spec.get("va_tu_khoa_mot") or []
    if one and not any(kw_in(name, k) for k in one):
        return None
    for k in spec.get("tru_tu_khoa") or []:
        if kw_in(name, k):
            return None
    return hit


# --------------------------------------------------------------------------- phân tích
@dataclass
class SheetAnalysis:
    name: str
    role: str
    tmap: object
    items: list = field(default_factory=list)
    totals: list = field(default_factory=list)
    total_value: float = 0.0
    hidden: bool = False
    errors: list = field(default_factory=list)
    hard_count: int = 0
    kind: str = "excel"


def analyze(book, overrides=None):
    """overrides: {tên sheet: {"role": ..., "header_row": int, "data_start": int, "cols": {...}}}"""
    overrides = overrides or {}
    plan = []
    for sh in book.sheets:
        if getattr(sh, "_tmap", None) is None:
            sh._tmap = detect(sh)                      # nhận diện một lần, lưu lại
        tmap = copy.deepcopy(sh._tmap)
        ov = overrides.get(sh.name) or {}
        if ov.get("cols") is not None:
            tmap.cols = {k: v for k, v in ov["cols"].items() if v is not None}
        if ov.get("header_row") is not None:
            tmap.header_row = ov["header_row"]
        if ov.get("data_start") is not None:
            tmap.data_start = ov["data_start"]
        # Sheet ẩn mặc định bỏ qua (thường là bản cũ, bản nháp); vẫn được báo ở CH-08
        role = ov.get("role") or ("Bỏ qua" if sh.hidden else guess_role(sh, tmap))
        plan.append((sh, tmap, role))
    if hasattr(book, "path"):
        from . import reader as _reader
        _reader.load_formulas(book, {sh.name for sh, _, role in plan if role == "Tổng hợp"})
    out = []
    for sh, tmap, role in plan:
        items, totals = parse(sh, tmap) if role != "Bỏ qua" else ([], [])
        a = SheetAnalysis(name=sh.name, role=role, tmap=tmap, items=items, totals=totals,
                          hidden=sh.hidden, errors=sh.errors, kind=sh.kind)
        a.total_value = sum(it.value for it in items if it.value and it.value > 0)
        a.hard_count = sum(1 for it in items if it.hard)
        out.append(a)
    return out


# --------------------------------------------------------------------------- phát hiện
@dataclass
class Finding:
    ma: str
    muc: str
    rule_id: str
    ten_quy_tac: str
    sheet: str
    vi_tri: str
    hang_muc: str
    khoi_luong: str
    ty_trong: float
    van_de: str
    cau_hoi: str
    kiem_chung: str
    phan: str
    trang_thai: str = ""
    ho_so_thieu: str = ""
    rows: list = field(default_factory=list)
    nguon: str = "Quy tắc"
    score: float = 0.0
    group: bool = True

    def to_point(self):
        return {"muc": self.muc, "vi_tri": self.vi_tri, "phan": self.phan or "Chung",
                "hang_muc": self.hang_muc, "khoi_luong": self.khoi_luong,
                "ty_trong": None if self.ty_trong is None else round(self.ty_trong, 1),
                "van_de": self.van_de, "cau_hoi": self.cau_hoi, "kiem_chung": self.kiem_chung,
                "ho_so_thieu": self.ho_so_thieu, "ma_quy_tac": self.ma, "nguon": self.nguon}


def _kl(it, key=None):
    key = key or it.q_key
    if not key:
        return ""
    lbl = {"kynay": "kỳ này", "luyke": "lũy kế", "kl": "", "hd": "HĐ", "truoc": "kỳ trước"}[key]
    s = f"{fmt(it.get(key))} {unit_label(it.dvt)}".strip()
    return f"{s} ({lbl})" if lbl else s


def _rows(items, n=8):
    labs = [x.label.replace("dòng ", "") for x in items]
    return ", ".join(labs[:n]) + ("…" if len(labs) > n else "")


def _tt(it, sa):
    if sa.total_value and it.value and it.value > 0:
        return it.value / sa.total_value * 100
    return None


class Runner:
    def __init__(self, rules, analyses, book, prev_analyses=None):
        self.rules = rules
        self.set = rules.get("cai_dat", {})
        self.an = analyses
        self.book = book
        self.prev = prev_analyses
        self.out = []
        summ = [a for a in analyses if a.role == "Tổng hợp" and a.items]
        self.summary = summ or [a for a in analyses if a.role != "Bỏ qua" and a.items]
        self.details = [a for a in analyses if a.role == "Chi tiết" and a.items]
        self.tol = float(self.set.get("nguong_lech_tuong_doi", 0.005))

    # -- tiện ích
    def add(self, r, sa, it=None, rid=None, group=True, **kw):
        vi_tri = f"{sa.name} / {it.label}" if (sa and it) else (sa.name if sa else "Toàn file")
        f = Finding(ma=r["ma"], muc=r["muc"], rule_id=rid or r.get("kiem_tra") or r.get("ten", r["ma"]),
                    ten_quy_tac=r.get("ten", ""), sheet=sa.name if sa else "", vi_tri=vi_tri,
                    hang_muc=kw.pop("hang_muc", it.ten if it else r.get("ten", "")),
                    khoi_luong=kw.pop("khoi_luong", _kl(it) if it else ""),
                    ty_trong=kw.pop("ty_trong", _tt(it, sa) if it and sa else None),
                    van_de=_fill(r.get("van_de"), **kw), cau_hoi=_fill(r.get("cau_hoi"), **kw),
                    kiem_chung=_fill(r.get("kiem_chung"), **kw),
                    phan=r.get("phan") or (it.phan if it else "Chung"),
                    trang_thai=r.get("trang_thai", ""), rows=[it.label] if it else [], group=group)
        self.out.append(f)
        return f

    def close(self, a, b):
        if a is None or b is None:
            return True
        return abs(a - b) <= max(0.01, self.tol * max(abs(a), abs(b)))

    # -- chạy
    def run(self):
        for r in self.rules.get("quy_tac", []):
            if not r.get("bat", True):
                continue
            loai = r.get("loai")
            try:
                if loai == "he_thong":
                    fn = getattr(self, "chk_" + r.get("kiem_tra", ""), None)
                    if fn:
                        fn(r)
                elif loai == "tu_khoa":
                    self.rule_keyword(r)
                elif loai == "don_vi":
                    self.rule_unit(r)
                elif loai == "chi_so":
                    self.rule_ratio(r)
            except Exception as e:     # một quy tắc lỗi không làm dừng cả bộ
                self.out.append(Finding(ma=r.get("ma", "?"), muc="Thấp", rule_id="loi", ten_quy_tac="Lỗi quy tắc",
                                        sheet="", vi_tri="", hang_muc=r.get("ten", ""), khoi_luong="",
                                        ty_trong=None, van_de=f"Quy tắc chạy lỗi: {e}", cau_hoi="",
                                        kiem_chung="", phan="Chung", nguon="Lỗi", group=False))
        return self.out

    # -- quy tắc hệ thống
    def chk_luy_ke_vuot_hd(self, r):
        ng = float((r.get("tham_so") or {}).get("nguong", 0.0001))
        for sa in self.summary:
            for it in sa.items:
                hd, lk = it.get("hd"), it.get("luyke")
                if lk is None and it.get("kynay") is not None and "truoc" in sa.tmap.cols:
                    lk = (it.get("truoc") or 0) + it.get("kynay")
                if hd and hd > 0 and lk and lk > hd * (1 + ng) and not self.close(lk, hd):
                    self.add(r, sa, it, khoi_luong=_kl(it, "luyke") or f"{fmt(lk)} {unit_label(it.dvt)}",
                             luyke=fmt(lk), hd=fmt(hd), vuot="+" + pct(lk / hd - 1))

    def chk_ky_truoc_cong_ky_nay(self, r):
        for sa in self.summary:
            c = sa.tmap.cols
            if not all(k in c for k in ("truoc", "kynay", "luyke")):
                continue
            zero = []
            for it in sa.items:
                tr, kn, lk = it.get("truoc") or 0, it.get("kynay") or 0, it.get("luyke")
                if lk is None or (tr == 0 and kn == 0):
                    continue
                if lk == 0 and kn == 0 and tr > 0:
                    zero.append(it)                     # đã nghiệm thu trước đây nhưng lũy kế về 0
                    continue
                if not self.close(tr + kn, lk):
                    self.add(r, sa, it, truoc=fmt(tr), kynay=fmt(kn), tong=fmt(tr + kn), luyke=fmt(lk))
            if zero:
                zero.sort(key=lambda x: -((x.get("truoc") or 0) * (x.get("dongia") or 0)))
                tts = [t for t in (_tt(x, sa) for x in zero) if t]
                rr = dict(r, muc="Trung bình",
                          van_de="{n} dòng đã có khối lượng kỳ trước nhưng lũy kế bằng 0 (giảm trừ toàn bộ hoặc "
                                 "chuyển sang dòng khác), các dòng: {cac_dong}.",
                          cau_hoi="Khối lượng các dòng này chuyển đi đâu? Có biên bản điều chỉnh, giảm trừ không? "
                                  "Dòng thay thế (nếu có) có cộng trùng không?",
                          kiem_chung="Đối chiếu biên bản nghiệm thu các đợt trước và biên bản điều chỉnh khối lượng.")
                f = self.add(rr, sa, None, rid="giam_tru_toan_bo", group=False,
                             hang_muc=f"{zero[0].ten[:60]} (và {len(zero) - 1} dòng khác)" if len(zero) > 1 else zero[0].ten,
                             khoi_luong="", ty_trong=None, n=len(zero), cac_dong=_rows(zero, 12))
                f.phan = zero[0].phan
                f.rows = [x.label for x in zero]

    def chk_dot_chiem_ty_le(self, r):
        tl = float((r.get("tham_so") or {}).get("ty_le", 0.6))
        for sa in self.summary:
            for it in sa.items:
                hd, kn, tr = it.get("hd"), it.get("kynay"), it.get("truoc")
                if not (hd and hd > 0 and kn and kn > 0) or tr:
                    continue
                t = _tt(it, sa)
                if kn / hd >= tl and (t is None or t >= 1.0):
                    self.add(r, sa, it, kynay=fmt(kn), ty_le=pct(kn / hd, 0))

    def chk_ngoai_hop_dong(self, r):
        for sa in self.summary:
            if "hd" not in sa.tmap.cols:
                continue
            for it in sa.items:
                if it.get("hd") in (None, 0) and ((it.get("kynay") or 0) > 0 or (it.get("luyke") or 0) > 0):
                    self.add(r, sa, it)

    def chk_trung_ten(self, r):
        for sa in self.summary:
            groups = defaultdict(list)
            for it in sa.items:
                groups[(norm(it.section), norm(it.ma), norm(it.ten), it.dvt)].append(it)
            for k, lst in groups.items():
                # chỉ coi là trùng khi các dòng nằm gần nhau (≤ 40 dòng); cùng tên ở phụ lục khác là bình thường
                lst = [x for x in lst if any(y is not x and abs(y.idx - x.idx) <= 40 for y in lst)]
                if len(lst) > 1 and any(x.q for x in lst):
                    top = max(lst, key=lambda x: x.value or 0)
                    f = self.add(r, sa, top, group=False, cac_dong=_rows(lst))
                    f.rows = [x.label for x in lst]
                    f.vi_tri = f"{sa.name} / " + ", ".join(x.label for x in lst[:6])

    def chk_so_tron(self, r):
        p = r.get("tham_so") or {}
        ng = float(p.get("nguong_gia_tri", 10))
        tl = float(p.get("ty_le_toi_thieu", 0.3))
        for sa in self.summary:
            pool = [it for it in sa.items if it.dvt in ("m3", "m2", "tan") and it.q]
            hits = [it for it in pool if abs(it.q) >= ng and float(it.q).is_integer()]
            if len(hits) >= 3 and (len(hits) / max(len(pool), 1) >= tl or len(hits) >= 10):
                hits.sort(key=lambda x: -(x.value or abs(x.q)))
                tt = sum(_tt(x, sa) or 0 for x in hits) or None
                rr = dict(r)
                if tt and tt >= 10:
                    rr["muc"] = "Trung bình"
                f = self.add(rr, sa, None, group=False, hang_muc=f"{len(hits)} dòng khối lượng số tròn",
                             khoi_luong="", ty_trong=None, so_dong=len(hits),
                             cac_dong=_rows(hits, 10))
                f.phan = hits[0].phan
                f.rows = [x.label for x in hits]

    def chk_dong_an(self, r):
        # Chỉ xét bảng tổng hợp; bảng diễn giải hay gom nhóm (outline) nên ẩn hàng loạt là bình thường
        for sa in self.summary:
            for it in sa.items:
                if it.hidden and (it.q or it.value):
                    self.add(r, sa, it)

    def chk_sheet_an_lien_ket(self, r):
        parts = []
        hid = [a.name for a in self.an if a.hidden]
        if hid:
            parts.append("Có sheet ẩn: " + ", ".join(hid) + ".")
        if self.book.external_links:
            parts.append(f"File có {self.book.external_links} liên kết sang file khác – số liệu có thể không cập nhật.")
        if parts:
            self.add(r, None, None, group=False, hang_muc="Cấu trúc file", noi_dung=" ".join(parts))

    def chk_o_loi(self, r):
        summ_names = {a.name for a in self.summary}
        others = []
        for sa in self.an:
            if not sa.errors:
                continue
            if sa.name in summ_names:
                self.add(r, sa, None, group=False, hang_muc="Ô lỗi công thức", so_o=len(sa.errors),
                         cac_o=", ".join(sa.errors[:10]) + ("…" if len(sa.errors) > 10 else ""))
            else:
                others.append(sa)
        if others:
            rr = dict(r, muc="Trung bình")
            n = sum(len(a.errors) for a in others)
            f = self.add(rr, None, None, group=False, hang_muc="Ô lỗi công thức ở bảng chi tiết", so_o=n,
                         cac_o="; ".join(f"{a.name} ({len(a.errors)} ô, vd {a.errors[0]})" for a in others[:8])
                         + ("…" if len(others) > 8 else ""))
            f.vi_tri = f"{len(others)} sheet chi tiết"

    def chk_nhap_tay(self, r):
        if not self.details:
            return
        for sa in self.summary:
            n = len(sa.items)
            if n and sa.hard_count / n >= 0.3:
                self.add(r, sa, None, group=False, hang_muc="Liên kết tổng hợp – chi tiết",
                         so_dong=sa.hard_count, tong_dong=n)

    def chk_dong_cong(self, r):
        base_tol = self.tol
        for sa in self.an:
            if not sa.totals:
                continue
            self.tol = 0.02 if sa.kind == "pdf" else base_tol      # số trên PDF có thể bị làm tròn khi in
            seq = sorted([("i", it) for it in sa.items] + [("t", t) for t in sa.totals], key=lambda x: x[1].idx)
            since, allv, subs = defaultdict(float), defaultdict(float), defaultdict(float)
            for kind, x in seq:
                vkeys = [k for k in ("tt_luyke", "tt_kynay", "thanhtien", "tt_hd") if x.get(k) not in (None,)]
                if kind == "i":
                    for k in vkeys:
                        v = x.get(k) or 0
                        since[k] += v
                        allv[k] += v
                    continue
                n = norm(x.ten)
                if "thue" in n or "vat" in n.split() or n.startswith(("gia tri sau", "gia tri truoc")):
                    continue                                  # dòng có thuế không phải phép cộng thuần
                bad = None
                for k in vkeys:
                    w = x.get(k)
                    if not w:
                        continue
                    if not any(self.close(w, c) for c in (since[k], allv[k], subs[k]) if c):
                        cand = since[k] or allv[k]
                        if cand:
                            bad = (k, w, cand)
                            break
                if bad:
                    k, w, cand = bad
                    self.add(r, sa, x, group=False, hang_muc=x.ten, khoi_luong="", ty_trong=None,
                             ten_dong=x.ten[:60], ghi=fmt(w, 0), tinh=fmt(cand, 0), lech=fmt(w - cand, 0))
                for k in vkeys:
                    subs[k] += x.get(k) or 0
                since = defaultdict(float)
        self.tol = base_tol

    def chk_tong_hop_chi_tiet(self, r):
        if not self.details:
            return
        idx = defaultdict(list)
        for da in self.details:
            for it in da.items:
                if it.ma and it.q:
                    idx[(norm(it.ma), it.dvt)].append((da, it))
        for sa in self.summary:
            if sa.role != "Tổng hợp":
                continue
            for it in sa.items:
                if not it.ma or not it.q:
                    continue
                lst = idx.get((norm(it.ma), it.dvt))
                if not lst or len({d.name for d, _ in lst}) > 1:
                    continue
                total = sum(x.q for _, x in lst)
                cands = [it.get(k) for k in ("kynay", "luyke", "kl", "hd") if it.get(k)]
                if cands and not any(self.close(total, c) for c in cands):
                    d0 = lst[0][0]
                    self.add(r, sa, it, tong_hop=" / ".join(fmt(c) for c in cands) + f" {unit_label(it.dvt)}",
                             chi_tiet=f"{fmt(total)} {unit_label(it.dvt)}", sheet_ct=d0.name,
                             dong_ct=", ".join(x.label.replace("dòng ", "") for _, x in lst[:5]))

    def chk_so_am(self, r):
        for sa in self.summary:
            for it in sa.items:
                for k in ("kynay", "luyke", "kl"):
                    v = it.get(k)
                    if v is not None and v < 0:
                        self.add(r, sa, it, gia_tri=fmt(v))
                        break

    _GRADE = [re.compile(r"(?:mac|\bm)\s*-?\s*(\d{3})\b"), re.compile(r"\bb\s*(\d{2}(?:[,.]5)?)\b"),
              re.compile(r"\bc\s*(\d{2})/\d{2}\b")]

    def chk_nhieu_mac(self, r):
        for sa in self.summary:
            gs, rows = set(), []
            for it in sa.items:
                n = norm(it.ten)
                if "be tong" not in n or "lot" in n:
                    continue
                for i, rx in enumerate(self._GRADE):
                    m = rx.search(n)
                    if m:
                        gs.add(("M" if i == 0 else "B" if i == 1 else "C") + m.group(1))
                        rows.append(it.label)
                        break
            if len(gs) > 1:
                f = self.add(r, sa, None, group=False, hang_muc="Bê tông nhiều mác/cấp", cac_mac=", ".join(sorted(gs)))
                f.rows = rows

    def _prev_index(self):
        if not self.prev:
            return None
        if hasattr(self, "_pidx"):
            return self._pidx
        idx, dup = {}, set()
        pa = [a for a in self.prev if a.role == "Tổng hợp" and a.items] or [a for a in self.prev if a.items]
        for a in pa:
            for it in a.items:
                for key in ((norm(it.ma), norm(it.ten)), ("", norm(it.ten), norm(it.section))):
                    if key in idx:
                        dup.add(key)
                    idx[key] = it
        for k in dup:
            idx.pop(k, None)              # trùng tên ở đợt trước → không ghép được chắc chắn
        self._pidx = idx
        return idx

    def _prev_of(self, idx, it):
        if self._dup_now(it):
            return None
        return idx.get((norm(it.ma), norm(it.ten))) or idx.get(("", norm(it.ten), norm(it.section)))

    def _dup_now(self, it):
        if not hasattr(self, "_cnt"):
            self._cnt = defaultdict(int)
            for sa in self.summary:
                for x in sa.items:
                    self._cnt[(norm(x.ma), norm(x.ten))] += 1
        return self._cnt[(norm(it.ma), norm(it.ten))] > 1

    def chk_so_voi_dot_truoc(self, r):
        idx = self._prev_index()
        if not idx:
            return
        for sa in self.summary:
            if "truoc" not in sa.tmap.cols:
                continue
            for it in sa.items:
                p = self._prev_of(idx, it)
                if not p:
                    continue
                plk = p.get("luyke")
                if plk is None and p.get("kynay") is not None:
                    plk = (p.get("truoc") or 0) + p.get("kynay")
                tr = it.get("truoc") or 0
                if plk is not None and not self.close(tr, plk):
                    self.add(r, sa, it, truoc=fmt(tr), luyke_truoc=fmt(plk))

    def chk_hd_thay_doi_so_dot_truoc(self, r):
        idx = self._prev_index()
        if not idx:
            return
        for sa in self.summary:
            for it in sa.items:
                p = self._prev_of(idx, it)
                if p and it.get("hd") is not None and p.get("hd") is not None and not self.close(it.get("hd"), p.get("hd")):
                    self.add(r, sa, it, hd=fmt(it.get("hd")), hd_truoc=fmt(p.get("hd")))

    def chk_don_vi_khac_dot_truoc(self, r):
        idx = self._prev_index()
        if not idx:
            return
        for sa in self.summary:
            for it in sa.items:
                p = self._prev_of(idx, it)
                if p and it.dvt and p.dvt and it.dvt != p.dvt:
                    self.add(r, sa, it, dvt=it.dvt_raw, dvt_truoc=p.dvt_raw)

    # -- quy tắc dữ liệu
    def rule_keyword(self, r):
        units = set(r.get("don_vi") or [])
        for sa in self.summary:
            for it in sa.items:
                if units and it.dvt not in units:
                    continue
                if r.get("chi_khi_co_kl") and not (it.q and it.q > 0):
                    continue
                hit = _match(it.ten, r)
                if hit:
                    self.add(r, sa, it, rid=r["ma"] + "|" + r.get("ten", ""), tu_khoa=hit)

    MEASURE = {"m3", "m2", "100m2", "100m3", "m", "100m", "tan", "kg", "10m3", "km"}

    def rule_unit(self, r):
        ok = set(r.get("don_vi_hop_le") or [])
        lim = int(r.get("vi_tri_toi_da", 20))      # từ khóa phải nằm ở đầu tên (công việc chính)
        for sa in self.summary:
            for it in sa.items:
                # chỉ báo khi đơn vị là đơn vị đo lường rõ ràng nhưng khác loại (m² thay m³…)
                if not it.dvt or it.dvt in ok or it.dvt not in self.MEASURE:
                    continue
                hit = _match(it.ten, r)
                if hit and 0 <= kw_pos(it.ten, hit) <= lim:
                    self.add(r, sa, it, rid=r["ma"] + "|" + r.get("ten", ""), group=False, dvt=it.dvt_raw)

    def rule_ratio(self, r):
        conv = {"kg": {"kg": 1, "tan": 1000}, "m2": {"m2": 1, "100m2": 100}, "m3": {"m3": 1, "100m3": 100}}
        ds = float(r.get("dung_sai", self.set.get("dung_sai_chi_so", 0.3)))
        lo = r.get("min")
        hi = r.get("max")
        for sa in self.summary:
            num_items = [it for it in sa.items if (not r["tu_so"].get("don_vi") or it.dvt in r["tu_so"]["don_vi"])
                         and _match(it.ten, r["tu_so"])]
            den_items = [it for it in sa.items if (not r["mau_so"].get("don_vi") or it.dvt in r["mau_so"]["don_vi"])
                         and _match(it.ten, r["mau_so"])]
            if not num_items or not den_items:
                continue
            du = set(r["mau_so"].get("don_vi") or [])
            cv_n = conv.get(r.get("quy_doi"), {})
            cv_d = conv.get(r.get("quy_doi_mau") or ("m3" if du and du <= {"m3", "100m3"} else r.get("quy_doi")), {})
            for col in ("luyke", "hd", "kl", "kynay"):
                nsum = sum((it.get(col) or 0) * cv_n.get(it.dvt, 1) for it in num_items)
                dsum = sum((it.get(col) or 0) * cv_d.get(it.dvt, 1) for it in den_items)
                if nsum > 0 and dsum > 0:
                    break
            else:
                continue
            if dsum < 1:
                continue
            ratio = nsum / dsum
            low = lo / (1 + ds) if lo is not None else None
            high = hi * (1 + ds) if hi is not None else None
            if (low is not None and ratio < low) or (high is not None and ratio > high):
                rr = dict(r)
                if col == "kynay":
                    rr["muc"] = "Thấp"
                lbl = {"luyke": "lũy kế", "hd": "hợp đồng", "kl": "khối lượng", "kynay": "kỳ này"}[col]
                tt = sum(_tt(x, sa) or 0 for x in num_items + den_items) or None
                f = self.add(rr, sa, None, rid=r["ma"] + "|" + r.get("ten", ""), group=False,
                             hang_muc=r.get("ten", ""), ty_trong=tt,
                             khoi_luong=f"{fmt(nsum)} / {fmt(dsum)} (cột {lbl})",
                             chi_so=fmt(ratio, 2), min=fmt(lo) if lo is not None else "–",
                             max=fmt(hi) if hi is not None else "–")
                f.rows = [x.label for x in num_items + den_items]
                f.vi_tri = f"{sa.name} / " + ", ".join(x.label for x in (num_items + den_items)[:6])


# --------------------------------------------------------------------------- gom và chọn
def _group(findings, threshold=3):
    """Gom phát hiện cùng quy tắc trong cùng sheet khi quá nhiều dòng."""
    buckets = defaultdict(list)
    keep = []
    for f in findings:
        if f.group:
            buckets[(f.rule_id, f.sheet)].append(f)
        else:
            keep.append(f)
    for (rid, sh), lst in buckets.items():
        if len(lst) <= threshold:
            keep.extend(lst)
            continue
        lst.sort(key=lambda x: -(x.ty_trong or 0))
        top = copy.copy(lst[0])
        labels = [x.rows[0] for x in lst if x.rows]
        top.rows = labels
        top.vi_tri = f"{sh} / {len(lst)} dòng: " + ", ".join(l.replace('dòng ', '') for l in labels[:12]) + \
                     ("…" if len(labels) > 12 else "")
        top.hang_muc = f"{lst[0].hang_muc} (và {len(lst) - 1} dòng khác)"
        tts = [x.ty_trong for x in lst if x.ty_trong]
        top.ty_trong = sum(tts) if tts else None
        top.van_de = f"{lst[0].ten_quy_tac}: {len(lst)} dòng. Dòng lớn nhất: {lst[0].van_de}"
        keep.append(top)
    return keep


def _merge_same_row(findings):
    """Một dòng công việc dính nhiều quy tắc → gộp thành một điểm, giữ mức cao nhất."""
    order = {k: i for i, k in enumerate(LEVELS)}
    buckets, keep = defaultdict(list), []
    for f in findings:
        if len(f.rows) == 1 and f.sheet and f.nguon == "Quy tắc":
            buckets[(f.sheet, f.rows[0])].append(f)
        else:
            keep.append(f)
    for lst in buckets.values():
        if len(lst) == 1:
            keep.append(lst[0])
            continue
        lst.sort(key=lambda x: order.get(x.muc, 9))
        m = copy.copy(lst[0])
        m.ma = ", ".join(dict.fromkeys(x.ma for x in lst))
        m.van_de = " ".join(dict.fromkeys(x.van_de for x in lst))
        m.cau_hoi = " ".join(dict.fromkeys(x.cau_hoi for x in lst if x.cau_hoi))
        m.kiem_chung = " ".join(dict.fromkeys(x.kiem_chung for x in lst if x.kiem_chung))
        m.ten_quy_tac = "; ".join(dict.fromkeys(x.ten_quy_tac for x in lst))
        keep.append(m)
    return keep


def _dedupe_sheets(findings):
    """Hồ sơ hay có nhiều bảng tổng hợp lặp lại cùng số liệu (THKL, giá trị, bảng vật tư…):
    cùng quy tắc + cùng hạng mục + cùng khối lượng thì chỉ giữ một điểm, ghi chú sheet trùng."""
    seen, out = {}, []
    for f in findings:
        if f.rows and len(f.rows) > 1:
            key = (f.ma, f.ten_quy_tac, len(f.rows))
        else:
            key = (f.ma, norm(f.hang_muc)[:80], f.van_de)
        if f.sheet and key in seen:
            g = seen[key]
            if f.sheet != g.sheet and f.sheet not in g.vi_tri:
                g.vi_tri += f" (trùng ở sheet {f.sheet})"
            continue
        seen[key] = f
        out.append(f)
    return out


def score(f, settings):
    dm = settings.get("diem_muc", {"Cao": 100, "Trung bình": 60, "Thấp": 20})
    hs = float(settings.get("he_so_ty_trong", 2.0))
    return dm.get(f.muc, 0) + hs * min(f.ty_trong or 0, 30)


def select(findings, settings, n=None):
    n = n or int(settings.get("so_diem_toi_da", 20))
    real = [f for f in findings if f.nguon != "Lỗi"]
    for f in real:
        f.score = score(f, settings)
    real.sort(key=lambda f: -f.score)
    return real[:n], real[n:]


# --------------------------------------------------------------------------- tổng hợp
def pareto(sa, cut=0.8):
    vals = sorted([it for it in sa.items if it.value and it.value > 0], key=lambda x: -x.value)
    if not vals or not sa.total_value:
        return [], 0, 0
    acc, out = 0, []
    for it in vals:
        acc += it.value
        out.append((it, it.value / sa.total_value, acc / sa.total_value))
        if acc / sa.total_value >= cut:
            break
    return out, len(out), acc / sa.total_value


def run_all(book, rules, overrides=None, prev_book=None, prev_overrides=None):
    analyses = analyze(book, overrides)
    prev = analyze(prev_book, prev_overrides) if prev_book else None
    runner = Runner(rules, analyses, book, prev)
    raw = runner.run()
    grouped = _dedupe_sheets(_merge_same_row(_group(raw)))
    return analyses, raw, grouped


def build_report(analyses, chosen, info, rules, book=None, rest=None):
    """Lập JSON theo cấu trúc qkl_report.py (SKILL.md mục 9)."""
    rest = rest or []
    cnt = {k: sum(1 for f in chosen if f.muc == k) for k in LEVELS}
    summ = [a for a in analyses if a.role == "Tổng hợp" and a.items] or [a for a in analyses if a.items]
    n_items = sum(len(a.items) for a in summ)
    p_txt, tap = "", []
    if summ:
        main = max(summ, key=lambda a: a.total_value)
        pl, k, share = pareto(main)
        if k:
            p_txt = f"{k} dòng = {pct(share, 0)}"
            tap = [f"{it.ten[:70]} ({pct(s, 1)})" for it, s, _ in pl[:5]]
    danh_gia = "Cao" if cnt["Cao"] else ("Trung bình" if cnt["Trung bình"] else "Thấp")
    nx = (f"Quét tự động {len(summ)} bảng tổng hợp, {n_items} dòng công việc theo bộ quy tắc "
          f"phiên bản {rules.get('phien_ban', '?')}. Chọn {len(chosen)} điểm cần rà soát "
          f"({cnt['Cao']} Cao, {cnt['Trung bình']} Trung bình, {cnt['Thấp']} Thấp)"
          + (f"; còn {len(rest)} phát hiện mức thấp hơn ở sheet Kết quả quét" if rest else "") + ". ")
    if p_txt:
        nx += f"Nhóm hạng mục chủ yếu: {p_txt} giá trị. "
    nx += "Kết quả là gợi ý của bộ lọc, người rà soát cần đối chiếu hồ sơ trước khi kết luận."
    if info.get("nhan_xet_them"):
        nx = info["nhan_xet_them"].strip() + " " + nx
    data = {
        "du_an": info.get("du_an", ""), "goi_thau": info.get("goi_thau", ""),
        "ho_so": info.get("ho_so", book.source if book else ""),
        "loai_ho_so": info.get("loai_ho_so", ""), "hop_dong": info.get("hop_dong", ""),
        "nguon_von": info.get("nguon_von", ""), "tinh_chat": info.get("tinh_chat", ""),
        "nguoi_lap": info.get("nguoi_lap", ""),
        "ngay": info.get("ngay") or datetime.date.today().isoformat(),
        "kho_giay": info.get("kho_giay", "A4"),
        "danh_gia_chung": danh_gia, "nhan_xet_chung": nx, "tap_trung": tap,
        "can_bo_sung": info.get("can_bo_sung", []),
        "nhom_chu_yeu": p_txt or "–",
        "diem": [f.to_point() for f in chosen],
        "thieu_ho_so": info.get("thieu_ho_so", []),
        "ket_qua_quet": scan_text(analyses, rest, book),
        "chu_ky": bool(info.get("chu_ky")),
        "chu_ky_ben": info.get("chu_ky_ben") or [["NGƯỜI RÀ SOÁT", info.get("nguoi_lap", "")],
                                                 ["TRƯỞNG PHÒNG KỸ THUẬT", "Hoàng Đức Vũ"]],
    }
    return data


def scan_text(analyses, rest, book=None):
    from .table import COL_KEYS
    L = []
    if book is not None:
        L.append(f"Hồ sơ: {book.source}")
        for w in book.warnings:
            L.append(f"Lưu ý: {w}")
    for a in analyses:
        cols = ", ".join(f"{COL_KEYS.get(k, k)}=cột {c + 1}" for k, c in a.tmap.cols.items() if k in COL_KEYS)
        L.append(f"Sheet {a.name}: {a.role}; {len(a.items)} dòng công việc; "
                 f"tiêu đề dòng {(a.tmap.header_row or 0) + 1 if a.tmap.header_row is not None else '–'}; {cols or 'không nhận diện được bảng'}")
    if rest:
        L.append("")
        L.append(f"Các phát hiện khác ({len(rest)}), xếp theo mức và tỷ trọng:")
        for f in rest:
            L.append(f"- [{f.muc}] {f.ma} | {f.vi_tri} | {f.hang_muc[:60]} | {f.van_de}")
    return "\n".join(L)
