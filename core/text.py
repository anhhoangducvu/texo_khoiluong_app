# -*- coding: utf-8 -*-
"""Chuẩn hóa chữ, số, đơn vị cho hồ sơ khối lượng tiếng Việt."""
import functools
import re
import unicodedata


def norm(s):
    """Chữ thường, bỏ dấu, đ→d, gộp khoảng trắng."""
    if s is None:
        return ""
    s = str(s).strip().lower().replace("đ", "d").replace("Đ", "d")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", s).strip()


@functools.lru_cache(maxsize=4096)
def _wre(kw_n):
    return re.compile(r"(?<![a-z0-9])" + re.escape(kw_n) + r"(?![a-z0-9])")


def contains_word(text_n, kw_n):
    """Từ khóa (đã chuẩn hóa) có trong chuỗi, khớp theo ranh giới từ."""
    if not kw_n:
        return False
    return _wre(kw_n).search(text_n) is not None


def has_marks(s):
    """Chuỗi có dấu tiếng Việt (ký tự ngoài bảng chữ ASCII)."""
    return any(ord(c) > 127 and c.isalpha() for c in str(s or ""))


def low(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s or "")).lower()).strip()


@functools.lru_cache(maxsize=8192)
def _kre(kw_low):
    return re.compile(r"(?<!\w)" + re.escape(kw_low) + r"(?!\w)")


@functools.lru_cache(maxsize=200000)
def _forms(name):
    """(chữ thường có dấu, chữ bỏ dấu, có dấu?) – tính một lần cho mỗi tên công việc."""
    return low(name), norm(name), has_marks(name)


def kw_pos(name, kw):
    """Vị trí từ khóa trong tên công việc (−1 nếu không có).
    Tên và từ khóa đều có dấu → so có dấu (phân biệt cốt/cột, đài/dài, đắp/đáp).
    Tên không dấu (gõ không dấu, font cũ) → bỏ dấu cả hai rồi so."""
    if not kw or not name:
        return -1
    lname, nname, marks = _forms(str(name))
    if marks and has_marks(kw):
        m = _kre(low(kw)).search(lname)
    else:
        m = _wre(norm(kw)).search(nname)
    return m.start() if m else -1


def kw_in(name, kw):
    return kw_pos(name, kw) >= 0


_THOUSANDS_DOT = re.compile(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$")
_THOUSANDS_COMMA = re.compile(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$")


def num(v):
    """Đọc số từ ô Excel hoặc chuỗi PDF (ưu tiên quy ước Việt Nam: 1.234,56)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(" ", "").replace(" ", "")
    if not t or t in ("-", "–"):
        return None
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    if t.endswith("%"):
        t = t[:-1]
    try:
        if _THOUSANDS_DOT.match(t):
            t = t.replace(".", "").replace(",", ".")
        elif _THOUSANDS_COMMA.match(t) and "." in t:
            t = t.replace(",", "")
        elif "," in t and "." in t:
            if t.rfind(",") > t.rfind("."):
                t = t.replace(".", "").replace(",", ".")
            else:
                t = t.replace(",", "")
        elif "," in t:
            t = t.replace(",", ".") if t.count(",") == 1 else t.replace(",", "")
        val = float(t)
        return -val if neg else val
    except ValueError:
        return None


UNIT_MAP = [
    (r"^100\s*m\s*(2|²|vuong)$", "100m2"),
    (r"^100\s*m\s*(3|³|khoi)$", "100m3"),
    (r"^100\s*m(et)?\s*(dai)?$", "100m"),
    (r"^m\s*(2|²|vuong)$", "m2"),
    (r"^m\s*(3|³|khoi)$", "m3"),
    (r"^(tan|t|ton|tonne)$", "tan"),
    (r"^kg$", "kg"),
    (r"^(md|m dai|m|met|mdai)$", "m"),
    (r"^(bo|bo\.)$", "bo"),
    (r"^(cai|chiec)$", "cai"),
    (r"^(ca|ca may)$", "ca"),
    (r"^(cong|ngay cong)$", "cong"),
    (r"^(coc|cay)$", "coc"),
    (r"^(km)$", "km"),
    (r"^(m3\s*\.?\s*km|m3km)$", "m3km"),
    (r"^(10\s*m\s*3|10m3)$", "10m3"),
    (r"^(tron goi|tg|ls|lumpsum|lump sum|goi)$", "goi"),
]


def unit(v):
    """Chuẩn hóa đơn vị: m3, m2, 100m2, tan, kg, m, bo, cai…; trả '' nếu trống."""
    u = norm(v).replace("³", "3").replace("²", "2").strip(" .")
    if not u:
        return ""
    for pat, val in UNIT_MAP:
        if re.match(pat, u):
            return val
    return u


def unit_label(u):
    return {"m3": "m³", "m2": "m²", "100m2": "100 m²", "100m3": "100 m³",
            "tan": "tấn", "bo": "bộ", "cai": "cái", "cong": "công"}.get(u, u)


def fmt(x, nd=2):
    """Định dạng số kiểu Việt Nam: 1.234,56."""
    if x is None:
        return "–"
    try:
        x = float(x)
    except (TypeError, ValueError):
        return str(x)
    if abs(x - round(x)) < 1e-9 and abs(x) >= 1:
        s = f"{x:,.0f}"
    else:
        s = f"{x:,.{nd}f}"
    return s.replace(",", "#").replace(".", ",").replace("#", ".")


def pct(x, nd=1):
    if x is None:
        return "–"
    return fmt(x * 100, nd) + "%"


ROMAN = re.compile(r"^(i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii|xiii|xiv|xv)[\.\)]?$")
LETTER = re.compile(r"^[a-h][\.\)]?$")


def is_section_code(stt):
    s = norm(stt)
    return bool(ROMAN.match(s) or LETTER.match(s) or re.match(r"^(phan|hang muc|muc)\b", s))


_TOTAL_RE = re.compile(r"^(tong cong|cong|tong)\s*(:|\(|-|$|hang muc|phan|muc|gia tri|so\b|cac\b|truoc thue|sau thue)")
_TAX_RE = re.compile(r"^(thue (vat|gtgt)|gia tri (truoc|sau) thue|gia tri de nghi thanh toan)")


def is_total_name(name):
    n = norm(name)
    return bool(_TOTAL_RE.match(n) or _TAX_RE.match(n))
