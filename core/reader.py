# -*- coding: utf-8 -*-
"""Đọc hồ sơ khối lượng: Excel (.xlsx/.xlsm/.xls) và PDF có lớp chữ.

Mọi nguồn được đưa về cùng một dạng `Sheet`:
  grid        : list các dòng, mỗi dòng là list giá trị ô (dùng để đọc số liệu)
  head_grid   : như grid nhưng ô gộp (merge) vùng tiêu đề được điền lại (dùng nhận diện tiêu đề)
  formula     : list các dòng, True nếu ô là công thức (chỉ Excel)
  hidden_rows : tập chỉ số dòng (0-based) đang bị ẩn
  labels      : nhãn vị trí từng dòng để ghi vào kết quả ("dòng 15", "tr.3")
"""
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

ERR_PREFIX = ("#REF", "#VAL", "#DIV", "#N/A", "#NAM", "#NUM", "#NUL")


@dataclass
class Sheet:
    name: str
    kind: str                      # "excel" | "pdf"
    grid: list
    head_grid: list
    formula: list = None
    hidden_rows: set = field(default_factory=set)
    hidden: bool = False
    labels: list = None
    errors: list = field(default_factory=list)     # ["F12", ...]

    def label(self, i):
        if self.labels and i < len(self.labels):
            return self.labels[i]
        return f"dòng {i + 1}"


@dataclass
class Book:
    source: str
    kind: str
    sheets: list
    external_links: int = 0
    warnings: list = field(default_factory=list)
    scanned_pdf: bool = False
    path: str = None


# --------------------------------------------------------------------------- Excel
MAX_ROWS = 8000          # số dòng đọc tối đa mỗi sheet (bảng tổng hợp hiếm khi vượt)
MAX_COLS = 60            # số cột đọc tối đa
EMPTY_STOP = 300         # gặp liên tiếp ngần này dòng trống thì coi như hết bảng
HEAD_ROWS = 60           # vùng tìm dòng tiêu đề


def _recalc_with_soffice(path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    out = tempfile.mkdtemp(prefix="qkl_recalc_")
    try:
        subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", out, path],
                       capture_output=True, timeout=180)
        cand = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
        return cand if os.path.exists(cand) else None
    except Exception:
        return None


def _clean_copy(path):
    """Bản sao gọn: bỏ khối <definedNames> (có file chứa hàng trăm nghìn tên rác, nặng hàng chục MB
    làm openpyxl treo) và ảnh/đối tượng vẽ. Kết quả công thức đã lưu sẵn trong ô nên không ảnh hưởng số liệu.
    Trả (đường dẫn, số KB tên rác đã bỏ)."""
    out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    removed = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            name = info.filename
            if name.startswith(("xl/media/", "xl/drawings/", "xl/printerSettings/", "xl/embeddings/")):
                continue
            data = zin.read(name)
            if name == "xl/workbook.xml":
                i, j = data.find(b"<definedNames"), data.find(b"</definedNames>")
                if i >= 0 and j > i:
                    removed = (j - i) // 1024
                    data = data[:i] + data[j + len(b"</definedNames>"):]
            elif name.startswith("xl/worksheets/_rels/") or name.startswith("xl/worksheets/sheet"):
                # bỏ liên kết tới bản vẽ đã loại
                data = re.sub(rb"<drawing [^>]*/>|<legacyDrawing [^>]*/>|<picture [^>]*/>", b"", data) \
                    if name.startswith("xl/worksheets/sheet") else \
                    re.sub(rb"<Relationship [^>]*(drawing|image|printerSettings|oleObject|vmlDrawing)[^>]*/>", b"", data)
            zout.writestr(info, data)
    return out, removed


def _sheet_parts(zf):
    """Tên sheet → đường dẫn XML trong gói."""
    wb = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    rmap = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        tgt = attrs.get("Target", "")
        tgt = tgt.lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        rmap[attrs.get("Id")] = tgt
    out = {}
    for m in re.finditer(r"<sheet\b([^>]*)/>", wb):
        name = re.search(r'name="([^"]*)"', m.group(1))
        rid = re.search(r'r:id="([^"]*)"', m.group(1))
        if name and rid:
            nm = (name.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                  .replace("&quot;", '"').replace("&apos;", "'"))
            out[nm] = rmap.get(rid.group(1))
    return out


_ROW_HID = re.compile(rb'<row\b[^>]*?\br="(\d+)"[^>]*?\bhidden="(?:1|true)"')
_ROW_HID2 = re.compile(rb'<row\b[^>]*?\bhidden="(?:1|true)"[^>]*?\br="(\d+)"')
_NOVAL_RE = re.compile(rb"</f>(?:<v></v>|<v/>)?</c>|<f [^>]*/>(?:<v></v>|<v/>)?</c>")
_NOVAL = [0]                      # đếm ô công thức không có giá trị lưu sẵn
_MERGE = re.compile(rb'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"')


def _col_num(letters):
    n = 0
    for ch in letters.decode() if isinstance(letters, bytes) else letters:
        n = n * 26 + ord(ch) - 64
    return n


def _scan_meta(zf, member, max_rows):
    """Đọc lướt XML sheet theo từng khúc: dòng ẩn và ô gộp vùng tiêu đề (không dựng cây XML)."""
    hidden, merges = set(), []
    if not member:
        return hidden, merges
    tail = b""
    try:
        with zf.open(member) as fh:
            while True:
                chunk = fh.read(1 << 20)
                if not chunk:
                    break
                buf = tail + chunk
                for rx in (_ROW_HID, _ROW_HID2):
                    for m in rx.finditer(buf):
                        r = int(m.group(1))
                        if r <= max_rows:
                            hidden.add(r - 1)
                if _NOVAL[0] <= 50:
                    _NOVAL[0] += len(_NOVAL_RE.findall(chunk))
                for m in _MERGE.finditer(buf):
                    r1, r2 = int(m.group(2)), int(m.group(4))
                    if r1 <= HEAD_ROWS and r2 - r1 <= 3:
                        merges.append((r1, _col_num(m.group(1)), r2, _col_num(m.group(3))))
                tail = buf[-400:]
    except KeyError:
        pass
    return hidden, list(dict.fromkeys(merges))


def _stream_values(ws):
    """Đọc giá trị theo luồng, dừng khi quá MAX_ROWS hoặc gặp EMPTY_STOP dòng trống liên tiếp."""
    grid, empty, truncated = [], 0, False
    for k, row in enumerate(ws.iter_rows(max_col=MAX_COLS, values_only=True)):
        if k >= MAX_ROWS:
            truncated = True
            break
        row = list(row)
        if any(v not in (None, "") for v in row):
            empty = 0
        else:
            empty += 1
            if empty > EMPTY_STOP:
                break
        grid.append(row)
    while grid and not any(v not in (None, "") for v in grid[-1]):
        grid.pop()
    nc = 0
    for row in grid:
        for c in range(len(row) - 1, -1, -1):
            if row[c] not in (None, ""):
                nc = max(nc, c + 1)
                break
    grid = [(r[:nc] + [None] * (nc - len(r[:nc]))) for r in grid]
    return grid, truncated


def _read_xlsx(path, display_name, _recalc=False):
    import warnings as _w
    import openpyxl
    _w.filterwarnings("ignore", module="openpyxl")
    warnings = []
    _NOVAL[0] = 0
    clean, removed = _clean_copy(path)
    if removed > 1024:
        warnings.append(f"File chứa khoảng {removed // 1024} MB tên vùng (defined names) thừa – app đã bỏ qua "
                        "để đọc nhanh. Nên dọn Name Manager trong file gốc.")
    wb_v = openpyxl.load_workbook(clean, read_only=True, data_only=True)
    ext = 0
    try:
        ext = len(wb_v._external_links or [])
    except Exception:
        pass
    zf = zipfile.ZipFile(clean)
    parts = _sheet_parts(zf)
    sheets, truncated = [], []
    for ws in wb_v.worksheets:
        grid, trunc = _stream_values(ws)
        if trunc:
            truncated.append(ws.title)
        nr, nc = len(grid), (len(grid[0]) if grid else 0)
        errors = []
        for i, row in enumerate(grid):
            for j, v in enumerate(row):
                if isinstance(v, str) and v.strip().upper().startswith(ERR_PREFIX):
                    errors.append(f"{get_column_letter(j + 1)}{i + 1}")
        hidden_rows, merges = _scan_meta(zf, parts.get(ws.title), max(nr, 1))
        head = [list(row) for row in grid[:HEAD_ROWS]]
        for r1, c1, r2, c2 in merges:
            if r1 - 1 >= len(head) or c1 - 1 >= nc:
                continue
            v = grid[r1 - 1][c1 - 1]
            for rr in range(r1, min(r2, len(head)) + 1):
                for cc in range(c1, min(c2, nc) + 1):
                    if head[rr - 1][cc - 1] is None:
                        head[rr - 1][cc - 1] = v
        head += grid[HEAD_ROWS:]
        sheets.append(Sheet(name=ws.title, kind="excel", grid=grid, head_grid=head, formula=None,
                            hidden_rows={r for r in hidden_rows if r < nr}, hidden=(ws.sheet_state != "visible"),
                            labels=None, errors=errors))
    zf.close()
    wb_v.close()
    if truncated:
        warnings.append(f"Sheet dài hơn {MAX_ROWS} dòng chỉ đọc {MAX_ROWS} dòng đầu: " + ", ".join(truncated[:6])
                        + ("…" if len(truncated) > 6 else "") + ".")
    book = Book(source=display_name, kind="excel", sheets=sheets, external_links=ext, warnings=warnings)
    book.path = clean
    # Ô công thức chưa có kết quả lưu sẵn (file do phần mềm khác xuất): tính lại nếu file không quá lớn
    if _NOVAL[0] > 20 and not _recalc:
        if os.path.getsize(clean) < 8 * 1024 * 1024:
            rc = _recalc_with_soffice(clean)
            if rc:
                b2 = _read_xlsx(rc, display_name, _recalc=True)
                b2.warnings = warnings + ["File chưa lưu sẵn kết quả công thức; app đã tính lại bằng LibreOffice."]
                return b2
        book.warnings.append("File có nhiều công thức chưa lưu sẵn kết quả (thường do phần mềm khác xuất). "
                             "Mở file bằng Excel, bấm Lưu rồi tải lại để kết quả quét đầy đủ.")
    return book


def load_formulas(book, names):
    """Nạp cờ 'ô là công thức' cho các sheet cần (lazy, chỉ bảng tổng hợp), đọc theo luồng."""
    path = getattr(book, "path", None)
    if not path or book.kind != "excel":
        return
    need = [sh for sh in book.sheets if sh.name in names and sh.formula is None]
    if not need:
        return
    import openpyxl
    wb_f = openpyxl.load_workbook(path, read_only=True, data_only=False)
    missing = 0
    for sh in need:
        nr, nc = len(sh.grid), (len(sh.grid[0]) if sh.grid else 0)
        fl = []
        if nc:
            for k, row in enumerate(wb_f[sh.name].iter_rows(max_row=nr, max_col=nc, values_only=True)):
                flags = [isinstance(v, str) and v.startswith("=") for v in row] + [False] * (nc - len(row))
                for j, f in enumerate(flags):
                    if f and sh.grid[k][j] is None:
                        missing += 1
                fl.append(flags)
        fl += [[False] * nc for _ in range(nr - len(fl))]
        sh.formula = fl
    wb_f.close()
    if missing > 20 and not getattr(book, "recalc_tried", False):
        book.recalc_tried = True
        book.warnings.append("Bảng tổng hợp có công thức chưa lưu sẵn kết quả (file do phần mềm khác xuất). "
                             "Mở file bằng Excel, bấm Lưu rồi tải lại để kết quả quét đầy đủ.")


def _read_xls(path, display_name):
    import xlrd
    try:
        wb = xlrd.open_workbook(path, formatting_info=True)
        fmt_ok = True
    except Exception:
        wb = xlrd.open_workbook(path)
        fmt_ok = False
    sheets = []
    for sh in wb.sheets():
        grid = []
        for r in range(sh.nrows):
            row = []
            for c in range(min(sh.ncols, 60)):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_ERROR:
                    row.append("#ERR")
                elif cell.ctype == xlrd.XL_CELL_EMPTY:
                    row.append(None)
                else:
                    row.append(cell.value)
            grid.append(row)
        head = [list(r) for r in grid]
        hidden = set()
        if fmt_ok:
            for (r1, r2, c1, c2) in sh.merged_cells:
                if r1 > 60 or r2 - r1 > 4:
                    continue
                v = grid[r1][c1] if r1 < len(grid) and c1 < len(grid[r1]) else None
                for rr in range(r1, min(r2, len(head))):
                    for cc in range(c1, min(c2, len(head[rr]))):
                        if head[rr][cc] in (None, ""):
                            head[rr][cc] = v
            for r, info in sh.rowinfo_map.items():
                if getattr(info, "hidden", 0):
                    hidden.add(r)
        sheets.append(Sheet(name=sh.name, kind="excel", grid=grid, head_grid=head, formula=None,
                            hidden_rows=hidden, hidden=bool(getattr(sh, "visibility", 0)),
                            labels=[f"dòng {i + 1}" for i in range(len(grid))]))
    return Book(source=display_name, kind="excel", sheets=sheets,
                warnings=["File .xls (định dạng cũ): không kiểm tra được công thức nhập tay. "
                          "Nên lưu lại thành .xlsx để quét đầy đủ."])


# --------------------------------------------------------------------------- PDF
def _read_pdf(path, display_name):
    import pdfplumber
    blocks = []   # [ncols, rows, labels, pages]
    total_chars = 0
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            total_chars += len(page.chars)
            tables = page.extract_tables()
            if not tables:
                tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text"})
            for t in tables:
                rows = [[(re.sub(r"\(cid:\d+\)", "", c).strip() if isinstance(c, str) else c) for c in row]
                        for row in t if row]
                rows = [r for r in rows if any(x not in (None, "") for x in r)]
                if not rows:
                    continue
                nc = max(len(r) for r in rows)
                rows = [r + [None] * (nc - len(r)) for r in rows]
                labels = [f"tr.{pno} d.{k}" for k in range(1, len(rows) + 1)]
                if blocks and blocks[-1][0] == nc and pno - blocks[-1][3][-1] <= 1:
                    blocks[-1][1].extend(rows)
                    blocks[-1][2].extend(labels)
                    if blocks[-1][3][-1] != pno:
                        blocks[-1][3].append(pno)
                else:
                    blocks.append([nc, rows, labels, [pno]])
    book = Book(source=display_name, kind="pdf", sheets=[])
    if total_chars < 50:
        book.scanned_pdf = True
        book.warnings.append("PDF không có lớp chữ (bản scan). Bản này chưa đọc được ảnh – "
                             "đề nghị xin file Excel gốc hoặc PDF xuất trực tiếp từ Excel.")
        return book
    for i, (nc, rows, labels, pages) in enumerate(blocks, 1):
        pg = f"tr.{pages[0]}" + (f"–{pages[-1]}" if len(pages) > 1 else "")
        head = [list(r) for r in rows]
        # Ô gộp ngang ở vùng tiêu đề PDF chỉ có chữ ở ô đầu: điền sang phải nếu dòng dưới có chữ
        for r in range(min(len(head) - 1, 10)):
            for c in range(1, nc):
                if head[r][c] in (None, "") and head[r][c - 1] not in (None, "") and head[r + 1][c] not in (None, ""):
                    head[r][c] = head[r][c - 1]
        book.sheets.append(Sheet(name=f"PDF bảng {i} ({pg})", kind="pdf", grid=rows, head_grid=head, labels=labels))
    if not book.sheets:
        book.warnings.append("Không tách được bảng nào trong PDF.")
    return book


# --------------------------------------------------------------------------- API
def load(file_or_path, name=None):
    """Đọc file từ đường dẫn hoặc đối tượng upload (có .read()/.name). Trả Book."""
    if hasattr(file_or_path, "read"):
        name = name or getattr(file_or_path, "name", "ho_so")
        data = file_or_path.read()
        suffix = os.path.splitext(name)[1].lower()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(data)
        tmp.close()
        path = tmp.name
    else:
        path = str(file_or_path)
        name = name or os.path.basename(path)
    ext = os.path.splitext(name)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(path, name)
    if ext == ".xls":
        return _read_xls(path, name)
    if ext == ".pdf":
        return _read_pdf(path, name)
    raise ValueError(f"Chưa hỗ trợ định dạng {ext}. Dùng .xlsx, .xlsm, .xls hoặc .pdf.")
