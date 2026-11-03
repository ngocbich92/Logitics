"""Biểu mẫu outbound (SI Form, DGD, VGM, CMA, ZIM, MSC, ONE, OOCL, Hapag-L, CVC, PK, INV, ...).

Mẫu = các sheet biểu mẫu gốc của file 'Database outbound shipping plan.xlsx' (templates/outbound_forms.xlsx), giữ nguyên định dạng và
công thức VLOOKUP gốc. Công thức được tính bằng xl_eval trên dữ liệu trong hệ thống, rồi ghi giá trị vào bản sao của mẫu —
nên biểu mẫu xuất ra giống hệt biểu mẫu trong Excel.
"""
import datetime as dt
import io
import re
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

import outbound
import rvc_form
import xl_eval
from xl_eval import XlError

TEMPLATE = Path(__file__).parent / "templates" / "outbound_forms.xlsx"

FORMS = [  # (sheet, tên hiển thị, loại khoá: ship = Shipment ID của Source_Data; inv = mã trong List shipment)
    ("SI Form", "SI Form — Shipping Instruction / BL draft (hàng DG)", "ship"),
    ("SI Non-DG", "SI Non-DG — Shipping Instruction (hàng thường)", "ship"),
    ("DGD WH", "DGD WH — Dangerous Goods Declaration", "ship"),
    ("DGD air", "DGD air — Khai báo hàng nguy hiểm đường hàng không", "ship"),
    ("VGM", "VGM — Verified Gross Mass", "ship"),
    ("Manifest", "Manifest", "ship"),
    ("CMA", "CMA CGM", "ship"),
    ("ZIM", "ZIM", "ship"),
    ("IMDG ZIM 2cont", "IMDG ZIM (2 container)", "ship"),
    ("MSC", "MSC", "ship"),
    ("ONE", "ONE", "ship"),
    ("OOCL", "OOCL", "ship"),
    ("Hapag-L", "Hapag-Lloyd", "ship"),
    ("CVC", "CVC", "ship"),
    ("PK", "Packing list", "inv"),
    ("INV", "Commercial invoice", "inv"),
]


def _available():
    """Chỉ giữ các biểu mẫu thực sự có trong mẫu (bạn có thể xoá bớt sheet trong file Excel gốc rồi tạo lại mẫu)."""
    try:
        wb = openpyxl.load_workbook(TEMPLATE, read_only=True)
        names = set(wb.sheetnames)
        wb.close()
    except OSError:
        return FORMS
    return [f for f in FORMS if f[0] in names]


FORMS = _available()
KIND = {s: k for s, _, k in FORMS}


def _walk(node, fn):
    fn(node)
    if node[0] == "func":
        for a in node[2]:
            _walk(a, fn)
    elif node[0] == "bin":
        _walk(node[2], fn)
        _walk(node[3], fn)
    elif node[0] == "neg":
        _walk(node[1], fn)


def _key_cells(ws, sheet):
    """Các ô tĩnh được dùng làm khoá tra cứu (đối số đầu của VLOOKUP / MATCH) — là ô người dùng nhập mã shipment."""
    found = set()

    def visit(n):
        if n[0] == "func" and n[1] in ("VLOOKUP", "MATCH") and n[2] and n[2][0][0] == "cell":
            _, sh, c, r = n[2][0]
            if (sh or sheet) == sheet and r:
                found.add((c, r))

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                try:
                    _walk(xl_eval.parse(cell.value), visit)
                except XlError:
                    pass
    return {(c, r) for c, r in found if not (isinstance(ws.cell(r, c).value, str) and str(ws.cell(r, c).value).startswith("="))}


class _Ctx:
    def __init__(self, conn, wb, overrides):
        self.conn, self.wb, self.overrides = conn, wb, overrides
        self.memo, self.stack, self.errors = {}, set(), []
        self.ev = xl_eval.Evaluator(self)
        self.row_cache = {}

    def today(self):
        return dt.date.today()

    def lookup_row(self, sheet, key):
        if sheet == "InvoiceData":
            return invoice_row(self.conn, key)
        tbl = outbound.SHEET_TO_TABLE.get(sheet)
        if tbl is None:
            return None  # bảng không tồn tại (vd tham chiếu #REF! trong file gốc) -> #N/A
        k = (tbl, str(key).lower())
        if k not in self.row_cache:
            self.row_cache[k] = outbound.find_row(self.conn, tbl, key)
        return self.row_cache[k]

    def cell_value(self, sheet, c, r):
        key = (sheet, c, r)
        if key in self.overrides:
            return self.overrides[key]
        if sheet not in self.wb.sheetnames:
            return None
        if key in self.memo:
            res = self.memo[key]
            if isinstance(res, XlError):
                raise res
            return res
        v = self.wb[sheet].cell(r, c).value
        if isinstance(v, str) and v.startswith("="):
            if sheet in ("INV", "PK") and "#REF!" in v:
                # bảng nguồn của INV / PK đã bị xoá trong file gốc (#REF!): nay lấy từ Invoice đã quét (cùng chỉ số cột như cũ)
                v = v.replace("#REF!", "InvoiceData!B:BH")
            if key in self.stack:
                raise XlError("#REF!")
            self.stack.add(key)
            try:
                res = self.ev.ev(xl_eval.parse(v), sheet)
            except XlError as e:
                res = e
            except (ValueError, TypeError, ZeroDivisionError):
                res = XlError("#VALUE!")
            finally:
                self.stack.discard(key)
            self.memo[key] = res
            if isinstance(res, XlError):
                raise res
            return res
        return outbound.dec(v) if isinstance(v, str) else v


INVOICE_MAP = {  # cột của "bảng nguồn" cũ (VLOOKUP(A2,#REF!,n,0): n=2 -> C, 6 -> G, 7 -> H, 36 -> AK, 53 -> BB, 54 -> BC)
    "C": "polarium_ref", "G": "deliver_to", "H": "deliver_to", "AK": "incoterms", "BB": "invoice_no", "BC": "invoice_date"}


def invoice_row(conn, key):
    inv = outbound.find_invoice(conn, key)
    if not inv:
        return None
    row = {}
    for letter, field in INVOICE_MAP.items():
        v = inv.get(field)
        row[letter] = outbound.dec(v) if isinstance(v, str) else v
    return row


def _fill_invoice_lines(conn, wb, inv_key):
    """Biểu mẫu INV: điền HS code / CoO / Unit price từ dòng hàng của Invoice đã quét (khớp theo mã hàng)."""
    inv = outbound.find_invoice(conn, inv_key)
    if not inv or "INV" not in wb.sheetnames:
        return
    import db
    lines = {}
    for g in outbound.goods_get(conn, inv["ship_key"]):
        if (not inv["invoice_no"] or g.get("inv_no") == inv["invoice_no"]) and g.get("ma_hang"):
            lines.setdefault(db.ma_key(g["ma_hang"]), g)
    ws = wb["INV"]
    for r in range(22, ws.max_row + 1):
        item = ws.cell(r, 3).value
        g = lines.get(db.ma_key(item)) if item else None
        if g:
            ws.cell(r, 7).value = inv.get("hs_code") or ws.cell(r, 7).value
            ws.cell(r, 8).value = g.get("xuat_xu") or ws.cell(r, 8).value
            ws.cell(r, 9).value = g.get("don_gia") if g.get("don_gia") is not None else ws.cell(r, 9).value


def build(conn, ship_id, inv_id=None, sheets=None):
    """Điền biểu mẫu cho shipment. Trả về (Workbook chỉ gồm các sheet yêu cầu, danh sách ô không tính được)."""
    sheets = [s for s in (sheets or [f[0] for f in FORMS]) if s in KIND]
    wb = openpyxl.load_workbook(TEMPLATE)
    inv_id = inv_id or ship_id
    overrides, statics = {}, {}
    for sheet in wb.sheetnames:
        for (c, r) in _key_cells(wb[sheet], sheet):
            val = ship_id if KIND.get(sheet) == "ship" else inv_id
            overrides[(sheet, c, r)] = val
            statics[(sheet, c, r)] = val
    ctx = _Ctx(conn, wb, overrides)
    results, errors = {}, []
    for sheet in sheets:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    try:
                        results[(sheet, cell.coordinate)] = ctx.cell_value(sheet, cell.column, cell.row)
                    except XlError as e:
                        results[(sheet, cell.coordinate)] = ""
                        errors.append({"sheet": sheet, "cell": cell.coordinate, "error": e.code})
    for name in list(wb.sheetnames):
        if name not in sheets:
            del wb[name]
    for (sheet, coord), v in results.items():
        cell = wb[sheet][coord]
        cell.value = v
        if isinstance(v, dt.datetime) and cell.number_format == "General":
            cell.number_format = "dd/mm/yyyy" if (v.hour, v.minute) == (0, 0) else "dd/mm/yyyy hh:mm"
    for (sheet, c, r), v in statics.items():
        if sheet in sheets:
            wb[sheet].cell(r, c).value = v
    if "INV" in sheets:
        _fill_invoice_lines(conn, wb, inv_id)
    return wb, errors


DRAWINGS = Path(__file__).parent / "templates" / "drawings"


def inject_drawings(data, sheet_names):
    """openpyxl bỏ các hình vẽ (shape) khi lưu — vd khung viền đỏ của DGD WH (hàng trăm shape). Ghép lại XML hình vẽ gốc vào file xuất."""
    todo = [(n, DRAWINGS / f"{n}.xml") for n in sheet_names if (DRAWINGS / f"{n}.xml").exists()]
    if not todo:
        return data
    zin = zipfile.ZipFile(io.BytesIO(data))
    files = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    wbx = files["xl/workbook.xml"].decode("utf8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf8")
    target = {}
    for m in re.finditer(r"<Relationship [^>]*/?>", rels):
        i, t = re.search(r"Id=\"(rId\d+)\"", m.group()), re.search(r"Target=\"([^\"]+)\"", m.group())
        if i and t:
            target[i.group(1)] = t.group(1)
    ct = files["[Content_Types].xml"].decode("utf8")
    for k, (name, path) in enumerate(todo):
        m = re.search(r"<sheet [^>]*name=\"%s\"[^>]*r:id=\"(rId\d+)\"" % re.escape(name), wbx)
        if not m:
            continue
        sp = target[m.group(1)].lstrip("/")
        sheet_path = sp if sp.startswith("xl/") else "xl/" + sp
        sx = files[sheet_path].decode("utf8")
        if "<drawing " in sx:  # sheet đã có drawing (hình ảnh) — không ghép đè
            continue
        dname = f"xl/drawings/drawing_orig{k + 1}.xml"
        files[dname] = path.read_bytes()
        if "xmlns:r=" not in sx[:600]:
            sx = sx.replace("<worksheet ", "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ", 1)
        tag = "<drawing r:id=\"rIdOrigDrw\"/>"
        for anchor in ("<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"):
            if anchor in sx:
                sx = sx.replace(anchor, tag + anchor, 1)
                break
        files[sheet_path] = sx.encode("utf8")
        rel_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rel = f"<Relationship Id=\"rIdOrigDrw\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing_orig{k + 1}.xml\"/>"
        if rel_path in files:
            files[rel_path] = files[rel_path].decode("utf8").replace("</Relationships>", rel + "</Relationships>").encode("utf8")
        else:
            files[rel_path] = ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + rel + "</Relationships>").encode("utf8")
        ct = ct.replace("</Types>", f"<Override PartName=\"/{dname}\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/></Types>")
    files["[Content_Types].xml"] = ct.encode("utf8")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    return out.getvalue()


def to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return inject_drawings(buf.getvalue(), wb.sheetnames)


def preview_html(conn, ship_id, inv_id, sheet):
    wb, errors = build(conn, ship_id, inv_id, [sheet])
    ws = wb[sheet]
    html_ = rvc_form.to_html(ws, bounds=rvc_form.print_bounds(ws))
    if (DRAWINGS / f"{sheet}.xml").exists():
        # biểu mẫu gốc có khung viền đỏ (hình vẽ) — bản xem trên web vẽ khung sọc đỏ tương tự quanh biểu mẫu
        html_ = ("<div style=\"display:inline-block;padding:10px;border:16px solid transparent;"
                 "border-image:repeating-linear-gradient(135deg,#fc0d2a 0 6px,#fff 6px 12px) 16\">" + html_ + "</div>")
    return html_, errors


def shipment_ids(conn):
    return [r["key"] for r in conn.execute("SELECT key FROM ob_rows WHERE tbl='source' AND key IS NOT NULL ORDER BY pos")]


def invoice_ids(conn):
    """Mã dùng cho PK / INV: tiền tố của 'Separate ID#' trong List shipment (vd VNBRINT00091 từ VNBRINT00091-44)."""
    seen = {}
    for r in conn.execute("SELECT key FROM ob_rows WHERE tbl='list' AND key IS NOT NULL ORDER BY pos"):
        k = r["key"]
        p = k.rsplit("-", 1)[0] if "-" in k else k
        seen[p] = seen.get(p, 0) + 1
    for r in conn.execute("SELECT so_to FROM ob_invoices WHERE so_to IS NOT NULL ORDER BY id"):
        seen.setdefault(r["so_to"], 1)
    return list(seen)
