"""Quét Commercial Invoice (Excel / PDF có chữ) -> thông tin Invoice + các dòng hàng, để điền tự động thay vì nhập tay.

Nhận dạng theo NHÃN của mẫu Invoice Polarium ("Invoice Number", "Invoice Date", "Polarium SO/TO number", "Polarium Ref (Shipment ID)",
"INCOTERMS", "Mode of Shipment", bảng ITEM / DESCRIPTION / QTY / UNIT / Country of origin / Unit Price (CCY) / Total Amount (CCY)...)
chứ không theo vị trí ô cố định, nên vẫn đọc được khi người dùng chèn / bớt dòng.
"""
import datetime as dt
import re

import openpyxl

from db import fold

HEADER_KEYS = ["invoice_no", "invoice_date", "customer_po", "so_to", "polarium_ref", "incoterms", "mode", "carrier", "shipped_on", "hs_code",
               "term_of_payment", "remarks", "sold_by", "sold_to", "shipper", "deliver_to"]
# (khoá, regex nhãn (đã bỏ dấu, chữ thường), cách lấy giá trị theo thứ tự ưu tiên)
LABELS = [
    ("invoice_no", r"^invoice\s*(number|no\.?)\b", ["inline", "below", "right"]),
    ("invoice_date", r"^invoice\s*date\b", ["inline", "below", "right"]),
    ("customer_po", r"^customer\s*po\b", ["inline", "right"]),
    ("so_to", r"^polarium\s*so\s*(/\s*to)?\s*(number|no)\b", ["inline", "right"]),
    ("polarium_ref", r"^polarium\s*ref\b", ["inline", "right"]),
    ("incoterms", r"^incoterms?\b", ["inline", "right"]),
    ("mode", r"^mode\s*of\s*shipment\b", ["inline"]),
    ("carrier", r"^carrier\b", ["inline", "below"]),
    ("shipped_on", r"^shipped\s*on\b", ["inline", "below"]),
    ("hs_code", r"^hs\s*code\b", ["inline", "below"]),
    ("term_of_payment", r"^term\s*of\s*payment\b", ["inline", "below", "right"]),
    ("remarks", r"^remarks\b", ["inline", "below"]),
    ("sold_by", r"^sold\s*by\b", ["block"]),
    ("sold_to", r"^(sold|bill)\s*to\b", ["block"]),
    ("shipper", r"^shipper\b", ["block"]),
    ("deliver_to", r"^deliver\s*to\b", ["block"]),
]
LINE_ALIASES = {
    "stt": [r"^item\s*no\.?$", r"^no\.?$", r"^#$"],
    "ma_hang": [r"^item(\s*code)?$", r"^part\s*(no|number)\.?$"],
    "ten_hang": [r"^description"],
    "qty": [r"^qty\b", r"^quantity"],
    "dvt": [r"^unit$", r"^uom$"],
    "xuat_xu": [r"^country\s*of\s*origin", r"^coo$", r"^origin$"],
    "hs": [r"^hs\s*code"],
    "don_gia": [r"^unit\s*price"],
    "amount": [r"^total\s*amount", r"^amount"],
}


def _t(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).replace("\xa0", " ").strip()


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.,\-]", "", str(v).replace("\xa0", ""))
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _date(v):
    """-> ISO yyyy-mm-dd hoặc None. Chuỗi m/d/yyyy (kiểu Mỹ như PDF Polarium) hoặc d/m/yyyy khi vế đầu > 12."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = _t(v)
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})", s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        d, mo = (a, b) if a > 12 else (b, a)  # mặc định m/d/yyyy (PDF Polarium: 8/26/2026)
        try:
            return dt.date(y, mo, d).isoformat()
        except ValueError:
            return None
    return None


def _fold(s):
    return fold(s).strip()


def _label_of(text):
    """Nhãn nhận dạng được của một ô (khoá) hoặc None."""
    f = _fold(text).rstrip(":").strip()
    for key, rx, _ in LABELS:
        if re.match(rx, f):
            return key
    return None


# ----------------------------------------------------------------------------------------------- Excel
def _grid(ws):
    return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if c.value not in (None, "")}


def _find_table_header(g, max_row):
    """Dòng tiêu đề bảng hàng: có cột ITEM, DESCRIPTION, QTY."""
    for r in range(1, max_row + 1):
        cols = {c: _fold(v) for (rr, c), v in g.items() if rr == r and isinstance(v, str)}
        has = lambda rx: any(re.match(rx, t) for t in cols.values())
        if has(r"^item(\s*code)?$") and has(r"^qty\b|^quantity") and has(r"^description"):
            return r
    return None


def invoice_sheets(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb:
        g = _grid(ws)
        title = any("commercial invoice" in _fold(v) for (r, c), v in g.items() if r <= 12 and isinstance(v, str))
        if _find_table_header(g, min(ws.max_row, 60)) and title:
            out.append(ws.title)
    return out


def scan_xlsx(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    cands = invoice_sheets(path)
    if not cands:
        raise ValueError("Không tìm thấy sheet Commercial Invoice (cần tiêu đề COMMERCIAL INVOICE và bảng có cột ITEM, DESCRIPTION, QTY)")
    pick = sheet if sheet in cands else next((n for n in cands if "customs" not in n.lower()), cands[0])
    ws = wb[pick]
    g = _grid(ws)
    hdr_row = _find_table_header(g, min(ws.max_row, 60))
    warnings = []
    header = {k: None for k in HEADER_KEYS}
    labels = {(r, c): _label_of(v) for (r, c), v in g.items() if isinstance(v, str) and r < hdr_row}
    labels = {k: v for k, v in labels.items() if v}

    def is_label(pos):
        v = g.get(pos)
        return pos in labels or (isinstance(v, str) and re.match(r"^(reference|remarks|bill to|sold to|deliver to|shipper|sold by|term of payment|carrier|shipped on|hs code|incoterms?|mode of)\b", _fold(v).rstrip(":")))

    for (r, c), key in sorted(labels.items()):
        if header[key]:
            continue
        mode = next(m for k, rx, m in LABELS if k == key)
        raw = _t(g[(r, c)])
        for how in mode:
            val = None
            if how == "inline":
                after = re.split(r":", raw, maxsplit=1)
                if len(after) == 2 and after[1].strip():
                    val = after[1].strip()
                else:
                    m = re.match(r"^(?:invoice\s*(?:number|no\.?|date)|mode\s*of\s*shipment|polarium[^:]*|customer\s*po|incoterms?)\s+(\S.*)$", raw, re.I)
                    val = m.group(1).strip() if m and key in ("mode",) else None
            elif how == "right":
                for cc in range(c + 1, c + 7):
                    if (r, cc) in g and not is_label((r, cc)):
                        val = g[(r, cc)]
                        break
                    if (r, cc) in g:
                        break
            elif how in ("below", "block"):
                for rr in range(r + 1, r + 4):
                    if (rr, c) in g and not is_label((rr, c)):
                        val = g[(rr, c)]
                        break
                    if (rr, c) in g:
                        break
            if val not in (None, ""):
                header[key] = val
                break
    for k in ("invoice_date", "shipped_on"):
        if header[k] is not None:
            header[k] = _date(header[k]) or _t(header[k])
    for k in HEADER_KEYS:
        if k not in ("invoice_date", "shipped_on") and header[k] is not None:
            header[k] = _t(header[k]) or None

    # ---- bảng dòng hàng ----
    cols, ccy = {}, None
    for (r, c), v in g.items():
        if r == hdr_row and isinstance(v, str):
            f = _fold(v)
            for field, rxs in LINE_ALIASES.items():
                if field not in cols and any(re.match(rx, f) for rx in rxs):
                    cols[field] = c
            m = re.search(r"\(([A-Za-z]{3})\)", v)
            if m and ("price" in f or "amount" in f) and not ccy:
                ccy = m.group(1).upper()
    lines, total, total_ccy = [], None, None
    for r in range(hdr_row + 1, ws.max_row + 1):
        rowtxt = " ".join(_fold(v) for (rr, c), v in g.items() if rr == r and isinstance(v, str))
        if "total value of goods" in rowtxt:
            for (rr, c), v in sorted(g.items()):
                if rr == r:
                    if isinstance(v, str) and re.fullmatch(r"[A-Za-z]{3}", v.strip()):
                        total_ccy = v.strip().upper()
                    elif _num(v) is not None and not isinstance(v, str):
                        total = _num(v)
            break
        get = lambda f: g.get((r, cols[f])) if f in cols else None
        item, desc, qty = _t(get("ma_hang")), _t(get("ten_hang")), _num(get("qty"))
        if not item:
            continue
        lines.append({"stt": _t(get("stt")), "ma_hang": item or None, "ten_hang": desc or None, "qty": qty, "dvt": _t(get("dvt")) or None,
                      "xuat_xu": _t(get("xuat_xu")) or None, "don_gia": _num(get("don_gia")), "amount": _num(get("amount")),
                      "hs": _t(get("hs")) or None, "row": r})
    if not lines:
        warnings.append("Không đọc được dòng hàng nào dưới tiêu đề bảng")
    return _finish({"source": "xlsx", "sheet": pick, "sheets": cands, "header": header, "lines": lines, "currency": total_ccy or ccy,
                    "total": total, "warnings": warnings})


# ----------------------------------------------------------------------------------------------- PDF
def scan_pdf(path):
    try:
        import pdfplumber
    except ImportError:
        raise ValueError("Cần cài thư viện đọc PDF: mở cửa sổ lệnh và chạy  pip install pdfplumber")
    texts, tables = [], []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            texts.append(p.extract_text() or "")
            tables.extend(p.extract_tables() or [])
    text = "\n".join(texts)
    if len(text.strip()) < 40:
        raise ValueError("PDF không có chữ để đọc (có thể là bản scan ảnh). Hệ thống chưa hỗ trợ OCR — hãy dùng file PDF xuất trực tiếp từ hệ thống hoặc file Excel.")
    header = {k: None for k in HEADER_KEYS}
    cells = [str(c) for t in tables for row in t for c in row if c]
    rx = lambda pat, src=text, flags=re.I: (re.search(pat, src, flags) or [None, None])[1]
    for cell in cells:
        head, _, rest = cell.partition("\n")
        k = _label_of(head)
        if k in ("sold_by", "sold_to", "shipper", "deliver_to") and rest:
            header[k] = rest.replace("\n", ", ").strip(", ")
        elif k in ("invoice_no", "invoice_date", "term_of_payment") and rest:
            header[k] = rest.strip()
        elif k == "hs_code" and rest.strip():
            header["hs_code"] = rest.strip()
        elif cell.lower().startswith("remarks"):
            header["remarks"] = rest.replace("\n", "; ").strip()
    header["customer_po"] = rx(r"Customer PO:\s*([^\n]+)")
    header["so_to"] = rx(r"Polarium SO(?:/TO)?\s*(?:number|no\.?):\s*([^\n]+?)(?:\s+Contact:|\n|$)")
    header["polarium_ref"] = rx(r"Polarium Ref\s*\(Shipment ID\):\s*(\S+)")
    header["incoterms"] = rx(r"INCOTERMS:\s*([A-Z]{3})")
    header["mode"] = rx(r"Mode of Shipment:\s*([A-Za-z]+)")
    header["carrier"] = header["carrier"] or rx(r"Carrier:\s*([^\n]*?)\s*Shipped on")
    if not header["invoice_no"]:
        m = re.search(r"Invoice Number\s+Invoice Date\s*\n(\S+)\s+(\S+)", text)
        if m:
            header["invoice_no"], header["invoice_date"] = m.group(1), m.group(2)
    if header["invoice_no"] and "\n" in str(header["invoice_no"]):
        header["invoice_no"] = header["invoice_no"].split("\n")[0]
    if header["invoice_date"]:
        header["invoice_date"] = _date(header["invoice_date"]) or header["invoice_date"]
    if not header["hs_code"]:
        header["hs_code"] = rx(r"HS code:\s*\n?\s*(\d{6,10})")
    for k in HEADER_KEYS:
        if isinstance(header[k], str):
            header[k] = header[k].strip() or None
    # ---- bảng dòng hàng ----
    lines, ccy = [], None
    for t in tables:
        if not t or not t[0]:
            continue
        hdr = [_fold(str(c or "").replace("\n", " ")) for c in t[0]]
        if not (any(re.match(r"^item$", h) or h == "item code" for h in hdr) and any(h.startswith("qty") for h in hdr)):
            continue
        cols = {}
        for i, h in enumerate(hdr):
            for field, rxs in LINE_ALIASES.items():
                if field not in cols and any(re.match(rx_, h) for rx_ in rxs):
                    cols[field] = i
            m = re.search(r"\(([a-z]{3})\)", h)
            if m and not ccy:
                ccy = m.group(1).upper()
        for n, row in enumerate(t[1:], start=2):
            get = lambda f: row[cols[f]] if f in cols and cols[f] < len(row) else None
            item, desc, qty = _t(get("ma_hang")), _t(get("ten_hang")).replace("\n", " "), _num(get("qty"))
            if not item:
                continue
            lines.append({"stt": _t(get("stt")), "ma_hang": item or None, "ten_hang": desc or None, "qty": qty, "dvt": _t(get("dvt")) or None,
                          "xuat_xu": _t(get("xuat_xu")) or None, "don_gia": _num(get("don_gia")), "amount": _num(get("amount")), "hs": None, "row": n})
    m = re.search(r"Total Value of goods:\s*([A-Z]{3})\s*([\d.,]+)", text, re.I)
    warnings = [] if lines else ["Không đọc được bảng dòng hàng trong PDF"]
    return _finish({"source": "pdf", "sheet": None, "sheets": [], "header": header, "lines": lines,
                    "currency": (m.group(1).upper() if m else ccy), "total": _num(m.group(2)) if m else None, "warnings": warnings})


# ----------------------------------------------------------------------------------------------- kiểm tra chéo
def _finish(res):
    lines = res["lines"]
    for l in lines:
        if l["amount"] is None and l["qty"] is not None and l["don_gia"] is not None:
            l["amount"] = round(l["qty"] * l["don_gia"], 2)
    s = round(sum(l["amount"] or 0 for l in lines), 2)
    res["sum_lines"] = s
    res["total_match"] = (res["total"] is None) or abs(s - res["total"]) < 0.05 + 1e-6 * abs(res["total"] or 0)
    w = res["warnings"]
    h = res["header"]
    if not h.get("polarium_ref"):
        w.append("Không thấy 'Polarium Ref (Shipment ID)' — hãy nhập Shipment ID")
    if not h.get("invoice_no"):
        w.append("Không đọc được Invoice Number")
    if res["total"] is not None and not res["total_match"]:
        w.append(f"Tổng các dòng ({s:,.2f}) khác Total Value of goods ({res['total']:,.2f}) — kiểm tra lại số liệu")
    for l in lines:
        if l["qty"] is not None and l["don_gia"] is not None and l["amount"] is not None and abs(l["qty"] * l["don_gia"] - l["amount"]) > 0.05 + 1e-4 * abs(l["amount"]):
            w.append(f"Dòng {l.get('stt') or l['row']}: Qty × Unit price ≠ Amount trong file")
            break
    return res


def scan(path, name="", sheet=None):
    ext = (name or str(path)).lower().rsplit(".", 1)[-1]
    if ext in ("xlsx", "xlsm"):
        return scan_xlsx(path, sheet)
    if ext == "pdf":
        return scan_pdf(path)
    raise ValueError("Chỉ quét được file Excel (.xlsx) hoặc PDF có chữ")
