"""Hệ thống Outbound shipment: dữ liệu (Source_Data, List shipment, Tracking, Return shipment, 2025 Outbound) lưu trong SQLite.

Mỗi bảng là một sheet của file 'Database outbound shipping plan.xlsx'. Ô lưu theo CHỮ CỘT của Excel ({"B": "VNOUT00001", ...})
để các công thức gốc của biểu mẫu (VLOOKUP(..., Source_Data!B:BH, 5, 0)) tính được nguyên vẹn.
"""
import datetime as dt
import json
import os
import re
import uuid
from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string, get_column_letter

import db
import grid
from db import fold
from importer import _col_type

TABLES = {
    "source": {"title": "Shipment (Source data — SI / BL / container)", "sheet": "Source_Data", "header": 1, "first": 4, "key": "B", "tag_row": 2},
    "list": {"title": "List shipment", "sheet": "List shipment", "header": 2, "first": 3, "key": "B"},
    "tracking": {"title": "Tracking xe / container", "sheet": "Tracking", "header": 3, "first": 4, "key": "C", "tag_row": 2},
    "return": {"title": "Return shipment", "sheet": "Return shipment", "header": 2, "first": 3, "key": "B"},
    "y2025": {"title": "2025 Outbound", "sheet": "2025 Outbound", "header": 1, "first": 2, "key": "D"},
}
SHEET_TO_TABLE = {v["sheet"]: k for k, v in TABLES.items()}

# Cột do hệ thống thêm (không có trong file Excel gốc). Khoá bắt đầu bằng '+' nên KHÔNG nằm trong vùng B:BH mà biểu mẫu tra cứu
# -> thêm / sửa các cột này không ảnh hưởng dữ liệu của biểu mẫu. Được giữ nguyên khi nạp lại file Excel.
EXTRA_COLS = {
    "source": [("+SoTK", "Số tờ khai", "text"), ("+NgayTK", "Ngày tờ khai", "date"), ("+Invoice", "Số Invoice", "text"),
               ("+TypeBL", "Type of BL", "text"), ("+UN", "UN", "text"), ("+ActualPU", "Actual pick up", "date")],
}
FILE_COL = {"letter": "_files", "name": "File", "type": "number", "tag": "", "computed": True}
GOODS_COL = {"letter": "_goods", "name": "Dòng hàng (Invoice)", "type": "number", "tag": "", "computed": True}


def ship_of(tbl, cells, key, cols=None):
    """Mã shipment mà một dòng thuộc về (khoá để gắn các dòng hàng): Source = Shipment ID; List shipment = cột Shipment# (G),
    nếu trống thì tiền tố của Separate ID# (VNOUT00004-1 -> VNOUT00004)."""
    if tbl == "source":
        return key
    if tbl == "list":
        v = cells.get("G")
        return str(v).strip() if v not in (None, "") else (key.rsplit("-", 1)[0] if key and "-" in key else key)
    return None
ALLOWED_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".txt", ".csv", ".xlsx", ".xls", ".docx", ".doc"}
INLINE_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}  # mở xem trực tiếp trên trình duyệt (không có HTML/SVG)
MAX_FILE = 30 * 1024 * 1024
CURRENCIES = ["USD", "EUR", "SEK"]
_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?$")


def enc(v):
    """Giá trị ô -> giá trị JSON (ngày/giờ -> chuỗi)."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if (v.hour, v.minute, v.second) == (0, 0, 0) else v.isoformat(timespec="seconds")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    if isinstance(v, str):
        return v.strip() if v.strip() else None
    return v


def dec(v):
    """JSON -> giá trị cho bộ tính công thức (chuỗi ISO -> datetime)."""
    if isinstance(v, str) and _ISO.match(v):
        try:
            return dt.datetime.fromisoformat(v)
        except ValueError:
            return v
    return v


def _clean(h):
    return " ".join(str(h).split()) if h is not None else ""


# ----------------------------------------------------------------------------------------------- nạp từ Excel
def read_sheet(ws, spec):
    """-> (cols, rows[{letter: value}])"""
    all_rows = list(ws.iter_rows(values_only=True))
    hdr = all_rows[spec["header"] - 1] if len(all_rows) >= spec["header"] else ()
    tags = all_rows[spec["tag_row"] - 1] if spec.get("tag_row") and len(all_rows) >= spec["tag_row"] else ()
    ncols = max((len(r) for r in all_rows[spec["header"] - 1:]), default=0)
    seen, cols = {}, []
    data = [r for r in all_rows[spec["first"] - 1:] if any(c not in (None, "") for c in r)]
    for i in range(ncols):
        letter = get_column_letter(i + 1)
        name = _clean(hdr[i]) if i < len(hdr) else ""
        vals = [r[i] if i < len(r) else None for r in data]
        if not name and not any(v not in (None, "") for v in vals):
            continue
        name = name or f"Cột {letter}"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name} ({seen[name]})"
        typed = [v if isinstance(v, (dt.datetime, dt.date)) else enc(v) for v in vals]
        cols.append({"letter": letter, "name": name, "type": _col_type(typed),
                     "tag": _clean(tags[i]) if i < len(tags) and tags[i] in ("Logs", "WH") else ""})
    rows = []
    for r in data:
        cells = {}
        for c in cols:
            i = column_index_from_string(c["letter"]) - 1
            v = enc(r[i]) if i < len(r) else None
            if v is not None:
                cells[c["letter"]] = v
        rows.append(cells)
    return cols, rows


def sheet_candidates(path):
    """Các sheet trong file khớp với 5 bảng dữ liệu của hệ thống."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return {tbl: spec["sheet"] for tbl, spec in TABLES.items() if spec["sheet"] in wb.sheetnames}, list(wb.sheetnames)
    finally:
        wb.close()


def import_workbook(conn, path, selection):
    """selection: {bảng: tên sheet}. Thay toàn bộ dữ liệu của các bảng được chọn."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    try:
        for tbl, sheet in selection.items():
            if tbl not in TABLES:
                continue
            if sheet not in wb.sheetnames:
                raise ValueError(f"Không có sheet '{sheet}' trong file")
            spec = TABLES[tbl]
            cols, rows = read_sheet(wb[sheet], spec)
            if not rows:
                raise ValueError(f"Sheet '{sheet}' không có dữ liệu (dòng dữ liệu bắt đầu từ dòng {spec['first']})")
            before = conn.execute("SELECT COUNT(*) FROM ob_rows WHERE tbl=?", (tbl,)).fetchone()[0]
            # giữ lại cột do hệ thống thêm (theo khoá dòng) và file đính kèm (theo khoá + thứ tự xuất hiện) qua lần nạp thay thế
            old_cells, old_files = _snapshot(conn, tbl)
            cols = cols + [{"letter": k, "name": n, "type": t, "tag": "", "extra": True} for k, n, t in EXTRA_COLS.get(tbl, [])]
            with conn:
                conn.execute("DELETE FROM ob_files WHERE row_id IN (SELECT id FROM ob_rows WHERE tbl=?)", (tbl,))
                conn.execute("DELETE FROM ob_rows WHERE tbl=?", (tbl,))
                conn.execute("INSERT INTO ob_tables(name,title,sheet,key_col,cols) VALUES(?,?,?,?,?) ON CONFLICT(name) DO UPDATE SET "
                             "title=excluded.title, sheet=excluded.sheet, key_col=excluded.key_col, cols=excluded.cols, updated_at=datetime('now','localtime')",
                             (tbl, spec["title"], sheet, spec["key"], json.dumps(cols, ensure_ascii=False)))
                seen_key, new_ids = {}, {}
                for pos, cells in enumerate(rows, start=1):
                    key = cells.get(spec["key"])
                    key = None if key is None else str(key)
                    klc = key.lower() if key else None
                    if klc:
                        n = seen_key[klc] = seen_key.get(klc, 0) + 1
                        cells.update(old_cells.get((klc, n), {}))
                    rid = conn.execute("INSERT INTO ob_rows(tbl,pos,key,key_lc,cells) VALUES(?,?,?,?,?)",
                                       (tbl, pos, key, klc, json.dumps(cells, ensure_ascii=False))).lastrowid
                    if klc:
                        new_ids[(klc, seen_key[klc])] = rid
                orphans = 0
                for (klc_n, files) in old_files.items():
                    rid = new_ids.get(klc_n)
                    for f in files:
                        if rid:
                            conn.execute("INSERT INTO ob_files(row_id,name,stored,size,mime,at) VALUES(?,?,?,?,?,?)", (rid, f["name"], f["stored"], f["size"], f["mime"], f["at"]))
                        else:
                            orphans += 1
            out[tbl] = {"sheet": sheet, "before": before, "after": len(rows), "columns": len(cols), "files_orphaned": orphans}
    finally:
        wb.close()
    return out


# ----------------------------------------------------------------------------------------------- truy vấn
def table_info(conn, tbl):
    r = conn.execute("SELECT * FROM ob_tables WHERE name=?", (tbl,)).fetchone()
    if r is None:
        return None
    d = dict(r)
    d["cols"] = json.loads(d["cols"])
    have = {c["letter"] for c in d["cols"]}
    missing = [(k, n, t) for k, n, t in EXTRA_COLS.get(tbl, []) if k not in have]
    if missing:  # bảng nạp từ trước khi có cột mới: tự bổ sung cột ở cuối bảng
        d["cols"] += [{"letter": k, "name": n, "type": t, "tag": "", "extra": True} for k, n, t in missing]
        with conn:
            conn.execute("UPDATE ob_tables SET cols=? WHERE name=?", (json.dumps(d["cols"], ensure_ascii=False), tbl))
    return d


def _snapshot(conn, tbl):
    """Trước khi nạp thay thế: giá trị cột do hệ thống thêm và file đính kèm của từng dòng, khoá theo (khoá dòng, thứ tự xuất hiện)."""
    cells, files, seen = {}, {}, {}
    for r in conn.execute("SELECT id, key_lc, cells FROM ob_rows WHERE tbl=? AND key_lc IS NOT NULL ORDER BY pos", (tbl,)):
        n = seen[r["key_lc"]] = seen.get(r["key_lc"], 0) + 1
        extras = {k: v for k, v in json.loads(r["cells"]).items() if k.startswith("+")}
        if extras:
            cells[(r["key_lc"], n)] = extras
        fl = [dict(f) for f in conn.execute("SELECT name, stored, size, mime, at FROM ob_files WHERE row_id=?", (r["id"],))]
        if fl:
            files[(r["key_lc"], n)] = fl
    return cells, files


def overview(conn):
    out = []
    for tbl, spec in TABLES.items():
        info = table_info(conn, tbl)
        n = conn.execute("SELECT COUNT(*) FROM ob_rows WHERE tbl=?", (tbl,)).fetchone()[0]
        out.append({"name": tbl, "title": spec["title"], "sheet": spec["sheet"], "rows": n,
                    "updated_at": info["updated_at"] if info else None, "loaded": bool(info)})
    return out


def find_row(conn, tbl, key):
    """Dòng có khoá = key (không phân biệt hoa thường) -> {chữ cột: giá trị} hoặc None. Dùng cho VLOOKUP."""
    r = conn.execute("SELECT cells FROM ob_rows WHERE tbl=? AND key_lc=? ORDER BY pos LIMIT 1", (tbl, str(key).lower())).fetchone()
    return None if r is None else {k: dec(v) for k, v in json.loads(r["cells"]).items()}


_PGF = [False]   # đang chạy trên PostgreSQL: cột số trong JSON phải ép kiểu số


def _expr_for(cols):
    by_name = {c["name"]: c for c in cols}

    def expr(name, ctype, params):
        params.append('$."%s"' % by_name[name]["letter"])
        return ("json_num" if (_PGF[0] and by_name[name]["type"] == "number") else "json_extract") + "(t.cells, ?)"
    return expr


def _where(conn, tbl, filters, q, cols):
    _PGF[0] = db.is_pg(conn)
    conn.create_function("fold", 1, fold, deterministic=True)
    types = {c["name"]: c["type"] for c in cols}
    expr = _expr_for(cols)
    conds, params = ["t.tbl = ?"], [tbl]
    for name, text in (filters or {}).items():
        if name in types:
            c = grid._condition(name, types[name], text, params, expr)
            if c:
                conds.append(c)
    if q and q.strip():
        conds.append("instr(fold(t.cells), ?) > 0")
        params.append(fold(q.strip()))
    return "WHERE " + " AND ".join(conds), params


def query(conn, tbl, filters, q, sort, direction, page, size):
    info = table_info(conn, tbl)
    if info is None:
        return {"total": 0, "columns": [], "rows": [], "ids": [], "col_info": []}
    cols = info["cols"]
    where, wp = _where(conn, tbl, filters, q, cols)
    total = conn.execute(f"SELECT COUNT(*) FROM ob_rows t {where}", wp).fetchone()[0]
    order, op = "ORDER BY t.pos", []
    if sort in {c["name"] for c in cols}:
        ex = _expr_for(cols)(sort, None, op)
        op = op * 2
        order = f"ORDER BY ({ex}) IS NULL, ({ex}) {'DESC' if direction == 'desc' else 'ASC'}, t.pos"
    rows = conn.execute(f"SELECT t.id AS _id, t.cells AS _cells FROM ob_rows t {where} {order} LIMIT ? OFFSET ?",
                        wp + op + [size, (page - 1) * size]).fetchall()
    ids = [r["_id"] for r in rows]
    counts = {}
    if ids:
        marks = ",".join("?" * len(ids))
        counts = {x["row_id"]: x["n"] for x in conn.execute(
            f"SELECT row_id, COUNT(*) n FROM ob_files WHERE row_id IN ({marks}) GROUP BY row_id", ids)}
    data = []
    gcount = {}
    with_goods = tbl in ("source", "list")
    if with_goods and rows:
        sl = set()
        for r in rows:
            c = json.loads(r["_cells"])
            s = ship_of(tbl, c, c.get(info["key_col"]))
            if s:
                sl.add(s.lower())
        if sl:
            marks2 = ",".join("?" * len(sl))
            gcount = {x["ship_lc"]: x["n"] for x in conn.execute(
                f"SELECT ship_lc, COUNT(*) n FROM ob_goods WHERE ship_lc IN ({marks2}) GROUP BY ship_lc", list(sl))}
    for r in rows:
        cells = json.loads(r["_cells"])
        line = [cells.get(c["letter"]) for c in cols] + [counts.get(r["_id"])]
        if with_goods:
            s = ship_of(tbl, cells, cells.get(info["key_col"]))
            line.append(gcount.get(s.lower()) if s else None)
        data.append(line)
    cols_out = cols + [FILE_COL] + ([GOODS_COL] if with_goods else [])
    return {"total": total, "columns": [c["name"] for c in cols_out], "col_info": cols_out, "rows": data, "ids": ids}


def distinct(conn, tbl, name, filters, q, limit=200):
    info = table_info(conn, tbl)
    cols = info["cols"] if info else []
    if name not in {c["name"] for c in cols}:
        return []
    others = {k: v for k, v in (filters or {}).items() if k != name}
    where, wp = _where(conn, tbl, others, q, cols)
    p = []
    ex = _expr_for(cols)(name, None, p)
    sql = f"SELECT v, COUNT(*) n FROM (SELECT {ex} v FROM ob_rows t {where}) s WHERE v IS NOT NULL GROUP BY v ORDER BY n DESC LIMIT ?"
    return [{"v": r["v"], "n": r["n"]} for r in conn.execute(sql, p + wp + [limit])]


def export_xlsx(conn, tbl, filters, q, sort, direction):
    res = query(conn, tbl, filters, q, sort, direction, 1, 200000)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(TABLES[tbl]["sheet"][:31])
    ws.append(res["columns"])
    for r in res["rows"]:
        ws.append([dec(v) if isinstance(v, str) else v for v in r])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------------------------- sửa dữ liệu
def get_row(conn, rid):
    r = conn.execute("SELECT * FROM ob_rows WHERE id=?", (rid,)).fetchone()
    if r is None:
        raise KeyError(rid)
    info = table_info(conn, r["tbl"])
    out = {"id": r["id"], "tbl": r["tbl"], "title": info["title"], "cols": info["cols"], "key": r["key"], "cells": json.loads(r["cells"]),
           "files": files_of(conn, r["id"])}
    cells = out["cells"]
    ship = ship_of(r["tbl"], cells, r["key"])
    out["ship"] = ship
    if ship and r["tbl"] in ("source", "list"):
        out["goods"] = goods_get(conn, ship)
        out["invoices"] = invoices_of(conn, ship)
    return out


def _coerce(v, ctype):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if ctype == "number":
        try:
            return float(str(v).replace(",", "."))
        except ValueError:
            raise ValueError(f"'{v}' không phải là số")
    if ctype == "date" and isinstance(v, str):
        s = v.strip()
        m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?", s)
        if m:
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            iso = f"{year:04d}-{month:02d}-{day:02d}"
            return f"{iso}T{int(m.group(4)):02d}:{m.group(5)}:00" if m.group(4) else iso
        if not _ISO.match(s):
            raise ValueError(f"'{v}' không phải là ngày (dd/mm/yyyy)")
        return s
    return v.strip() if isinstance(v, str) else v


def update_row(conn, rid, changes):
    row = get_row(conn, rid)
    types = {c["letter"]: c["type"] for c in row["cols"]}
    cells = row["cells"]
    for letter, v in changes.items():
        if letter not in types:
            continue
        nv = _coerce(v, types[letter])
        if nv is None:
            cells.pop(letter, None)
        else:
            cells[letter] = nv
    info = table_info(conn, row["tbl"])
    key = cells.get(info["key_col"])
    key = None if key is None else str(key)
    with conn:
        conn.execute("UPDATE ob_rows SET cells=?, key=?, key_lc=? WHERE id=?",
                     (json.dumps(cells, ensure_ascii=False), key, key.lower() if key else None, rid))
    return {"saved": True}


def add_row(conn, tbl, changes):
    info = table_info(conn, tbl)
    if info is None:
        raise ValueError("Bảng chưa có dữ liệu — hãy nạp file Excel trước")
    pos = (conn.execute("SELECT MAX(pos) FROM ob_rows WHERE tbl=?", (tbl,)).fetchone()[0] or 0) + 1
    with conn:
        rid = conn.execute("INSERT INTO ob_rows(tbl,pos,key,key_lc,cells) VALUES(?,?,NULL,NULL,'{}')", (tbl, pos)).lastrowid
    update_row(conn, rid, changes)
    return rid


def delete_row(conn, rid):
    for f in files_of(conn, rid):
        _unlink_if_unused(conn, f["id"])
    with conn:
        conn.execute("DELETE FROM ob_files WHERE row_id=?", (rid,))
        conn.execute("DELETE FROM ob_rows WHERE id=?", (rid,))


# ----------------------------------------------------------------------------------------------- file đính kèm
def files_dir():
    p = db.DB_PATH.parent / "uploads"
    p.mkdir(parents=True, exist_ok=True)
    return p


def files_of(conn, rid):
    return [dict(r) for r in conn.execute("SELECT id, name, size, mime, at FROM ob_files WHERE row_id=? ORDER BY id", (rid,))]


def add_file(conn, rid, name, data):
    if conn.execute("SELECT 1 FROM ob_rows WHERE id=?", (rid,)).fetchone() is None:
        raise KeyError(rid)
    ext = Path(name or "").suffix.lower()
    if ext not in ALLOWED_EXT:
        raise ValueError(f"Định dạng '{ext or name}' không được phép. Cho phép: hình ảnh, pdf, txt/csv, Excel, Word")
    if not data:
        raise ValueError("File rỗng")
    if len(data) > MAX_FILE:
        raise ValueError("File quá lớn (tối đa 30 MB)")
    stored = f"{uuid.uuid4().hex}{ext}"
    (files_dir() / stored).write_bytes(data)
    import mimetypes
    mime = mimetypes.guess_type(stored)[0] or "application/octet-stream"
    safe = os.path.basename(name).replace("\\", "_")[:150]
    with conn:
        fid = conn.execute("INSERT INTO ob_files(row_id,name,stored,size,mime) VALUES(?,?,?,?,?)", (rid, safe, stored, len(data), mime)).lastrowid
    return {"id": fid, "name": safe, "size": len(data), "mime": mime}


def get_file(conn, fid):
    r = conn.execute("SELECT * FROM ob_files WHERE id=?", (fid,)).fetchone()
    if r is None:
        raise KeyError(fid)
    p = files_dir() / r["stored"]
    if not p.exists():
        raise KeyError("File không còn trên đĩa")
    return dict(r), p


def _unlink_if_unused(conn, fid):
    r = conn.execute("SELECT stored FROM ob_files WHERE id=?", (fid,)).fetchone()
    if r and conn.execute("SELECT COUNT(*) FROM ob_files WHERE stored=?", (r["stored"],)).fetchone()[0] <= 1:
        try:
            (files_dir() / r["stored"]).unlink()
        except OSError:
            pass


def delete_file(conn, fid):
    _unlink_if_unused(conn, fid)
    with conn:
        conn.execute("DELETE FROM ob_files WHERE id=?", (fid,))


# ----------------------------------------------------------------------------------------------- Information of goods
GOODS_FIELDS = ["ma_hang", "ten_hang", "qty", "dvt", "don_gia", "ccy", "xuat_xu", "ter", "net_w", "gross_w"]
GOODS_ALIASES = {
    "ship": ["shipment id", "shipment", "shipment#", "shipment no", "ma shipment", "vnout"],
    "ma_hang": ["ma hang", "ma vat tu", "item", "item no", "item code", "part no", "part number", "material", "sku", "ma"],
    "ten_hang": ["ten hang", "description", "item description", "product description", "ten", "goods description"],
    "qty": ["so luong", "qty", "quantity", "q'ty", "qty (pcs)"],
    "dvt": ["dvt", "unit", "uom", "don vi tinh", "unit of measure"],
    "don_gia": ["don gia", "unit price", "price", "gia", "unit price (usd)", "unit price (eur)", "unit price (sek)"],
    "ccy": ["currency", "tien te", "ccy", "loai tien", "cur"],
    "xuat_xu": ["xuat xu", "origin", "country of origin", "coo", "made in", "country"],
    "amount": ["amount", "thanh tien", "total amount", "value", "total", "line amount"],
    "ter": ["ter of delivery", "incoterm", "incoterms", "term of delivery", "delivery term", "dieu kien giao hang", "term"],
    "net_w": ["net weight", "net weight (kg)", "nw", "net wt", "trong luong tinh", "n.w"],
    "gross_w": ["gross weight", "gross weight (kg)", "gw", "gross wt", "trong luong ca bi", "g.w"],
}


def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", ".")) if isinstance(v, str) else float(v)
    except (TypeError, ValueError):
        return None


def _amount(qty, price, given=None):
    return round(qty * price, 2) if (qty is not None and price is not None) else given


def goods_get(conn, ship):
    out = []
    for r in conn.execute("SELECT * FROM ob_goods WHERE ship_lc=? ORDER BY pos", (str(ship).lower(),)):
        d = dict(r)
        d["amount"] = _amount(d["qty"], d["don_gia"], d["amount"])
        out.append(d)
    return out


def _clean_goods(lines):
    out = []
    for l in lines:
        if not any(str(l.get(k) or "").strip() for k in ("ma_hang", "ten_hang")):
            continue
        qty, price = _num(l.get("qty")), _num(l.get("don_gia"))
        ccy = (l.get("ccy") or "USD").strip().upper() or "USD"
        out.append({"ma_hang": (l.get("ma_hang") or "").strip() or None, "ten_hang": (l.get("ten_hang") or "").strip() or None, "qty": qty,
                    "dvt": (l.get("dvt") or "").strip() or None, "don_gia": price, "ccy": ccy, "xuat_xu": (l.get("xuat_xu") or "").strip() or None,
                    "ter": (l.get("ter") or "").strip() or None, "net_w": _num(l.get("net_w")), "gross_w": _num(l.get("gross_w")),
                    "inv_no": (l.get("inv_no") or None),
                    "amount": _amount(qty, price, _num(l.get("amount")))})
    return out


def goods_save(conn, ship, lines):
    """Thay toàn bộ các dòng hàng của shipment."""
    ship = str(ship)
    rows = _clean_goods(lines)
    with conn:
        conn.execute("DELETE FROM ob_goods WHERE ship_lc=?", (ship.lower(),))
        for pos, d in enumerate(rows, start=1):
            conn.execute("""INSERT INTO ob_goods(ship_key,ship_lc,pos,ma_hang,ten_hang,qty,dvt,don_gia,ccy,xuat_xu,amount,ter,net_w,gross_w,inv_no)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                         (ship, ship.lower(), pos, d["ma_hang"], d["ten_hang"], d["qty"], d["dvt"], d["don_gia"], d["ccy"], d["xuat_xu"],
                          d["amount"], d["ter"], d["net_w"], d["gross_w"], d["inv_no"]))
    return len(rows)


def goods_parse_file(path):
    """Excel 'Information of goods': dòng tiêu đề có Mã hàng + Qty (nhận cả tên cột tiếng Anh); cột Shipment ID tuỳ chọn.
    -> (rows[{ship?, ma_hang, ...}], warnings)"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    hdr_i, cols = None, {}
    for i, r in enumerate(rows[:15]):
        c = {}
        for j, h in enumerate(r):
            f = fold(h).strip() if h else ""
            for field, names in GOODS_ALIASES.items():
                if f in names and field not in c:
                    c[field] = j
        if "ma_hang" in c and "qty" in c:
            hdr_i, cols = i, c
            break
    if hdr_i is None:
        raise ValueError("Không tìm thấy dòng tiêu đề. Cần tối thiểu 2 cột: Mã hàng (Item) và Số lượng (Qty). "
                         "Các cột khác: Tên hàng/Description, ĐVT/Unit, Unit price, Currency, Origin, Amount, Ter of delivery, Net weight, Gross weight, Shipment ID.")
    out, warn = [], []
    for n, r in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
        g = lambda f: r[cols[f]] if f in cols and cols[f] < len(r) else None
        if all(v in (None, "") for v in r):
            continue
        d = {f: g(f) for f in GOODS_FIELDS + ["amount"]}
        d["ship"] = str(g("ship")).strip() if g("ship") not in (None, "") else None
        for f in ("ma_hang", "ten_hang", "dvt", "ccy", "xuat_xu", "ter"):
            d[f] = None if d[f] is None else str(d[f]).strip()
        if d["ccy"] and d["ccy"].upper() not in CURRENCIES:
            warn.append(f"Dòng {n}: tiền tệ '{d['ccy']}' không thuộc USD/EUR/SEK")
        d["row"] = n
        out.append(d)
    if not out:
        raise ValueError("File không có dòng dữ liệu")
    return out, warn


def goods_import_bulk(conn, rows, dry_run=True, mode="replace", default_ship=None):
    """Nạp dòng hàng (Invoice) cho nhiều shipment (cột Shipment ID trong file; dòng không có thì gán default_ship).

    mode='replace': thay toàn bộ dòng hàng cũ của các shipment có trong file; mode='append': thêm vào cuối dòng hàng hiện có."""
    by, problems = {}, []
    for r in rows:
        ship = r.get("ship") or default_ship
        if not ship:
            problems.append({"row": r["row"], "msg": "Thiếu Shipment ID (thêm cột Shipment ID vào file hoặc nhập Shipment mặc định)"})
        else:
            by.setdefault(ship, []).append(r)
    known = {x["key_lc"] for x in conn.execute("SELECT key_lc FROM ob_rows WHERE tbl='source' AND key_lc IS NOT NULL")} | _list_ships(conn)
    ships = []
    for ship, lines in by.items():
        exists = ship.lower() in known
        if not exists:
            problems.append({"row": lines[0]["row"], "level": "warn", "msg": f"Shipment {ship} chưa có trong Source data / List shipment (vẫn lưu)"})
        old = goods_get(conn, ship) if mode == "append" else []
        if not dry_run:
            goods_save(conn, ship, old + lines)
        ships.append({"ship": ship, "lines": len(lines), "exists": exists,
                      "replaces": 0 if mode == "append" else conn.execute("SELECT COUNT(*) FROM ob_goods WHERE ship_lc=?", (ship.lower(),)).fetchone()[0]})
    errs = [p for p in problems if p.get("level") != "warn"]
    return {"ships": ships, "problems": problems, "n_errors": len(errs), "mode": mode}


def _list_ships(conn):
    """Các Shipment# (chữ thường) có trong List shipment."""
    info = table_info(conn, "list")
    if not info:
        return set()
    out = set()
    for r in conn.execute("SELECT key, cells FROM ob_rows WHERE tbl='list'"):
        s = ship_of("list", json.loads(r["cells"]), r["key"])
        if s:
            out.add(s.lower())
    return out


# ----- trang riêng "Dòng hàng (Invoice)": mọi dòng hàng của mọi shipment, tự link sang List shipment theo Shipment# -----
GOODS_COLS = [  # (tên cột hiển thị, cột SQL, kiểu)
    ("Shipment", "ship_key", "text"), ("Item", "ma_hang", "text"), ("Description", "ten_hang", "text"), ("Qty", "qty", "number"),
    ("Unit", "dvt", "text"), ("COO", "xuat_xu", "text"), ("Unit price", "don_gia", "number"), ("Currency", "ccy", "text"),
    ("Amount", "CASE WHEN qty IS NOT NULL AND don_gia IS NOT NULL THEN ROUND(qty*don_gia,2) ELSE amount END", "number"),
    ("Ter of delivery", "ter", "text"), ("Net weight", "net_w", "number"), ("Gross weight", "gross_w", "number")]


def _goods_where(conn, filters, q):
    conn.create_function("fold", 1, fold, deterministic=True)
    by = {n: (s, t) for n, s, t in GOODS_COLS}
    conds, params = [], []
    for name, text in (filters or {}).items():
        if name in by:
            sql, typ = by[name]
            c = grid._condition(name, typ, text, params, lambda n, ct, p, _s=sql: _s)
            if c:
                conds.append(c)
    if q and q.strip():
        conds.append("instr(fold(COALESCE(ship_key,'')||' '||COALESCE(ma_hang,'')||' '||COALESCE(ten_hang,'')||' '||COALESCE(xuat_xu,'')), ?) > 0")
        params.append(fold(q.strip()))
    return ("WHERE " + " AND ".join(conds)) if conds else "", params


def goods_query(conn, filters, q, sort, direction, page, size):
    where, wp = _goods_where(conn, filters, q)
    by = {n: s for n, s, _ in GOODS_COLS}
    order = "ORDER BY id"
    if sort in by:
        order = f"ORDER BY ({by[sort]}) IS NULL, ({by[sort]}) {'DESC' if direction == 'desc' else 'ASC'}, id"
    total = conn.execute(f"SELECT COUNT(*) FROM ob_goods {where}", wp).fetchone()[0]
    sel = ", ".join(f"{s} AS c{i}" for i, (_, s, _) in enumerate(GOODS_COLS))
    rows = conn.execute(f"SELECT id, ship_lc, {sel} FROM ob_goods {where} {order} LIMIT ? OFFSET ?", wp + [size, (page - 1) * size]).fetchall()
    # link sang List shipment: số dòng List shipment cùng Shipment#
    counts = {}
    info = table_info(conn, "list")
    if info and rows:
        for r in conn.execute("SELECT key, cells FROM ob_rows WHERE tbl='list'"):
            s = ship_of("list", json.loads(r["cells"]), r["key"])
            if s:
                counts[s.lower()] = counts.get(s.lower(), 0) + 1
    src = {x["key_lc"] for x in conn.execute("SELECT key_lc FROM ob_rows WHERE tbl='source' AND key_lc IS NOT NULL")}
    data = [[r[f"c{i}"] for i in range(len(GOODS_COLS))] for r in rows]
    summ = conn.execute(f"""SELECT COUNT(*) n, COUNT(DISTINCT ship_lc) ships FROM ob_goods {where}""", wp).fetchone()
    by_ccy = [{"ccy": x["ccy"], "amount": x["a"]} for x in conn.execute(
        f"SELECT ccy, ROUND(SUM(CASE WHEN qty IS NOT NULL AND don_gia IS NOT NULL THEN qty*don_gia ELSE COALESCE(amount,0) END),2) a "
        f"FROM ob_goods {where} GROUP BY ccy ORDER BY a DESC", wp)]
    return {"total": total, "columns": [n for n, _, _ in GOODS_COLS], "types": [t for _, _, t in GOODS_COLS], "rows": data,
            "ids": [r["id"] for r in rows], "list_lines": [counts.get(r["ship_lc"], 0) for r in rows], "in_source": [r["ship_lc"] in src for r in rows],
            "summary": {"lines": summ["n"], "ships": summ["ships"], "by_ccy": by_ccy}}


def goods_line_get(conn, gid):
    r = conn.execute("SELECT * FROM ob_goods WHERE id=?", (gid,)).fetchone()
    if r is None:
        raise KeyError(gid)
    d = dict(r)
    d["amount"] = _amount(d["qty"], d["don_gia"], d["amount"])
    return d


def goods_line_save(conn, gid, ship, fields):
    """Thêm (gid=None) hoặc sửa một dòng hàng."""
    ship = (ship or "").strip()
    if not ship:
        raise ValueError("Thiếu Shipment")
    d = (_clean_goods([fields]) or [None])[0]
    if d is None:
        raise ValueError("Cần ít nhất Item hoặc Description")
    cols = (d["ma_hang"], d["ten_hang"], d["qty"], d["dvt"], d["don_gia"], d["ccy"], d["xuat_xu"], d["amount"], d["ter"], d["net_w"], d["gross_w"])
    with conn:
        if gid:
            conn.execute("""UPDATE ob_goods SET ship_key=?, ship_lc=?, ma_hang=?, ten_hang=?, qty=?, dvt=?, don_gia=?, ccy=?, xuat_xu=?, amount=?,
                            ter=?, net_w=?, gross_w=? WHERE id=?""", (ship, ship.lower(), *cols, gid))
            return gid
        pos = (conn.execute("SELECT MAX(pos) FROM ob_goods WHERE ship_lc=?", (ship.lower(),)).fetchone()[0] or 0) + 1
        return conn.execute("""INSERT INTO ob_goods(ship_key,ship_lc,pos,ma_hang,ten_hang,qty,dvt,don_gia,ccy,xuat_xu,amount,ter,net_w,gross_w)
                               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (ship, ship.lower(), pos, *cols)).lastrowid


def goods_line_delete(conn, gid):
    with conn:
        conn.execute("DELETE FROM ob_goods WHERE id=?", (gid,))


def goods_export_xlsx(conn, filters, q, sort, direction):
    res = goods_query(conn, filters, q, sort, direction, 1, 500000)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Dong hang")
    ws.append(res["columns"] + ["Số dòng List shipment"])
    for r, n in zip(res["rows"], res["list_lines"]):
        ws.append(r + [n])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ----------------------------------------------------------------------------------------------- tìm kiếm shipment
def _hit_field(cols, cells, term):
    """Cột đầu tiên chứa từ khoá -> (tên cột, đoạn trích)."""
    for c in cols:
        v = cells.get(c["letter"])
        if v is not None and term in fold(str(v)):
            s = str(v).replace("\n", " ")
            i = max(0, fold(s).find(term) - 25)
            return c["name"], s[i:i + 90]
    return None, None


def search(conn, q, limit=60):
    """Tìm shipment theo bất kỳ dữ liệu nào: Source data, List shipment, Tracking, Information of goods (không phân biệt hoa thường / dấu)."""
    term = fold((q or "").strip())
    if len(term) < 2:
        return {"q": q, "results": [], "total": 0}
    conn.create_function("fold", 1, fold, deterministic=True)
    found = {}  # shipment id (chữ thường) -> kết quả

    def add(ship, tag, field, text, score):
        if not ship:
            return
        r = found.setdefault(str(ship).lower(), {"ship": str(ship), "row_id": None, "found": [], "score": 0})
        r["score"] = max(r["score"], score)
        if all(f["tbl"] != tag for f in r["found"]):
            r["found"].append({"tbl": tag, "field": field, "text": text})

    infos = {t: table_info(conn, t) for t in ("source", "list", "tracking")}
    for tbl, label in (("source", "Shipment"), ("list", "List shipment"), ("tracking", "Tracking")):
        info = infos[tbl]
        if not info:
            continue
        cols = info["cols"]
        ship_letter = {"source": info["key_col"], "tracking": info["key_col"]}.get(tbl)
        if tbl == "list":
            ship_letter = next((c["letter"] for c in cols if c["name"].lower().startswith("shipment#")), info["key_col"])
        for r in conn.execute("SELECT id, key, cells FROM ob_rows WHERE tbl=? AND instr(fold(cells), ?) > 0 ORDER BY pos DESC LIMIT 400", (tbl, term)):
            cells = json.loads(r["cells"])
            ship = cells.get(ship_letter) or (r["key"].rsplit("-", 1)[0] if r["key"] and tbl == "list" else r["key"])
            field, text = _hit_field(cols, cells, term)
            exact = fold(str(ship or "")) == term
            add(ship, label, field, text, 3 if exact else (2 if tbl == "source" else 1))
            if tbl == "source":
                found[str(ship).lower()]["row_id"] = r["id"]
    for r in conn.execute("SELECT ship_key, ma_hang, ten_hang FROM ob_goods WHERE instr(fold(COALESCE(ma_hang,'')||' '||COALESCE(ten_hang,'')), ?) > 0 LIMIT 300", (term,)):
        add(r["ship_key"], "Information of goods", "Mã hàng / Tên hàng", f"{r['ma_hang'] or ''} {r['ten_hang'] or ''}".strip()[:90], 1)
    for r in conn.execute("SELECT ship_key, invoice_no, so_to, customer_po FROM ob_invoices WHERE instr(fold(COALESCE(invoice_no,'')||' '||COALESCE(so_to,'')||' '||COALESCE(customer_po,'')||' '||COALESCE(polarium_ref,'')), ?) > 0 LIMIT 100", (term,)):
        add(r["ship_key"], "Invoice", "Invoice / SO-TO / PO", f"{r['invoice_no'] or ''} {r['so_to'] or ''}".strip()[:90], 2)
    # thông tin hiển thị lấy từ Source data
    info = infos["source"]
    letter = {c["name"]: c["letter"] for c in info["cols"]} if info else {}
    pick = {"consignee": "Consignee (Complete Name and Address)", "pod": "Port of Discharge", "fwd": "FWD", "status": "Status", "booking": "Booking No.", "bl": "B/L No.", "etd": "ETD"}
    out = sorted(found.values(), key=lambda r: (-r["score"], r["ship"]))[:limit]
    for r in out:
        row = conn.execute("SELECT id, cells FROM ob_rows WHERE tbl='source' AND key_lc=? ORDER BY pos LIMIT 1", (r["ship"].lower(),)).fetchone()
        if row:
            cells = json.loads(row["cells"])
            r["row_id"] = row["id"]
            for k, name in pick.items():
                v = cells.get(letter.get(name))
                r[k] = str(v).split("\n")[0][:60] if v is not None else None
    return {"q": q, "results": out, "total": len(found)}


# ----------------------------------------------------------------------------------------------- báo cáo tổng quan
CANCELLED = "Huỷ (Delete)"
MODE_LABELS = {"AIR": "Air", "SEA": "Sea", "TRUCK": "Truck", "OTHER": "Chưa phân loại"}


def _norm_status(v):
    s = str(v or "").strip()
    if not s:
        return None
    f = fold(s)
    if f.startswith(("delet", "delel", "cancel", "combine/delet", "huy")):
        return CANCELLED
    if f.startswith("done"):
        return "Done"  # gộp "DOne", "Done/Roll the BCF", "Done/Return VNOUT00073"…
    return s[:1].upper() + s[1:]


def _norm_mode(*values):
    """Air / Sea / Truck từ 'Type of Movement' (Source data): AIR, FCL/LCL = Sea, TRUCK."""
    for v in values:
        f = fold(str(v or ""))
        if not f or f.startswith("#") or f == "0":
            continue
        if "air" in f:
            return "AIR"
        if "truck" in f or "road" in f or "land" in f or "xe tai" in f:
            return "TRUCK"
        if "sea" in f or "fcl" in f or "lcl" in f or "ocean" in f:
            return "SEA"
    return "OTHER"


def _norm_fwd(v):
    s = " ".join(str(v or "").split())
    return s.upper() if s and s != "0" else None


def _base_id(i):
    """Shipment gốc: bỏ hậu tố -DG (bản khai hàng nguy hiểm) và -1, -2… (dòng container phụ). 'V-23-00008' giữ nguyên."""
    s = re.sub(r"-DG$", "", str(i).strip(), flags=re.I)
    return re.sub(r"-\d{1,2}$", "", s)


def _containers(*values):
    """Các số container / biển số xe trong ô 'Cont no.' (nhiều số ngăn bởi xuống dòng, dấu phẩy, ;); bỏ chữ chung như 'Air/LCL'."""
    out = set()
    for v in values:
        for tok in re.split(r"[\n,;/]+", str(v or "")):
            tok = tok.strip().upper()
            if len(tok) >= 6 and re.search(r"\d", tok) and re.search(r"[A-Z]", tok):
                out.add(tok)
    return out


def shipment_summary(conn):
    """Một dòng cho mỗi shipment, tính từ SOURCE DATA (gộp các dòng -DG và dòng container phụ -1, -2… về shipment gốc):
    tình trạng (Status), phương thức (Type of Movement), ngày pick up, forwarder (FWD), số lượng (Q'ty of CTNS), số container/xe (Cont no.)."""
    info = table_info(conn, "source")
    if not info:
        return []
    have = {c["letter"] for c in info["cols"]}
    first = lambda cells, *ls: next((cells.get(l) for l in ls if l in have and cells.get(l) not in (None, "", 0)), None)
    iso = re.compile(r"^\d{4}-\d{2}-\d{2}")
    by = {}
    for r in conn.execute("SELECT key, cells FROM ob_rows WHERE tbl='source' AND key IS NOT NULL ORDER BY pos"):
        by.setdefault(_base_id(r["key"]), []).append((r["key"], json.loads(r["cells"])))
    out = []
    for ship, rows in by.items():
        rows.sort(key=lambda kc: (kc[0].upper() != ship.upper(), len(kc[0])))  # dòng gốc trước
        cells_list = [c for _, c in rows]
        sts = {_norm_status(first(c, "AT", "AF")) for c in cells_list} - {None}
        live = sts - {CANCELLED}
        if not sts:
            status = "Chưa có status"
        elif not live:
            status = CANCELLED
        elif live == {"Done"}:
            status = "Done"
        elif len(live) == 1:
            status = next(iter(live))
        else:
            status = "Processing" if "Done" in live else sorted(live)[0]
        mode = _norm_mode(*[first(c, "Q") for c in cells_list])
        dates = [str(d)[:10] for d in (first(c, "+ActualPU", "O", "AS", "P") for c in cells_list) if d and iso.match(str(d))]
        fwd = next((x for x in (_norm_fwd(first(c, "AU")) for c in cells_list) if x), None)
        qty = 0.0
        for c in cells_list:
            v = c.get("S")
            if isinstance(v, (int, float)):
                qty += v
        conts = _containers(*[c.get(l) for c in cells_list for l in ("Z", "AV") if l in have])
        cust = next((str(first(c, "G")).split("\n")[0] for c in cells_list if first(c, "G")), None)
        out.append({"ship": ship, "status": status, "mode": mode, "date": min(dates) if dates else None, "forwarder": fwd,
                    "qty": qty, "items": len(rows), "customer": cust, "containers": len(conts)})
    return out

def report(conn, statuses=None, date_from=None, date_to=None):
    """Báo cáo tổng quan (dữ liệu từ SOURCE DATA): số shipment theo phương thức (Air / Sea / Truck), theo tháng, theo forwarder, tiến độ đã vận chuyển.

    statuses: danh sách tình trạng được chọn (mặc định: mọi tình trạng trừ 'Huỷ').
    date_from / date_to: lọc theo ngày pick up thực tế (hoặc ETD nếu chưa có) — shipment chưa có ngày chỉ nằm trong báo cáo khi không lọc ngày.
    """
    ships = shipment_summary(conn)
    if date_from or date_to:
        ships = [s for s in ships if s["date"] and (not date_from or s["date"] >= date_from) and (not date_to or s["date"] <= date_to)]
    counts = {}
    for s in ships:
        counts[s["status"]] = counts.get(s["status"], 0) + 1
    all_statuses = [{"status": k, "count": v} for k, v in sorted(counts.items(), key=lambda kv: -kv[1])]
    sel = set(statuses) if statuses else {k for k in counts if k != CANCELLED}
    chosen = [s for s in ships if s["status"] in sel]
    modes = {k: 0 for k in MODE_LABELS}
    months, fwd = {}, {}
    for s in chosen:
        modes[s["mode"]] += 1
        mk = (s["date"] or "")[:7] or "Chưa có ngày"
        months.setdefault(mk, {k: 0 for k in MODE_LABELS})[s["mode"]] += 1
        if s["forwarder"]:
            fwd[s["forwarder"]] = fwd.get(s["forwarder"], 0) + 1
    active = [s for s in ships if s["status"] != CANCELLED]
    done = sum(1 for s in active if s["status"] == "Done")
    trips = sum(s["containers"] for s in active)  # số container / xe (Cont no. trong Source data) của các shipment chưa huỷ
    ships_with_trips = sum(1 for s in active if s["containers"])
    return {
        "statuses": all_statuses, "selected": sorted(sel), "total": len(chosen), "qty": round(sum(s["qty"] for s in chosen), 2),
        "modes": [{"mode": k, "label": MODE_LABELS[k], "count": v} for k, v in modes.items()],
        "months": [{"month": k, **v} for k, v in sorted(months.items())],
        "forwarders": [{"v": k, "n": v} for k, v in sorted(fwd.items(), key=lambda kv: -kv[1])[:12]],
        "progress": {"done": done, "active": len(active), "pct": round(100 * done / len(active), 1) if active else 0,
                     "cancelled": len(ships) - len(active)},
        "containers": {"count": trips, "shipments": ships_with_trips},
        "shipments": chosen,
    }


def report_xlsx(conn, statuses=None, date_from=None, date_to=None):
    r = report(conn, statuses, date_from, date_to)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tổng hợp"
    ws.append(["Báo cáo Outbound shipment"])
    ws.append(["Tình trạng được chọn", ", ".join(r["selected"])])
    ws.append(["Khoảng ngày pick up", f"{date_from or '…'} → {date_to or '…'}"])
    ws.append(["Tổng shipment", r["total"]])
    ws.append(["Tiến độ: đã vận chuyển (Done) / tổng (không tính Huỷ)", f"{r['progress']['done']} / {r['progress']['active']} ({r['progress']['pct']}%)"])
    ws.append([])
    ws.append(["Phương thức", "Số shipment"])
    for m in r["modes"]:
        ws.append([m["label"], m["count"]])
    ws.append([])
    ws.append(["Tình trạng", "Số shipment (toàn bộ, theo khoảng ngày)"])
    for s in r["statuses"]:
        ws.append([s["status"], s["count"]])
    ws2 = wb.create_sheet("Theo tháng")
    ws2.append(["Tháng", "Air", "Sea", "Truck", "Chưa phân loại", "Tổng"])
    for m in r["months"]:
        ws2.append([m["month"], m["AIR"], m["SEA"], m["TRUCK"], m["OTHER"], m["AIR"] + m["SEA"] + m["TRUCK"] + m["OTHER"]])
    ws3 = wb.create_sheet("Danh sách shipment")
    ws3.append(["Shipment", "Tình trạng", "Phương thức", "Ngày pick up", "Forwarder", "Khách hàng", "Số dòng item", "Số lượng"])
    for s in sorted(r["shipments"], key=lambda x: (x["date"] or "", x["ship"])):
        ws3.append([s["ship"], s["status"], MODE_LABELS[s["mode"]], s["date"], s["forwarder"], s["customer"], s["items"], s["qty"]])
    for w_, widths in ((ws, (52, 40)), (ws2, (14, 8, 8, 8, 16, 8)), (ws3, (16, 16, 14, 14, 24, 40, 12, 12))):
        for i, wd in enumerate(widths, start=1):
            w_.column_dimensions[get_column_letter(i)].width = wd
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------------------------- Invoice quét từ file
INV_FIELDS = ["invoice_no", "invoice_date", "customer_po", "so_to", "polarium_ref", "incoterms", "mode", "carrier", "shipped_on", "hs_code",
              "term_of_payment", "remarks", "sold_by", "sold_to", "shipper", "deliver_to"]


def invoices_of(conn, ship):
    return [dict(r) for r in conn.execute("SELECT * FROM ob_invoices WHERE ship_lc=? ORDER BY id", (str(ship).lower(),))]


def invoices_list(conn):
    return [dict(r) for r in conn.execute("SELECT * FROM ob_invoices ORDER BY id DESC")]


def invoice_delete(conn, iid, delete_lines=False):
    r = conn.execute("SELECT ship_lc, invoice_no FROM ob_invoices WHERE id=?", (iid,)).fetchone()
    with conn:
        if r and delete_lines and r["invoice_no"]:
            conn.execute("DELETE FROM ob_goods WHERE ship_lc=? AND inv_no=?", (r["ship_lc"], r["invoice_no"]))
        conn.execute("DELETE FROM ob_invoices WHERE id=?", (iid,))


def find_invoice(conn, key):
    """Tìm Invoice theo số SO/TO, số Invoice hoặc Shipment (không phân biệt hoa thường) — dùng cho biểu mẫu INV / PK."""
    k = str(key or "").strip().lower()
    if not k:
        return None
    for col in ("so_to_lc", "invoice_no_lc", "ship_lc"):
        r = conn.execute(f"SELECT * FROM ob_invoices WHERE {col}=? ORDER BY id DESC LIMIT 1", (k,)).fetchone()
        if r:
            return dict(r)
    r = conn.execute("SELECT * FROM ob_invoices WHERE instr(',' || REPLACE(LOWER(COALESCE(so_to,'')), ' ', '') || ',', ?) > 0 ORDER BY id DESC LIMIT 1",
                     (f",{k},",)).fetchone()  # nhiều số SO ngăn bởi dấu phẩy
    return dict(r) if r else None


def apply_invoice(conn, ship, header, lines, ccy=None, total=None, mode="replace", file_path=None, file_name=None, source=None):
    """Lưu Invoice đã quét: thông tin Invoice, các dòng hàng (gắn theo Shipment + số Invoice), Số Invoice ở Source data, file gốc đính kèm."""
    ship = (ship or "").strip()
    if not ship:
        raise ValueError("Thiếu Shipment ID")
    inv_no = (header.get("invoice_no") or "").strip() or None
    good = [l for l in lines if (l.get("ma_hang") or l.get("ten_hang"))]
    if not good:
        raise ValueError("Invoice không có dòng hàng nào để lưu")
    inc = (header.get("incoterms") or "").strip() or None
    new = []
    for l in good:
        new.append({"ma_hang": l.get("ma_hang"), "ten_hang": l.get("ten_hang"), "qty": l.get("qty"), "dvt": l.get("dvt"), "xuat_xu": l.get("xuat_xu"),
                    "don_gia": l.get("don_gia"), "amount": l.get("amount"), "ccy": (ccy or "USD"), "ter": inc, "inv_no": inv_no})
    old = goods_get(conn, ship)
    keep = [g for g in old if not (mode == "replace" and inv_no and g.get("inv_no") == inv_no)]
    if mode == "replace" and not inv_no:
        keep = []
    goods_save(conn, ship, keep + new)
    vals = [(header.get(k) or None) for k in INV_FIELDS]
    cols = ",".join(INV_FIELDS)
    with conn:
        ex = conn.execute("SELECT id FROM ob_invoices WHERE ship_lc=? AND COALESCE(invoice_no_lc,'')=?", (ship.lower(), (inv_no or "").lower())).fetchone()
        so = header.get("so_to") or ""
        extra = (ccy, total, len(new), file_name, source, so.lower() or None, (inv_no or "").lower() or None)
        if ex:
            conn.execute(f"UPDATE ob_invoices SET {','.join(k + '=?' for k in INV_FIELDS)}, ccy=?, total_value=?, n_lines=?, file_name=?, source=?, so_to_lc=?, invoice_no_lc=?, "
                         f"at=datetime('now','localtime') WHERE id=?", (*vals, *extra, ex["id"]))
            iid = ex["id"]
        else:
            iid = conn.execute(f"INSERT INTO ob_invoices(ship_key,ship_lc,{cols},ccy,total_value,n_lines,file_name,source,so_to_lc,invoice_no_lc) "
                               f"VALUES(?,?,{','.join('?' * len(INV_FIELDS))},?,?,?,?,?,?,?)", (ship, ship.lower(), *vals, *extra)).lastrowid
    # Số Invoice ở Source data (cột bổ sung — không ảnh hưởng biểu mẫu) + đính kèm file gốc vào dòng Source
    src = conn.execute("SELECT id FROM ob_rows WHERE tbl='source' AND key_lc=? ORDER BY pos LIMIT 1", (ship.lower(),)).fetchone()
    attached = False
    if src and inv_no:
        nos = [x["invoice_no"] for x in conn.execute("SELECT invoice_no FROM ob_invoices WHERE ship_lc=? AND invoice_no IS NOT NULL ORDER BY id", (ship.lower(),))]
        update_row(conn, src["id"], {"+Invoice": ", ".join(dict.fromkeys(nos))})
        if file_path and file_name:
            already = any(f["name"] == file_name for f in files_of(conn, src["id"]))
            if not already:
                add_file(conn, src["id"], file_name, Path(file_path).read_bytes())
                attached = True
    return {"id": iid, "ship": ship, "invoice_no": inv_no, "lines": len(new), "in_source": bool(src), "attached": attached}


# ----------------------------------------------------------------------------------------------- xuất file đầy đủ
def full_export_xlsx(conn, filters=None, q="", by="source", ship=None):
    """Một file Excel: mọi shipment và chi tiết đầy đủ — mọi trường Source data, các dòng hàng (mã hàng, số lượng, đơn giá…), Invoice, List shipment, Tracking.

    Sheet 'Chi tiết shipment': mỗi DÒNG HÀNG một dòng, lặp lại toàn bộ thông tin shipment (shipment chưa có dòng hàng vẫn có 1 dòng) — lọc theo Shipment ID là thấy đủ.
    filters / q: chỉ xuất các shipment khớp bộ lọc (mặc định tất cả). by="source": lọc theo bảng Shipment; by="list": lọc theo bảng List shipment
    (shipment = Shipment# của các dòng khớp, gồm cả bản -DG / -1 trong Source data). Shipment chưa có dòng hàng Invoice thì lấy Item / Description / Qty từ List shipment."""
    from openpyxl.styles import Alignment, Font, PatternFill
    info = table_info(conn, "source")
    if not info:
        raise ValueError("Chưa có dữ liệu Shipment — hãy nạp file Excel trước")
    cols = info["cols"]
    linfo = table_info(conn, "list")
    list_by_ship = {}
    if linfo:
        for r in conn.execute("SELECT key, cells FROM ob_rows WHERE tbl='list' ORDER BY pos"):
            c = json.loads(r["cells"])
            s = ship_of("list", c, r["key"])
            if s:
                list_by_ship.setdefault(s.lower(), []).append(c)
    scoped = bool(ship) or bool(filters) or bool(q)
    if ship or (by == "list" and (filters or q) and linfo):
        want = {}
        if ship:  # đúng 1 shipment (mở từ chi tiết List shipment / Source data) — không phụ thuộc bộ lọc
            want[str(ship).strip().lower()] = str(ship).strip()
        else:
            lwhere, lwp = _where(conn, "list", filters, q, linfo["cols"])
            for r in conn.execute(f"SELECT t.key, t.cells FROM ob_rows t {lwhere} ORDER BY t.pos", lwp):
                s = ship_of("list", json.loads(r["cells"]), r["key"])
                if s:
                    want.setdefault(s.lower(), s)
        allsrc = [(r["id"], r["key"], json.loads(r["cells"])) for r in conn.execute("SELECT id, key, cells FROM ob_rows WHERE tbl='source' AND key IS NOT NULL ORDER BY pos")]
        src_rows, have = [], set()
        for rid, key, cells in allsrc:
            base = re.sub(r"-DG$", "", key, flags=re.I)
            base = re.sub(r"-\d{1,2}$", "", base)
            if key.lower() in want or base.lower() in want:
                src_rows.append((rid, key, cells))
                have.add(base.lower() if base.lower() in want else key.lower())
        for lc, s in want.items():  # shipment chỉ có trong List shipment (chưa có ở Source data)
            if lc not in have:
                src_rows.append((None, s, {}))
        ships = set(want) | {k.lower() for _, k, _ in src_rows if k}
    else:
        where, wp = _where(conn, "source", filters, q, cols)
        src_rows = [(r["id"], r["key"], json.loads(r["cells"])) for r in conn.execute(f"SELECT t.id, t.key, t.cells FROM ob_rows t {where} ORDER BY t.pos", wp)]
        ships = {k.lower() for _, k, _ in src_rows if k}
    goods = {}
    for r in conn.execute("SELECT * FROM ob_goods ORDER BY ship_lc, pos"):
        d = dict(r)
        d["amount"] = _amount(d["qty"], d["don_gia"], d["amount"])
        goods.setdefault(d["ship_lc"], []).append(d)
    invs = {}
    for r in conn.execute("SELECT * FROM ob_invoices ORDER BY id"):
        invs.setdefault(r["ship_lc"], []).append(dict(r))
    nfiles = {r["k"]: r["n"] for r in conn.execute(
        "SELECT o.key_lc k, COUNT(*) n FROM ob_files f JOIN ob_rows o ON o.id=f.row_id WHERE o.tbl='source' GROUP BY o.key_lc")}

    def val(v):
        return dec(v) if isinstance(v, str) else v

    head_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(bold=True, color="FFFFFF")

    def sheet(wb, title, headers, rows, widths=None, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title[:31]
        ws.append(headers)
        for c in ws[1]:
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for r in rows:
            ws.append(r)
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions
        for i, h in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(i, max(11, min(34, len(str(h)) + 3)))
        return ws

    line_heads = ["Dòng #", "Nguồn dòng hàng", "Item (mã hàng)", "Description", "Qty", "Unit", "COO", "Unit price", "Currency", "Amount", "Ter of delivery", "Net weight", "Gross weight", "Số Invoice (dòng)"]
    inv_heads = ["Invoice Number", "Invoice Date", "Polarium SO/TO", "Customer PO", "Incoterms (Invoice)", "Mode (Invoice)", "Invoice total", "Invoice currency"]
    src_heads = [c["name"] for c in cols]
    line_of = lambda g, i: [i, g.get("_src", "Invoice"), g["ma_hang"], g["ten_hang"], g["qty"], g["dvt"], g["xuat_xu"], g["don_gia"], g["ccy"], g["amount"], g["ter"], g["net_w"], g["gross_w"], g["inv_no"]]
    inv_of = lambda ship, g: next((x for x in invs.get(ship, []) if g and g.get("inv_no") and x["invoice_no"] == g["inv_no"]), (invs.get(ship) or [None])[0])
    inv_cells = lambda x: ([x["invoice_no"], val(x["invoice_date"]), x["so_to"], x["customer_po"], x["incoterms"], x["mode"], x["total_value"], x["ccy"]] if x else [None] * 8)

    wb = Workbook()
    # 1. Chi tiết shipment: 1 dòng hàng = 1 dòng, đủ thông tin shipment
    detail = []
    for rid, key, cells in src_rows:
        sl = (key or "").lower()
        shipment_part = [val(cells.get(c["letter"])) for c in cols]
        base = re.sub(r"-\d{1,2}$", "", re.sub(r"-DG$", "", sl, flags=re.I))
        lines = goods.get(sl) or goods.get(base)
        if not lines:  # chưa có dòng hàng Invoice: dùng Item / Description / Qty của List shipment
            li = [{"_src": "List shipment", "ma_hang": c.get("H"), "ten_hang": c.get("I"), "qty": c.get("E"), "dvt": None, "xuat_xu": None, "don_gia": None,
                   "ccy": None, "amount": None, "ter": None, "net_w": None, "gross_w": None, "inv_no": None, "ship_lc": sl} for c in (list_by_ship.get(sl) or list_by_ship.get(base) or [])]
            lines = li or [None]
        for i, g in enumerate(lines, start=1):
            detail.append([key] + (line_of(g, i) if g else [None] * len(line_heads)) + inv_cells(inv_of(sl if sl in invs else base, g)) + shipment_part)
    sheet(wb, "Chi tiết shipment", ["Shipment ID"] + line_heads + inv_heads + src_heads, detail, {3: 30}, first=True)
    # 2. Shipment: 1 dòng / shipment
    summ = []
    for rid, key, cells in src_rows:
        sl = (key or "").lower()
        gl = goods.get(sl, [])
        tot = {}
        for g in gl:
            if g["amount"] is not None:
                tot[g["ccy"] or "USD"] = tot.get(g["ccy"] or "USD", 0) + g["amount"]
        ivs = invs.get(sl, [])
        summ.append([key, len(gl), "; ".join(f"{v:,.2f} {c}" for c, v in tot.items()) or None, "; ".join(x["invoice_no"] or "" for x in ivs) or None,
                     "; ".join(x["so_to"] or "" for x in ivs if x["so_to"]) or None, nfiles.get(sl)] + [val(cells.get(c["letter"])) for c in cols])
    sheet(wb, "Shipment", ["Shipment ID", "Số dòng hàng", "Tổng Amount", "Invoice", "SO/TO", "Số file đính kèm"] + src_heads, summ)
    # 3. Dòng hàng
    gl_rows = []
    for rid, key, cells in src_rows:
        for i, g in enumerate(goods.get((key or "").lower(), []), start=1):
            gl_rows.append([key] + line_of(g, i))
    sheet(wb, "Dòng hàng", ["Shipment ID"] + line_heads, gl_rows, {3: 40})
    # 4. Invoice
    inv_rows = [[x["ship_key"], x["invoice_no"], val(x["invoice_date"]), x["so_to"], x["customer_po"], x["polarium_ref"], x["incoterms"], x["mode"], x["carrier"], x["shipped_on"],
                 x["hs_code"], x["total_value"], x["ccy"], x["n_lines"], x["sold_to"], x["deliver_to"], x["shipper"], x["remarks"], x["file_name"]]
                for k, v in invs.items() if k in ships for x in v]
    sheet(wb, "Invoice", ["Shipment ID", "Invoice Number", "Invoice Date", "SO/TO", "Customer PO", "Polarium Ref", "Incoterms", "Mode", "Carrier", "Shipped on", "HS code",
                          "Total value", "Currency", "Số dòng", "Sold to", "Deliver to", "Shipper", "Remarks", "File"], inv_rows)
    # 5-6. List shipment và Tracking của các shipment đã chọn
    for tbl, title in (("list", "List shipment"), ("tracking", "Tracking")):
        tinfo = table_info(conn, tbl)
        if not tinfo:
            continue
        rows = []
        for r in conn.execute("SELECT key, cells FROM ob_rows WHERE tbl=? ORDER BY pos", (tbl,)):
            c = json.loads(r["cells"])
            ship = ship_of(tbl, c, r["key"]) if tbl == "list" else r["key"]
            if not scoped or (ship and ship.lower() in ships) or (by == "list" and tbl == "tracking" and ship and re.sub(r"-(DG|\d{1,2})$", "", ship, flags=re.I).lower() in ships):
                rows.append([val(c.get(x["letter"])) for x in tinfo["cols"]])
        sheet(wb, title, [x["name"] for x in tinfo["cols"]], rows)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
