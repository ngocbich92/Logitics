"""BOM = định mức NVL cho 1 đơn vị thành phẩm (dùng để tính NVL của lô C/O)."""
import io
import re

import openpyxl
from openpyxl.styles import Font

import engine
from db import fold, ma_key
from importer import num, txt

ALIASES = {
    "ma_tp": ["ma thanh pham", "ma tp", "ma san xuat hang", "thanh pham"],
    "ten_hang": ["ten hang hoa", "ten thanh pham", "ten hang"],
    "hs6": ["ma hs", "ma hs 6 so", "ma hs hang hoa", "hs"],
    "ma_hang": ["ma hang", "ma hang nvl", "ma nvl", "ma hang (nvl)"],
    "mo_ta": ["item description", "mo ta", "ten nvl", "description"],
    "dinh_muc": ["qty used in fg", "dinh muc", "qty", "so luong", "dinh muc/sp"],
    "dvt": ["unit", "dvt", "don vi tinh", "don vi"],
}


def _hdr(row):
    cols = {}
    for i, h in enumerate(row):
        f = fold(h).strip() if h else ""
        for field, names in ALIASES.items():
            if f in names and field not in cols:
                cols[field] = i
    return cols


def parse_table(path):
    """File nhiều BOM: Mã thành phẩm | (Tên hàng hóa) | (Mã HS) | Mã hàng NVL | (Item description) | Qty used in FG | (Unit)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    hi = next((i for i, r in enumerate(rows[:15]) if {"ma_tp", "ma_hang", "dinh_muc"} <= set(_hdr(r))), None)
    if hi is None:
        raise ValueError("Không tìm thấy dòng tiêu đề. Cần các cột: Mã thành phẩm, Mã hàng (NVL), Qty used in FG (định mức).")
    c = _hdr(rows[hi])
    out = []
    for n, r in enumerate(rows[hi + 1:], start=hi + 2):
        g = lambda f: r[c[f]] if f in c and c[f] < len(r) else None
        if all(v in (None, "") for v in r):
            continue
        out.append({"row": n, "ma_tp": txt(g("ma_tp")) or "", "ten_hang": txt(g("ten_hang")), "hs6": txt(g("hs6")),
                    "ma_hang": txt(g("ma_hang")) or "", "mo_ta": txt(g("mo_ta")), "dinh_muc": num(g("dinh_muc")), "dvt": txt(g("dvt"))})
    if not out:
        raise ValueError("File không có dòng dữ liệu")
    return out


def parse_any(path):
    """File Excel bất kỳ: nếu có sheet tên 'BOM' (hoặc 'BOM (2)', ...) theo mẫu sổ cũ thì đọc tự động từng sheet = một thành phẩm;
    nếu không thì đọc sheet đầu tiên theo dạng bảng nhiều thành phẩm (xem file mẫu).
    Trả về (rows, sheets_used)."""
    import importer
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        names = importer.sheets_named(wb, r"bom(\s*\(\d+\))?")
        if not names:
            wb.close()
            return parse_table(path), []
        rows = []
        for sn in names:
            header, lines = importer.read_bom_sheet(wb[sn])
            if not lines:
                continue
            for i, l in enumerate(lines, start=17):
                rows.append({"row": f"{sn}!{i}", "ma_tp": header["ma_tp"] or "", "ten_hang": header["ten_hang"], "hs6": header["hs6"],
                             "dvt_tp": header["dvt"], "ma_hang": l["ma_hang"], "mo_ta": l["mo_ta"], "dinh_muc": l["qty"], "dvt": l["dvt"]})
        if not rows:
            raise ValueError(f"Sheet {', '.join(names)} không có dòng định mức NVL nào")
        return rows, names
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass


def apply_table(conn, rows, dry_run=True):
    """Upsert các BOM trong file (thay toàn bộ dòng NVL của BOM đã có). Trả về tóm tắt + vấn đề."""
    problems, groups = [], {}
    for r in rows:
        if not r["ma_tp"]:
            problems.append({"row": r["row"], "msg": "Thiếu Mã thành phẩm (ô B4 của sheet BOM)"})
        elif not r["ma_hang"]:
            problems.append({"row": r["row"], "msg": "Thiếu Mã hàng NVL"})
        elif r["dinh_muc"] is None or r["dinh_muc"] <= 0:
            problems.append({"row": r["row"], "ma_tp": r["ma_tp"], "ma_hang": r["ma_hang"], "msg": "Định mức không hợp lệ"})
        else:
            groups.setdefault(r["ma_tp"], []).append(r)
    known = {r[0] for r in conn.execute("SELECT DISTINCT ma_key FROM import_lines")}
    result = []
    try:
        for tp, lines in groups.items():
            ex = conn.execute("SELECT id FROM boms WHERE ma_tp=?", (tp,)).fetchone()
            first = lambda f: next((x[f] for x in lines if x.get(f)), None)
            if ex:
                bid = ex["id"]
                conn.execute("UPDATE boms SET ten_hang=COALESCE(?,ten_hang), hs6=COALESCE(?,hs6), dvt=COALESCE(?,dvt), updated_at=datetime('now','localtime') WHERE id=?",
                             (first("ten_hang"), first("hs6"), first("dvt_tp"), bid))
                conn.execute("DELETE FROM bom_lines WHERE bom_id=?", (bid,))
            else:
                bid = conn.execute("INSERT INTO boms(ma_tp,ten_hang,hs6,dvt) VALUES(?,?,?,?)", (tp, first("ten_hang"), first("hs6"), first("dvt_tp"))).lastrowid
            unknown = 0
            for x in lines:
                conn.execute("INSERT INTO bom_lines(bom_id,ma_hang,ma_key,mo_ta,dvt,dinh_muc) VALUES(?,?,?,?,?,?)",
                             (bid, x["ma_hang"], ma_key(x["ma_hang"]), x["mo_ta"], x["dvt"], x["dinh_muc"]))
                if ma_key(x["ma_hang"]) not in known:
                    unknown += 1
                    problems.append({"row": x["row"], "ma_tp": tp, "ma_hang": x["ma_hang"], "level": "warn",
                                     "msg": "Mã NVL này chưa có trong dữ liệu tờ khai nhập"})
            result.append({"ma_tp": tp, "lines": len(lines), "updated": bool(ex), "unknown_codes": unknown})
        if dry_run:
            conn.rollback()
        else:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    errs = [p for p in problems if p.get("level") != "warn"]
    return {"boms": result, "problems": problems, "n_errors": len(errs), "n_warnings": len(problems) - len(errs)}


def save_bom(conn, bom_id, data):
    """Lưu thông tin + thay toàn bộ dòng NVL của BOM."""
    lines = [l for l in data.get("lines", []) if (l.get("ma_hang") or "").strip()]
    for l in lines:
        if num(l.get("dinh_muc")) is None or num(l.get("dinh_muc")) <= 0:
            raise ValueError(f"Định mức không hợp lệ ở mã {l.get('ma_hang')}")
    with conn:
        conn.execute("UPDATE boms SET ten_hang=?, hs6=?, dvt=?, note=?, updated_at=datetime('now','localtime') WHERE id=?",
                     (data.get("ten_hang"), data.get("hs6"), data.get("dvt"), data.get("note"), bom_id))
        conn.execute("DELETE FROM bom_lines WHERE bom_id=?", (bom_id,))
        for l in lines:
            conn.execute("INSERT INTO bom_lines(bom_id,ma_hang,ma_key,mo_ta,dvt,dinh_muc) VALUES(?,?,?,?,?,?)",
                         (bom_id, l["ma_hang"].strip(), ma_key(l["ma_hang"]), l.get("mo_ta"), l.get("dvt"), num(l["dinh_muc"])))


def create_bom(conn, ma_tp, data=None):
    ma_tp = (ma_tp or "").strip()
    if not ma_tp:
        raise ValueError("Cần Mã thành phẩm")
    if conn.execute("SELECT 1 FROM boms WHERE ma_tp=?", (ma_tp,)).fetchone():
        raise ValueError("BOM của mã thành phẩm này đã tồn tại")
    with conn:
        bid = conn.execute("INSERT INTO boms(ma_tp) VALUES(?)", (ma_tp,)).lastrowid
    if data:
        save_bom(conn, bid, data)
    return bid


def detail(conn, bom_id, settings):
    b = conn.execute("SELECT * FROM boms WHERE id=?", (bom_id,)).fetchone()
    if b is None:
        raise KeyError(bom_id)
    types = engine.eligible_types(settings, "CO")
    cleared = settings.get("only_cleared") == "1"
    lines, max_lots = [], None
    for l in conn.execute("SELECT * FROM bom_lines WHERE bom_id=? ORDER BY id", (bom_id,)):
        stock = sum(s["remain"] for s in engine.stock_rows(conn, l["ma_key"], types, "LIFO", cleared))
        lots = stock / l["dinh_muc"] if l["dinh_muc"] else None
        max_lots = lots if max_lots is None else min(max_lots, lots)
        lines.append({**dict(l), "ton": round(stock, 6), "lo_toi_da": int(lots) if lots is not None else None})
    return {"bom": dict(b), "lines": lines, "max_lots": int(max_lots) if max_lots is not None else None}


def lines_for_event(conn, bom_id):
    return [{"ma_hang": r["ma_hang"], "mo_ta": r["mo_ta"], "qty": r["dinh_muc"], "dvt": r["dvt"]}
            for r in conn.execute("SELECT * FROM bom_lines WHERE bom_id=? ORDER BY id", (bom_id,))]


def template_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.append(["Mã thành phẩm", "Tên hàng hóa", "Mã HS 6 số", "Mã hàng (NVL)", "Item description", "Qty used in FG", "Unit"])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.append(["800-00332", "Pin lithium 800-00332", "850760", "110-00063", "Cell 18650", 15, "EACH"])
    ws.append(["800-00332", "", "", "410-00002", "Keo Loctite", 0.44, "ML"])
    ws.append(["800-00400", "Pin lithium 800-00400", "850760", "110-00063", "Cell 18650", 20, "EACH"])
    for col, w in zip("ABCDEFG", [16, 28, 12, 16, 36, 14, 8]):
        ws.column_dimensions[col].width = w
    ws2 = wb.create_sheet("Huong dan")
    for t in ["Mỗi dòng = một NVL trong BOM của một thành phẩm. Nhiều thành phẩm trong cùng một file.",
              "Cột bắt buộc: Mã thành phẩm, Mã hàng (NVL), Qty used in FG (định mức cho 1 đơn vị thành phẩm).",
              "Tên hàng hóa, Mã HS 6 số: tuỳ chọn (lấy ở dòng đầu của mỗi thành phẩm). Item description, Unit: tuỳ chọn.",
              "BOM đã có (cùng Mã thành phẩm) sẽ bị thay toàn bộ dòng NVL bằng nội dung file. Xoá 3 dòng ví dụ trước khi nạp."]:
        ws2.append([t])
    ws2.column_dimensions["A"].width = 120
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
