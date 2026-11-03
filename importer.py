"""Nhập dữ liệu từ Excel: báo cáo tờ khai nhập (01_DATABASE), BOM, danh sách NVL xuất."""
import datetime as dt
import json
import re
from collections import Counter

import openpyxl

from db import ma_key

DB_HEADERS = {  # tên cột trong báo cáo hải quan -> field
    "Số TK": "so_tk", "Ngày ĐK": "ngay_dk", "Mã loại hình": "loai_hinh", "Tên đối tác": "doi_tac",
    "Vận đơn": "van_don", "Số hóa đơn TM": "so_hd", "Ngày HĐTM": "ngay_hd", "Ng.Tệ hóa đơn": "ccy",
    "Tỷ giá VNĐ": "ty_gia", "Trạng thái": "trang_thai", "Mã HS": "ma_hs", "Mã hàng": "ma_hang",
    "Tên hàng": "ten_hang", "Đơn vị tính": "dvt", "Số lượng": "qty", "Đơn giá hóa đơn": "don_gia_hd",
    "Đơn giá tính thuế": "don_gia_tt", "Trị giá hóa đơn": "tri_gia_hd", "Nước xuất xứ": "nuoc_xx",
}
# Báo cáo hải quan dùng "Mã loại hình" ở cột E nhưng sổ cũ có thể ghi "Mã loại hình" ở nhiều nơi — lấy lần đầu.


def to_iso(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return s[:10]


def num(v, default=None):
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def txt(v):
    return None if v is None else str(v).strip()


def _find_header(ws, must, max_rows=10):
    for r_i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, values_only=True), start=1):
        vals = [txt(c) for c in row]
        if must in vals:
            return r_i, vals
    raise ValueError(f"Không tìm thấy dòng tiêu đề chứa '{must}' trong sheet {ws.title}")


def _json_val(v):
    if v is None or v == "":
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        return to_iso(v)
    if isinstance(v, str):
        return v.strip() or None
    return v


def _col_type(values):
    """number / date / text — dùng cho bộ lọc. Số nguyên >= 1e7 (số TK, mã HS...) coi là text."""
    vals = [v for v in values if v is not None]
    if not vals:
        return "text"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and abs(v) >= 10**7) for v in vals):
        return "number"
    if all(isinstance(v, (dt.datetime, dt.date)) for v in vals):
        return "date"
    return "text"


def parse_database_sheet(ws):
    """Đọc sheet 01_DATABASE -> dict(parsed, names, types, n_rows). Chưa ghi gì vào DB."""
    all_rows = list(ws.iter_rows(values_only=True))
    hdr_idx = next((i for i, r in enumerate(all_rows[:10]) if "Số TK" in [txt(c) for c in r]), None)
    if hdr_idx is None:
        raise ValueError("Không tìm thấy dòng tiêu đề chứa 'Số TK'")
    hdr = [txt(c) for c in all_rows[hdr_idx]]
    names, seen_names = {}, Counter()
    for i, h in enumerate(hdr):
        if h:
            seen_names[h] += 1
            names[i] = h if seen_names[h] == 1 else f"{h} ({seen_names[h]})"
    col = {}
    for i, h in enumerate(hdr):
        if h in DB_HEADERS and DB_HEADERS[h] not in col:
            col[DB_HEADERS[h]] = i
    need = {"so_tk", "ma_hang", "qty", "loai_hinh"}
    if not need <= set(col):
        raise ValueError(f"Thiếu cột bắt buộc: {sorted(need - set(col))}")

    parsed, types_src = [], {n: [] for n in names.values()}
    counter = Counter()
    for row in all_rows[hdr_idx + 1:]:
        g = lambda f: row[col[f]] if f in col and col[f] < len(row) else None
        ma, qty = txt(g("ma_hang")), num(g("qty"))
        if not ma or qty is None:
            continue
        raw = {}
        for i, n in names.items():
            v = row[i] if i < len(row) else None
            types_src[n].append(v if not isinstance(v, str) else (v.strip() or None))
            jv = _json_val(v)
            if jv is not None:
                raw[n] = jv
        so_tk = str(int(g("so_tk"))) if isinstance(g("so_tk"), (int, float)) else txt(g("so_tk"))
        key = f"{so_tk}|{ma}|{qty}|{num(g('don_gia_hd'))}"
        counter[key] += 1
        parsed.append({
            "key": key, "ord": counter[key], "dedupe": f"{key}|#{counter[key]}",
            "core": (so_tk, to_iso(g("ngay_dk")), (txt(g("loai_hinh")) or "").upper(), txt(g("doi_tac")), txt(g("van_don")),
                     txt(g("so_hd")), to_iso(g("ngay_hd")), txt(g("ccy")), num(g("ty_gia")), txt(g("trang_thai")),
                     txt(g("ma_hs")), ma, ma_key(ma), txt(g("ten_hang")), txt(g("dvt")), qty, num(g("don_gia_hd")),
                     num(g("don_gia_tt")), num(g("tri_gia_hd")), txt(g("nuoc_xx"))),
            "raw": json.dumps(raw, ensure_ascii=False),
            "rawd": raw,
        })
    return {"parsed": parsed, "types": {n: _col_type(v) for n, v in types_src.items()}}


def duplicate_groups(parsed, decisions=None):
    """Các nhóm dòng trùng khoá (Số TK + Mã hàng + Số lượng + Đơn giá) xuất hiện >1 lần TRONG FILE.

    kind='identical': giống hệt mọi cột (trừ STT) — nhiều khả năng là dòng nhập trùng.
    kind='differs'  : trùng khoá nhưng khác ở một số cột — có thể là các dòng hàng thật sự khác nhau.
    """
    by = {}
    for p in parsed:
        by.setdefault(p["key"], []).append(p)
    out = []
    for key, rows in by.items():
        if len(rows) < 2:
            continue
        first = {k: v for k, v in rows[0]["rawd"].items() if k != "STT"}
        diff = sorted({k for r in rows[1:] for k in set(first) | set(r["rawd"]) if k != "STT" and first.get(k) != r["rawd"].get(k)})
        so_tk, ma, qty, price = key.split("|")
        prior = (decisions or {}).get(key)
        out.append({"key": key, "so_tk": so_tk, "ma_hang": ma, "qty": float(qty), "don_gia": None if price == "None" else float(price),
                    "n": len(rows), "kind": "differs" if diff else "identical", "diff_cols": diff[:8],
                    "ten_hang": rows[0]["core"][13], "prior": prior["keep"] if prior and prior["n"] == len(rows) else None})
    out.sort(key=lambda g: (g["kind"] != "identical", g["so_tk"], g["ma_hang"]))
    return out


def load_decisions(conn):
    return {r["key"]: {"n": r["n_file"], "keep": r["keep"]} for r in conn.execute("SELECT * FROM dup_decisions")}


def apply_database_import(conn, data, filename="", full_sync=True, dedupe_keys=(), dry_run=False):
    """Ghi dữ liệu đã parse vào DB.

    full_sync=True : file là toàn bộ dữ liệu — dòng mới thêm, dòng đổi cập nhật, dòng không còn trong file bị xoá
                     (trừ dòng đã bị trừ lùi: giữ lại và báo).
    full_sync=False: chỉ thêm dòng mới / cập nhật dòng trùng, không xoá gì.
    dedupe_keys    : các khoá trùng mà người dùng đã xác nhận LOẠI BỎ phần thừa (chỉ giữ 1 dòng).
    dry_run        : tính kết quả rồi rollback — dùng để xem trước.
    """
    drop = set(dedupe_keys)
    parsed = [p for p in data["parsed"] if not (p["key"] in drop and p["ord"] > 1)]
    skipped_dups = len(data["parsed"]) - len(parsed)
    existing = {r["dedupe"]: r["id"] for r in conn.execute("SELECT id, dedupe FROM import_lines")}
    min_seq = conn.execute("SELECT COALESCE(MIN(seq),1) FROM import_lines").fetchone()[0]
    n_total = len(parsed)
    added = updated = removed = kept = 0
    cols_sql = ("so_tk,ngay_dk,loai_hinh,doi_tac,van_don,so_hd,ngay_hd,ccy,ty_gia,trang_thai,ma_hs,ma_hang,ma_key,"
                "ten_hang,dvt,qty,don_gia_hd,don_gia_tt,tri_gia_hd,nuoc_xx")
    set_sql = ",".join(f"{c}=?" for c in cols_sql.split(","))
    try:
        for n, p in enumerate(parsed, start=1):
            # file đầy đủ: thứ tự dòng trong file = ưu tiên (dòng trên = mới hơn); file bổ sung: dòng mới xếp lên trên cùng
            seq = n if full_sync else (min_seq - n_total + n)
            lid = existing.pop(p["dedupe"], None)
            if lid is None:
                conn.execute(f"INSERT INTO import_lines(seq,{cols_sql},dedupe,raw) VALUES(?,{','.join('?' * 20)},?,?)",
                             (seq, *p["core"], p["dedupe"], p["raw"]))
                added += 1
            else:
                cur = conn.execute(f"UPDATE import_lines SET {set_sql}, raw=? WHERE id=? AND raw IS NOT ?",
                                   (*p["core"], p["raw"], lid, p["raw"]))
                updated += cur.rowcount
                conn.execute("UPDATE import_lines SET seq=? WHERE id=? AND seq IS NOT ?", (seq, lid, seq))  # đổi vị trí không tính là sửa
        if full_sync and existing:
            for did in existing.values():
                if conn.execute("SELECT 1 FROM allocations WHERE import_line_id=?", (did,)).fetchone():
                    kept += 1
                else:
                    conn.execute("DELETE FROM import_lines WHERE id=?", (did,))
                    removed += 1
        if dry_run:
            conn.rollback()
        else:
            meta = {"columns": [{"name": n, "type": t} for n, t in data["types"].items()]}
            conn.execute("INSERT INTO meta(key,value) VALUES('grid_cols',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                         (json.dumps(meta, ensure_ascii=False),))
            conn.execute("INSERT INTO import_log(filename,mode,rows_in_file,added,updated,removed,kept) VALUES(?,?,?,?,?,?,?)",
                         (filename, "full" if full_sync else "incremental", n_total, added, updated, removed, kept))
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    return {"rows_in_file": n_total, "added": added, "updated": updated, "unchanged": n_total - added - updated,
            "removed": removed, "kept_because_already_allocated": kept, "duplicates_dropped": skipped_dups}


def import_database_sheet(conn, ws, filename="", full_sync=True):
    """Nạp thẳng, giữ nguyên mọi dòng trùng (dùng cho nạp tự động / kiểm thử)."""
    return apply_database_import(conn, parse_database_sheet(ws), filename, full_sync)


def read_item_list(ws, header_cell_text, first_data_row, cols):
    """Đọc danh sách (mã, mô tả, qty, đvt) từ BOM / Export Material."""
    items = []
    for row in ws.iter_rows(min_row=first_data_row, values_only=True):
        ma = txt(row[cols["ma"]]) if len(row) > cols["ma"] else None
        if not ma:
            continue
        items.append({
            "ma_hang": ma,
            "mo_ta": txt(row[cols["mo_ta"]]) if len(row) > cols["mo_ta"] else None,
            "qty": num(row[cols["qty"]], 0.0) if len(row) > cols["qty"] else 0.0,
            "dvt": txt(row[cols["dvt"]]) if len(row) > cols["dvt"] else None,
        })
    return items


def sheets_named(wb, pattern):
    """Các sheet có tên khớp pattern (không phân biệt hoa thường), vd r'bom(\\s*\\(\\d+\\))?'."""
    return [n for n in wb.sheetnames if re.fullmatch(pattern, n.strip(), re.IGNORECASE)]


def read_export_material(ws):
    """Sheet 'Export Material': thông tin lô ở A4:B7, danh sách NVL thực xuất từ dòng 11 (B mã, C mô tả, D Qty, E Unit)."""
    rows = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
    info = {txt(r[0]): r[1] for r in rows if r and r[0]}
    header = {"so_tk_xuat": txt(info.get("Số TK xuất khẩu")), "ngay_tk": to_iso(info.get("Ngày ĐK (TK xuất)")),
              "so_inv": txt(info.get("Số INV")), "ngay_inv": to_iso(info.get("Ngày INV"))}
    lines = [l for l in read_item_list(ws, "", 11, {"ma": 1, "mo_ta": 2, "qty": 3, "dvt": 4}) if l["qty"] and l["qty"] > 0]
    return header, lines


def read_bom_sheet(ws):
    """Sheet 'BOM': thông tin ở A4:B13, danh sách định mức từ dòng 17 (A mã NVL, B mô tả, C Qty used in FG, D Unit)."""
    rows = list(ws.iter_rows(min_row=1, max_row=16, values_only=True))
    info = {txt(r[0]): r[1] for r in rows if r and r[0]}
    header = {"ma_tp": txt(info.get("Mã thành phẩm (Mã sản xuất hàng)")), "ten_hang": txt(info.get("Tên hàng hóa")),
              "hs6": txt(info.get("Mã HS hàng hóa (06 số)")), "dvt": txt(info.get("ĐVT lô hàng"))}
    lines = read_item_list(ws, "", 17, {"ma": 0, "mo_ta": 1, "qty": 2, "dvt": 3})
    return header, lines


def create_event(conn, kind, code, header, lines, multiply_by_lot=False):
    """Tạo lô DRAFT. lines: [{ma_hang, mo_ta, qty, dvt}]. CO: need = định mức × số lượng lô."""
    qty_lo = num(header.get("qty_lo"), 1.0) or 1.0
    with conn:
        cur = conn.execute(
            """INSERT INTO events(kind,code,so_tk_xuat,ngay_tk,so_inv,ngay_inv,ma_tp,ten_hang,hs6,qty_lo,dvt_lo,fob_usd,note,nhap_tu,nhap_den)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (kind, code, header.get("so_tk_xuat"), to_iso(header.get("ngay_tk")), header.get("so_inv") or code,
             to_iso(header.get("ngay_inv")), header.get("ma_tp"), header.get("ten_hang"), header.get("hs6"), qty_lo,
             header.get("dvt_lo"), num(header.get("fob_usd")), header.get("note"),
             to_iso(header.get("nhap_tu")), to_iso(header.get("nhap_den"))))
        eid = cur.lastrowid
        for it in lines:
            need = it["qty"] * (qty_lo if multiply_by_lot else 1.0)
            conn.execute(
                "INSERT INTO event_lines(event_id,ma_hang,ma_key,mo_ta,dvt,dinh_muc,need) VALUES(?,?,?,?,?,?,?)",
                (eid, it["ma_hang"], ma_key(it["ma_hang"]), it.get("mo_ta"), it.get("dvt"),
                 it["qty"] if multiply_by_lot else None, round(need, 6)))
    return eid


def _bom_master(conn, header, lines):
    """Tạo BOM (định mức theo mã thành phẩm) từ sheet BOM của sổ cũ, nếu chưa có."""
    tp = header.get("ma_tp")
    if not tp or not lines or conn.execute("SELECT 1 FROM boms WHERE ma_tp=?", (tp,)).fetchone():
        return
    with conn:
        bid = conn.execute("INSERT INTO boms(ma_tp,ten_hang,hs6,dvt) VALUES(?,?,?,?)",
                           (tp, header.get("ten_hang"), header.get("hs6"), header.get("dvt_lo"))).lastrowid
        for l in lines:
            if l["qty"] and l["qty"] > 0:
                conn.execute("INSERT INTO bom_lines(bom_id,ma_hang,ma_key,mo_ta,dvt,dinh_muc) VALUES(?,?,?,?,?,?)",
                             (bid, l["ma_hang"], ma_key(l["ma_hang"]), l["mo_ta"], l["dvt"], l["qty"]))


def import_legacy_workbook(conn, path, filename=""):
    """Nạp toàn bộ từ file 'SỔ THEO DÕI NVL.xlsm' cũ: 01_DATABASE + BOM (lô C/O) + Export Material (lô XK)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_vba=False)
    out = {"database": import_database_sheet(conn, wb["01_DATABASE"], filename)}

    for sn in ("BOM", "BOM (2)"):
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        rows = list(ws.iter_rows(min_row=1, max_row=16, values_only=True))
        info = {txt(r[0]): r[1] for r in rows if r and r[0]}
        header = {
            "ma_tp": txt(info.get("Mã thành phẩm (Mã sản xuất hàng)")), "ten_hang": txt(info.get("Tên hàng hóa")),
            "so_tk_xuat": txt(info.get("Số TK xuất khẩu")), "ngay_tk": info.get("Ngày ĐK (TK xuất)"),
            "so_inv": txt(info.get("Số INV")), "ngay_inv": info.get("Ngày INV"), "qty_lo": info.get("Số lượng lô hàng"),
            "dvt_lo": txt(info.get("ĐVT lô hàng")), "hs6": txt(info.get("Mã HS hàng hóa (06 số)")),
            "fob_usd": info.get("Trị giá FOB (USD)"), "note": "Nhập từ sổ Excel cũ (sheet BOM)",
        }
        lines = read_item_list(ws, "", 17, {"ma": 0, "mo_ta": 1, "qty": 2, "dvt": 3})
        _bom_master(conn, header, lines)
        if sn == "BOM" and lines and header["so_inv"]:  # lô C/O nháp chỉ tạo từ sheet BOM đang dùng
            exists = conn.execute("SELECT id FROM events WHERE kind='CO' AND code=?", (header["so_inv"],)).fetchone()
            if not exists:
                out["bom_event"] = create_event(conn, "CO", header["so_inv"], header, lines, multiply_by_lot=True)

    if "Export Material" in wb.sheetnames:
        ws = wb["Export Material"]
        rows = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
        info = {txt(r[0]): r[1] for r in rows if r and r[0]}
        lines = read_item_list(ws, "", 11, {"ma": 1, "mo_ta": 2, "qty": 3, "dvt": 4})
        code = txt(info.get("Số INV"))
        if lines and code:
            exists = conn.execute("SELECT id FROM events WHERE kind='XK' AND code=?", (code,)).fetchone()
            if not exists:
                header = {"so_tk_xuat": txt(info.get("Số TK xuất khẩu")), "ngay_tk": info.get("Ngày ĐK (TK xuất)"),
                          "so_inv": code, "ngay_inv": info.get("Ngày INV"), "note": "Nhập từ sổ Excel cũ (Export Material)"}
                out["xk_event"] = create_event(conn, "XK", code, header, lines)
    wb.close()
    return out


def import_item_list_workbook(conn, path):
    """File đơn giản: dòng tiêu đề có 'Mã hàng', 'Qty' (hoặc 'Số lượng'), tuỳ chọn 'Item description','Unit'."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    items = []
    hdr = None
    for row in ws.iter_rows(values_only=True):
        vals = [txt(c) for c in row]
        if hdr is None:
            low = [(v or "").lower() for v in vals]
            if any(x.startswith("mã hàng") for x in low):
                hdr = {"ma": next(i for i, x in enumerate(low) if x.startswith("mã hàng"))}
                for i, x in enumerate(low):
                    if x in ("qty", "số lượng", "qty used in fg", "định mức"):
                        hdr.setdefault("qty", i)
                    elif x in ("item description", "tên hàng", "mô tả"):
                        hdr.setdefault("mo_ta", i)
                    elif x in ("unit", "đvt", "đơn vị tính"):
                        hdr.setdefault("dvt", i)
                if "qty" not in hdr:
                    raise ValueError("Không thấy cột Qty / Số lượng")
            continue
        if row[hdr["ma"]] in (None, ""):
            continue
        items.append({
            "ma_hang": txt(row[hdr["ma"]]), "qty": num(row[hdr["qty"]], 0.0),
            "mo_ta": txt(row[hdr["mo_ta"]]) if "mo_ta" in hdr else None,
            "dvt": txt(row[hdr["dvt"]]) if "dvt" in hdr else None,
        })
    wb.close()
    if hdr is None:
        raise ValueError("Không tìm thấy dòng tiêu đề 'Mã hàng'")
    return items
