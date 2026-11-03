"""Nạp lô cũ khi đã biết chính xác tờ khai nhập nào bị trừ bao nhiêu (không dùng LIFO).

File Excel, mỗi dòng = một lần trừ:  Mã lô/INV | Loại | Số TK nhập | Mã hàng | Số lượng  (+ cột tuỳ chọn thông tin lô).
Mỗi lô tạo ra một lô ĐÃ KHOÁ (mặc định 'cố định - đã nộp'), phần trừ ghi thẳng vào đúng tờ khai nhập.
"""
import io
import re
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font

import db
from db import fold, ma_key
from importer import num, to_iso, txt

EPS = 1e-9
KINDS = {"co": "CO", "c/o": "CO", "xk": "XK", "xuat khau": "XK", "mdk": "MDK", "muc dich khac": "MDK",
         "th": "TH", "tieu huy": "TH"}
ALIASES = {  # tên cột (đã bỏ dấu, chữ thường) -> field
    "lot": ["ma lo", "so inv", "inv", "lo", "ma lo / inv", "ma lo/inv", "lo hang"],
    "kind": ["loai", "loai lo"],
    "so_tk_nhap": ["so tk nhap", "tk nhap", "so to khai nhap", "so tk"],
    "ma_hang": ["ma hang", "ma nvl"],
    "qty": ["so luong", "so luong tru", "qty", "luong tru"],
    "so_tk_xuat": ["so tk xuat", "tk xuat"],
    "ngay_tk": ["ngay tk xuat", "ngay tk"],
    "so_co": ["so c/o", "so co"],
    "ngay_co": ["ngay c/o", "ngay co"],
    "ma_tp": ["ma thanh pham"],
    "ngay_inv": ["ngay inv"],
}
REQUIRED = ["lot", "so_tk_nhap", "ma_hang", "qty"]


def _map_headers(row):
    cols = {}
    for i, h in enumerate(row):
        f = fold(h).strip() if h else ""
        for field, names in ALIASES.items():
            if f in names and field not in cols:
                cols[field] = i
    return cols


def parse_file(path):
    """-> list[dict] các dòng trừ (kèm số dòng Excel)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    hdr_i = next((i for i, r in enumerate(rows[:15]) if all(f in _map_headers(r) for f in REQUIRED)), None)
    if hdr_i is None:
        raise ValueError("Không tìm thấy dòng tiêu đề. Cần các cột: Mã lô, Số TK nhập, Mã hàng, Số lượng "
                         "(tải file mẫu để xem định dạng).")
    cols = _map_headers(rows[hdr_i])
    out = []
    for n, r in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
        g = lambda f: r[cols[f]] if f in cols and cols[f] < len(r) else None
        if all(v in (None, "") for v in r):
            continue
        tk = g("so_tk_nhap")
        out.append({
            "row": n, "lot": txt(g("lot")) or "", "kind": txt(g("kind")) or "",
            "so_tk": str(int(tk)) if isinstance(tk, (int, float)) else (txt(tk) or ""),
            "ma_hang": txt(g("ma_hang")) or "", "qty": num(g("qty")),
            "so_tk_xuat": txt(g("so_tk_xuat")), "ngay_tk": to_iso(g("ngay_tk")), "so_co": txt(g("so_co")),
            "ngay_co": to_iso(g("ngay_co")), "ma_tp": txt(g("ma_tp")), "ngay_inv": to_iso(g("ngay_inv")),
        })
    if not out:
        raise ValueError("File không có dòng dữ liệu")
    return out


def _remaining(conn, planned, tk, key):
    res = []
    for r in conn.execute("""SELECT i.id, i.qty - COALESCE((SELECT SUM(a.qty) FROM allocations a WHERE a.import_line_id=i.id),0) rem,
                                    i.loai_hinh, i.trang_thai FROM import_lines i WHERE i.so_tk=? AND i.ma_key=? ORDER BY i.seq""", (tk, key)):
        res.append((r["id"], r["rem"] - planned.get(r["id"], 0.0), r["loai_hinh"], r["trang_thai"]))
    return res


def apply(conn, rows, default_kind="XK", replace=False, frozen=True, dry_run=True, settings=None):
    """Kiểm tra và (nếu dry_run=False) ghi các lô. Dòng lỗi được bỏ qua và liệt kê trong 'problems'."""
    settings = settings or db.get_settings(conn)
    problems, lots = [], OrderedDict()
    try:
        # gom theo lô
        for r in rows:
            if not r["lot"]:
                problems.append({**_p(r), "msg": "Thiếu Mã lô / INV"})
                continue
            k = KINDS.get(fold(r["kind"]).strip(), None) if r["kind"] else default_kind
            if k is None:
                problems.append({**_p(r), "msg": f"Loại '{r['kind']}' không hợp lệ (CO / XK / MDK / TH)"})
                continue
            lots.setdefault((r["lot"], k), []).append(r)
        result = []
        for (code, kind), lrows in lots.items():
            planned = {}  # phần đã giữ chỗ trong lô này; sau khi lô ghi xuống DB thì tính qua bảng allocations
            ex = conn.execute("SELECT id, status FROM events WHERE code=? AND kind=?", (code, kind)).fetchone()
            if ex and not replace:
                for r in lrows:
                    problems.append({**_p(r), "msg": "Lô đã tồn tại — chọn 'Thay thế lô cùng mã' nếu muốn nạp lại"})
                continue
            if ex and replace:
                conn.execute("DELETE FROM events WHERE id=?", (ex["id"],))  # cascade: dòng NVL + phần đã trừ
            warn_types = {t.strip().upper() for t in settings["types_co" if kind == "CO" else "types_other"].split(",")}
            per_line = OrderedDict()  # ma_key -> list[(import_id, qty)]
            ok = 0
            for r in lrows:
                msg = None
                if r["qty"] is None or r["qty"] <= 0:
                    msg = "Số lượng không hợp lệ"
                elif not r["so_tk"] or not r["ma_hang"]:
                    msg = "Thiếu Số TK nhập hoặc Mã hàng"
                else:
                    key = ma_key(r["ma_hang"])
                    lines = _remaining(conn, planned, r["so_tk"], key)
                    if not lines:
                        msg = "Không có tờ khai nhập này / mã hàng này trong dữ liệu"
                    else:
                        left, parts = r["qty"], []
                        for lid, rem, loai, tt in lines:
                            if left <= EPS:
                                break
                            if rem <= EPS:
                                continue
                            take = min(rem, left)
                            parts.append((lid, round(take, 6), loai, tt))
                            left -= take
                        if left > EPS:
                            have = sum(p[1] for p in parts)
                            msg = f"Vượt tồn còn lại của TK {r['so_tk']} (còn {round(have, 6)}, cần {r['qty']})"
                        else:
                            for lid, take, loai, tt in parts:
                                planned[lid] = planned.get(lid, 0.0) + take
                                per_line.setdefault((r["ma_hang"], key), []).append((lid, take))
                                if (loai or "").upper() not in warn_types:
                                    problems.append({**_p(r), "level": "warn",
                                                     "msg": f"Loại hình {loai} không nằm trong danh sách loại hình được trừ cho {kind}"})
                                elif settings.get("only_cleared") == "1" and "thong quan" not in fold(tt or ""):
                                    problems.append({**_p(r), "level": "warn", "msg": f"Tờ khai nhập '{tt}' chưa thông quan"})
                if msg:
                    problems.append({**_p(r), "msg": msg})
                else:
                    ok += 1
            if per_line:
                h = lrows[0]
                first = lambda f: next((x[f] for x in lrows if x.get(f)), None)
                cur = conn.execute(
                    """INSERT INTO events(kind,code,so_tk_xuat,ngay_tk,so_inv,ngay_inv,ma_tp,co_so,co_ngay,status,frozen,locked_at,note)
                       VALUES(?,?,?,?,?,?,?,?,?,'LOCKED',?,datetime('now','localtime'),'Nạp lô cũ theo tờ khai chỉ định')""",
                    (kind, code, first("so_tk_xuat"), first("ngay_tk"), code, first("ngay_inv"), first("ma_tp"), first("so_co"),
                     first("ngay_co"), 1 if frozen else 0))
                eid = cur.lastrowid
                for (ma, key), parts in per_line.items():
                    need = round(sum(q for _, q in parts), 6)
                    lcur = conn.execute("INSERT INTO event_lines(event_id,ma_hang,ma_key,need) VALUES(?,?,?,?)", (eid, ma, key, need))
                    for lid, q in parts:
                        conn.execute("INSERT INTO allocations(event_id,event_line_id,import_line_id,qty) VALUES(?,?,?,?)",
                                     (eid, lcur.lastrowid, lid, q))
            result.append({"lot": code, "kind": kind, "rows": len(lrows), "ok_rows": ok, "replaced": bool(ex),
                           "created": bool(per_line)})
        if dry_run:
            conn.rollback()
        else:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    errors = [p for p in problems if p.get("level") != "warn"]
    return {"lots": result, "problems": problems, "n_errors": len(errors), "n_warnings": len(problems) - len(errors),
            "rows": len(rows), "ok_rows": sum(x["ok_rows"] for x in result)}


def _p(r):
    return {"row": r["row"], "lot": r["lot"], "so_tk": r["so_tk"], "ma_hang": r["ma_hang"], "qty": r["qty"]}


def template_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lo cu"
    heads = ["Mã lô / INV", "Loại (CO/XK/MDK/TH)", "Số TK nhập", "Mã hàng", "Số lượng", "Số TK xuất", "Ngày TK xuất",
             "Số C/O", "Ngày C/O", "Mã thành phẩm"]
    ws.append(heads)
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.append(["VNOUT00500", "CO", "108615947530", "825-00058", 12, "308727606999", "15/07/2026", "", "", "800-00332"])
    ws.append(["VNOUT00500", "CO", "108610281140", "840-00331", 3, "308727606999", "15/07/2026", "", "", "800-00332"])
    ws.append(["PXK0044", "MDK", "108610281140", "840-00182", 20])
    ws.append(["BB-TH-001", "TH", "108609957520", "110-00006", 150])
    for col, w in zip("ABCDEFGHIJ", [16, 20, 16, 16, 10, 16, 14, 12, 12, 16]):
        ws.column_dimensions[col].width = w
    ws2 = wb.create_sheet("Huong dan")
    for line in ["Mỗi dòng = một lần trừ: lô nào lấy bao nhiêu của mã hàng nào từ tờ khai nhập nào.",
                 "Cột bắt buộc: Mã lô / INV, Số TK nhập, Mã hàng, Số lượng. Cột Loại để trống = dùng loại mặc định chọn trên màn hình.",
                 "Cột Số TK xuất, Ngày TK xuất, Số C/O, Ngày C/O, Mã thành phẩm: tuỳ chọn, lấy từ dòng đầu của lô.",
                 "Một mã hàng có thể trừ từ nhiều tờ khai: viết nhiều dòng cùng Mã lô + Mã hàng, khác Số TK nhập.",
                 "Xoá 3 dòng ví dụ trước khi nạp."]:
        ws2.append([line])
    ws2.column_dimensions["A"].width = 120
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
