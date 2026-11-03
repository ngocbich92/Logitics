"""Dữ liệu chi tiết cho lô / sổ NVL theo loại / form RVC.

Đơn giá CIF (USD) = Đơn giá tính thuế (VNĐ, cột BN của file Data) / Tỷ giá hải quan USD/VNĐ của tuần chứa ngày đăng ký tờ khai nhập
(tỷ giá hải quan công bố theo tuần Thứ Hai–Chủ nhật; xem class Rates).
"""
import datetime as _dt

import engine

KIND_LABEL = {"CO": "C/O", "XK": "Xuất khẩu", "MDK": "Xuất mục đích khác", "TH": "Tiêu huỷ"}


class Rates:
    """Tỷ giá hải quan USD/VNĐ theo TUẦN (thứ Hai–Chủ nhật, mọi ngày trong tuần cùng một tỷ giá).

    Ưu tiên: (1) bảng tỷ giá người dùng nhập / lấy từ website hải quan (bảng fx_rates);
             (2) tỷ giá suy ra từ chính dữ liệu tờ khai: giá trị 'Tỷ giá VNĐ' phổ biến nhất của các dòng USD trong tuần đó;
             (3) tuần gần nhất trước đó có tỷ giá (kèm ghi chú).
    """

    def __init__(self, conn):
        from collections import Counter
        cnt = {}
        for r in conn.execute("SELECT ngay_dk, ty_gia FROM import_lines WHERE UPPER(ccy)='USD' AND ty_gia>0 AND ngay_dk IS NOT NULL"):
            try:
                d = _dt.date.fromisoformat(r["ngay_dk"])
            except ValueError:
                continue
            mon = d - _dt.timedelta(days=d.weekday())
            cnt.setdefault(mon, Counter())[round(r["ty_gia"], 4)] += 1
        self.derived = {m: c.most_common(1)[0][0] for m, c in cnt.items()}
        self.distinct = {m: len(c) for m, c in cnt.items()}
        self.nlines = {m: sum(c.values()) for m, c in cnt.items()}
        self.weeks = sorted(self.derived)
        self.manual = [(r["date_from"], r["date_to"], r["usd"]) for r in
                       conn.execute("SELECT * FROM fx_rates ORDER BY date_from")]

    def get(self, date):
        """-> (tỷ giá, nhãn kỳ áp dụng, ghi chú)"""
        if not date:
            return None, None, "thiếu tỷ giá"
        for f, t, usd in self.manual:
            if f <= date <= t:
                return usd, f"{f}→{t}", ""
        try:
            d = _dt.date.fromisoformat(date)
        except ValueError:
            return None, None, "ngày không hợp lệ"
        mon = d - _dt.timedelta(days=d.weekday())
        label = lambda m: f"{m.isoformat()}→{(m + _dt.timedelta(days=6)).isoformat()}"
        if mon in self.derived:
            return self.derived[mon], label(mon), ""
        prev = [m for m in self.weeks if m < mon]
        if prev:
            m = prev[-1]
            return self.derived[m], label(m), f"tuần {label(mon)} chưa có tỷ giá USD — dùng tuần gần nhất trước đó; hãy nhập tỷ giá hải quan ở Cài đặt"
        return None, None, "thiếu tỷ giá"

    def table(self):
        """Danh sách kỳ tỷ giá để hiển thị: nhập tay + suy ra từ dữ liệu."""
        out = [{"from": f, "to": t, "usd": u, "source": "Nhập tay / website hải quan"} for f, t, u in self.manual]
        for m in reversed(self.weeks):
            out.append({"from": m.isoformat(), "to": (m + _dt.timedelta(days=6)).isoformat(), "usd": self.derived[m],
                        "source": "Suy ra từ dữ liệu tờ khai", "lines": self.nlines[m],
                        "warn": "Tuần này có nhiều tỷ giá USD khác nhau trong dữ liệu" if self.distinct[m] > 1 else ""})
        return out

def save_fx(conn, data):
    """Thêm / sửa một kỳ tỷ giá hải quan USD/VNĐ (từ ngày – đến ngày – tỷ giá)."""
    import re
    from importer import to_iso, num
    f, t, usd = to_iso(data.get("date_from")), to_iso(data.get("date_to")), num(data.get("usd"))
    iso = lambda s: bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", s or ""))
    if not (iso(f) and iso(t)):
        raise ValueError("Ngày không hợp lệ (dd/mm/yyyy)")
    if t < f:
        raise ValueError("Đến ngày phải sau Từ ngày")
    if not usd or usd <= 0:
        raise ValueError("Tỷ giá không hợp lệ")
    with conn:
        conn.execute("INSERT INTO fx_rates(date_from,date_to,usd,note) VALUES(?,?,?,?) ON CONFLICT(date_from) DO UPDATE SET "
                     "date_to=excluded.date_to, usd=excluded.usd, note=excluded.note", (f, t, usd, data.get("note")))
    return {"saved": True}


def import_fx_file(conn, path):
    """Excel: cột 'Từ ngày', 'Đến ngày', 'Tỷ giá USD' (copy từ bảng tỷ giá trên website hải quan)."""
    import openpyxl
    from db import fold
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    names = {"from": ["tu ngay", "from"], "to": ["den ngay", "to"], "usd": ["ty gia usd", "usd", "ty gia", "usd/vnd", "ty gia usd/vnd"]}
    hi, cols = None, {}
    for i, r in enumerate(rows[:15]):
        c = {}
        for j, h in enumerate(r):
            f = fold(h).strip() if h else ""
            for k, ns in names.items():
                if f in ns and k not in c:
                    c[k] = j
        if len(c) == 3:
            hi, cols = i, c
            break
    if hi is None:
        raise ValueError("Cần các cột: Từ ngày, Đến ngày, Tỷ giá USD")
    n, bad = 0, []
    for k, r in enumerate(rows[hi + 1:], start=hi + 2):
        if all(v in (None, "") for v in r):
            continue
        try:
            save_fx(conn, {"date_from": r[cols["from"]], "date_to": r[cols["to"]], "usd": r[cols["usd"]], "note": "Nạp từ file"})
            n += 1
        except ValueError as e:
            bad.append(f"dòng {k}: {e}")
    return {"imported": n, "errors": bad}


def _info(conn, ids):
    ids = list(set(ids))
    out = {}
    for i in range(0, len(ids), 500):
        chunk = ids[i:i + 500]
        for r in conn.execute(f"""SELECT id, so_tk, ngay_dk, loai_hinh, ma_hang, ten_hang, dvt, nuoc_xx, ma_hs, don_gia_hd, don_gia_tt, ccy, trang_thai
                                  FROM import_lines WHERE id IN ({','.join('?' * len(chunk))})""", chunk):
            out[r["id"]] = r
    return out


def _row(base, imp, qty, rates):
    cif = rate = rate_date = None
    note = ""
    if imp is not None:
        rate, rate_date, note = rates.get(imp["ngay_dk"])
        if rate and imp["don_gia_tt"]:
            cif = round(imp["don_gia_tt"] / rate, 6)
        elif not imp["don_gia_tt"]:
            note = (note + "; " if note else "") + "thiếu đơn giá tính thuế"
    return {**base,
            "so_tk": imp["so_tk"] if imp else None, "ngay_tk_nhap": imp["ngay_dk"] if imp else None,
            "loai_hinh": imp["loai_hinh"] if imp else None, "trang_thai_tk": imp["trang_thai"] if imp else None,
            "ten_hang": imp["ten_hang"] if imp else None, "nuoc_xx": imp["nuoc_xx"] if imp else None,
            "ma_hs": imp["ma_hs"] if imp else None, "dvt": (imp["dvt"] if imp else None) or base.get("dvt"),
            "don_gia_hd": imp["don_gia_hd"] if imp else None, "ccy": imp["ccy"] if imp else None,
            "don_gia_tt": imp["don_gia_tt"] if imp else None,
            "ty_gia_usd": round(rate, 4) if rate else None, "ty_gia_ngay": rate_date, "rate_note": note,
            "qty": qty, "don_gia_cif": cif, "tri_gia_usd": round(cif * qty, 4) if (cif is not None and qty) else None}


def lot_rows(conn, event_id, settings, rates=None):
    """Mỗi dòng = một lần trừ (mã hàng × tờ khai nhập). Mã thiếu tồn có thêm dòng 'THIẾU'."""
    rates = rates or Rates(conn)
    ev = conn.execute("SELECT * FROM events WHERE id=?", (event_id,)).fetchone()
    plan = engine.locked_plan(conn, event_id) if ev["status"] == "LOCKED" else engine.plan_event(conn, event_id, settings)
    lines = {r["id"]: r for r in conn.execute("SELECT * FROM event_lines WHERE event_id=?", (event_id,))}
    info = _info(conn, [x["import_line_id"] for p in plan for x in p["parts"]])
    rows = []
    for stt, p in enumerate(plan, start=1):
        ln = lines[p["event_line_id"]]
        base = {"stt": stt, "ma_hang": p["ma_hang"], "mo_ta": p["mo_ta"], "dvt": p["dvt"], "dinh_muc": ln["dinh_muc"],
                "need": p["need"], "shortfall": p["shortfall"]}
        for part in p["parts"]:
            rows.append(_row(base, info.get(part["import_line_id"]), part["qty"], rates))
        if p["shortfall"] > 1e-9:
            rows.append(_row({**base, "is_short": True}, None, p["shortfall"], rates))
    return rows


def ledger(conn, kind, settings=None):
    """Sổ NVL theo loại (C/O, XK, tiêu huỷ, xuất khác): kết quả trừ lùi của TẤT CẢ các lô loại đó.

    Lô đã khoá = đã ghi sổ; lô nháp hiển thị phương án dự kiến (trang_thai = 'Nháp — dự kiến', không tính vào tồn kho).
    """
    import db
    settings = settings or db.get_settings(conn)
    rates = Rates(conn)
    out = []
    evs = conn.execute("SELECT * FROM events WHERE kind=? ORDER BY COALESCE(ngay_tk,'') DESC, id DESC", (kind,)).fetchall()
    for e in evs:
        trang_thai = ("Đã ghi sổ · cố định" if e["frozen"] else "Đã ghi sổ") if e["status"] == "LOCKED" else "Nháp — dự kiến"
        for r in lot_rows(conn, e["id"], settings, rates):
            if r.get("is_short"):
                continue
            out.append({**r, "eid": e["id"], "lo": e["code"], "so_tk_xuat": e["so_tk_xuat"], "ngay_tk_xuat": e["ngay_tk"],
                        "so_co": e["co_so"], "ngay_co": e["co_ngay"], "ma_tp": e["ma_tp"], "trang_thai": trang_thai})
    return out


def rvc(conn, event_id, settings):
    ev = conn.execute("SELECT * FROM events WHERE id=?", (event_id,)).fetchone()
    if ev is None:
        raise KeyError(event_id)
    rows = lot_rows(conn, event_id, settings)
    # Cột "có / không có xuất xứ FTA" (G, H) do người dùng điền tay — không tự phân loại.
    return {"event": dict(ev), "rows": rows, "total_usd": round(sum(r["tri_gia_usd"] or 0 for r in rows), 2),
            "company": {"name": settings["company"], "mst": settings["mst"], "place": settings["place"]},
            "missing_rate": sum(1 for r in rows if r.get("so_tk") and r["don_gia_cif"] is None),
            "shortfall_rows": sum(1 for r in rows if r.get("is_short"))}
