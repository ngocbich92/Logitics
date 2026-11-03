"""Trừ lùi (LIFO/FIFO) nguyên vật liệu theo từng lô xuất.

Quy tắc lấy từ sổ Excel hiện tại:
  * Chỉ trừ vào dòng nhập có loại hình hợp lệ (C/O: E11,E15 — XK/MDK/TH: E11,E15,E13).
  * Mỗi mã hàng trừ lần lượt từ dòng nhập mới nhất (LIFO) cho tới khi đủ nhu cầu.
  * Một dòng nhập đã bị trừ ở lô trước (đã khoá) thì chỉ còn phần dư cho lô sau.
Khác với Excel: nếu thiếu tồn thì hiển thị phần THIẾU thay vì lặng lẽ trừ ít hơn.
"""
import re

EPS = 1e-9


def _round(x):
    return round(x + 0.0, 6)


def eligible_types(conn_settings, kind):
    key = "types_co" if kind == "CO" else "types_other"
    return [t.strip().upper() for t in conn_settings[key].split(",") if t.strip()]


def stock_rows(conn, ma_key, types, order, only_cleared=False, date_from=None, date_to=None):
    """Các dòng nhập còn tồn của 1 mã hàng, theo thứ tự ưu tiên trừ."""
    direction = "DESC" if order == "LIFO" else "ASC"
    marks = ",".join("?" * len(types))
    cleared = "AND fold(i.trang_thai) LIKE '%thong quan%'" if only_cleared else ""
    extra = []
    if date_from:  # khoảng ngày ĐK tờ khai nhập do người dùng chọn cho từng lô
        cleared += " AND i.ngay_dk >= ?"
        extra.append(date_from)
    if date_to:
        cleared += " AND i.ngay_dk <= ?"
        extra.append(date_to)
    # cùng ngày: giữ thứ tự dòng nguồn (seq nhỏ = đứng trên = mới hơn)
    sql = f"""
        SELECT i.id, i.so_tk, i.ngay_dk, i.qty,
               i.qty - COALESCE((SELECT SUM(a.qty) FROM allocations a WHERE a.import_line_id=i.id),0) AS remain
        FROM import_lines i
        WHERE i.ma_key=? AND UPPER(i.loai_hinh) IN ({marks}) {cleared}
        ORDER BY i.ngay_dk {direction}, i.seq ASC
    """
    return [r for r in conn.execute(sql, [ma_key, *types, *extra]) if r["remain"] > EPS]


def plan_event(conn, event_id, settings):
    """Tính phương án trừ cho 1 lô (chưa ghi DB). Trả về list theo từng dòng nhu cầu."""
    ev = conn.execute("SELECT * FROM events WHERE id=?", (event_id,)).fetchone()
    if ev is None:
        raise KeyError(event_id)
    types = eligible_types(settings, ev["kind"])
    iso = lambda s: s if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s or "") else None
    d_from, d_to = iso(ev["nhap_tu"]), iso(ev["nhap_den"])  # trống = không giới hạn
    used = {}  # import_line_id -> đã trừ trong chính phương án này (mã hàng trùng dòng)
    result = []
    for line in conn.execute("SELECT * FROM event_lines WHERE event_id=? ORDER BY id", (event_id,)):
        need = line["need"] or 0.0
        left = need
        parts = []
        for s in stock_rows(conn, line["ma_key"], types, settings["order"], settings.get("only_cleared") == "1", d_from, d_to):
            if left <= EPS:
                break
            avail = s["remain"] - used.get(s["id"], 0.0)
            if avail <= EPS:
                continue
            take = min(avail, left)
            used[s["id"]] = used.get(s["id"], 0.0) + take
            left -= take
            parts.append({"import_line_id": s["id"], "so_tk": s["so_tk"], "ngay_dk": s["ngay_dk"], "qty": _round(take)})
        result.append({
            "event_line_id": line["id"], "ma_hang": line["ma_hang"], "mo_ta": line["mo_ta"],
            "dvt": line["dvt"], "need": _round(need), "allocated": _round(need - max(left, 0.0)),
            "shortfall": _round(max(left, 0.0)), "parts": parts,
        })
    return result


def lock_event(conn, event_id, settings, allow_shortfall=False, frozen=False):
    """Ghi phương án trừ và khoá lô. Mặc định từ chối nếu còn thiếu tồn."""
    ev = conn.execute("SELECT status FROM events WHERE id=?", (event_id,)).fetchone()
    if ev is None:
        raise KeyError(event_id)
    if ev["status"] == "LOCKED":
        raise ValueError("Lô đã khoá")
    plan = plan_event(conn, event_id, settings)
    short = [p for p in plan if p["shortfall"] > EPS]
    if short and not allow_shortfall:
        raise ValueError(f"Thiếu tồn ở {len(short)} mã hàng — kiểm tra lại hoặc khoá với allow_shortfall")
    with conn:
        for p in plan:
            for part in p["parts"]:
                conn.execute(
                    "INSERT INTO allocations(event_id,event_line_id,import_line_id,qty) VALUES(?,?,?,?)",
                    (event_id, p["event_line_id"], part["import_line_id"], part["qty"]))
        conn.execute("UPDATE events SET status='LOCKED', frozen=?, locked_at=datetime('now','localtime') WHERE id=?",
                     (1 if frozen else 0, event_id))
    return plan


def unlock_event(conn, event_id):
    """Huỷ khoá: trả lại số lượng đã trừ vào tồn kho."""
    with conn:
        conn.execute("DELETE FROM allocations WHERE event_id=?", (event_id,))
        conn.execute("UPDATE events SET status='DRAFT', frozen=0, locked_at=NULL WHERE id=?", (event_id,))


def locked_plan(conn, event_id):
    """Phương án đã khoá (đọc từ bảng allocations), cùng định dạng plan_event."""
    out = []
    for line in conn.execute("SELECT * FROM event_lines WHERE event_id=? ORDER BY id", (event_id,)):
        parts = [dict(r) for r in conn.execute("""
            SELECT a.import_line_id, i.so_tk, i.ngay_dk, a.qty FROM allocations a
            JOIN import_lines i ON i.id=a.import_line_id WHERE a.event_line_id=? ORDER BY a.id""", (line["id"],))]
        got = _round(sum(p["qty"] for p in parts))
        out.append({"event_line_id": line["id"], "ma_hang": line["ma_hang"], "mo_ta": line["mo_ta"],
                    "dvt": line["dvt"], "need": _round(line["need"] or 0), "allocated": got,
                    "shortfall": _round(max((line["need"] or 0) - got, 0)), "parts": parts})
    return out


def priority_key(ev, settings):
    """Thứ tự ưu tiên khi tính lại: C/O trước (như sổ Excel: 02_CO trừ trước 03_XK...), rồi theo ngày TK xuất, rồi theo ID."""
    co_first = settings.get("priority", "CO_FIRST") == "CO_FIRST"
    return (0 if (co_first and ev["kind"] == "CO") else 1, ev["ngay_tk"] or "9999-12-31", ev["id"])


def _signature(conn, event_id):
    return sorted((r["event_line_id"], r["import_line_id"], round(r["qty"], 6))
                  for r in conn.execute("SELECT event_line_id, import_line_id, qty FROM allocations WHERE event_id=?", (event_id,)))


def rebuild(conn, settings, dry_run=True):
    """Tính lại phương án trừ của mọi lô đã khoá (trừ lô 'đã nộp/cố định') theo đúng thứ tự ưu tiên.

    Dùng sau khi bổ sung các lô cũ, hoặc thêm lô C/O sau khi đã có lô XK. Trả về danh sách lô bị thay đổi.
    dry_run=True chỉ xem trước (rollback).
    """
    evs = [e for e in conn.execute("SELECT * FROM events WHERE status='LOCKED' AND frozen=0")]
    evs.sort(key=lambda e: priority_key(e, settings))
    old = {e["id"]: _signature(conn, e["id"]) for e in evs}
    old_short = {e["id"]: sum(p["shortfall"] for p in locked_plan(conn, e["id"])) for e in evs}
    changes = []
    try:
        for e in evs:
            conn.execute("DELETE FROM allocations WHERE event_id=?", (e["id"],))
        for e in evs:
            plan = plan_event(conn, e["id"], settings)
            for p in plan:
                for part in p["parts"]:
                    conn.execute("INSERT INTO allocations(event_id,event_line_id,import_line_id,qty) VALUES(?,?,?,?)",
                                 (e["id"], p["event_line_id"], part["import_line_id"], part["qty"]))
            new = _signature(conn, e["id"])
            if new != old[e["id"]]:
                short = sum(p["shortfall"] for p in plan)
                diff_lines = len({x[0] for x in set(new) ^ set(old[e["id"]])})
                changes.append({"id": e["id"], "kind": e["kind"], "code": e["code"], "ngay_tk": e["ngay_tk"],
                                "lines_changed": diff_lines, "shortfall_before": _round(old_short[e["id"]]), "shortfall_after": _round(short)})
        if dry_run:
            conn.rollback()
        else:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    return {"lots_checked": len(evs), "changed": changes}


def set_frozen(conn, event_id, frozen):
    ev = conn.execute("SELECT status FROM events WHERE id=?", (event_id,)).fetchone()
    if ev is None:
        raise KeyError(event_id)
    if frozen and ev["status"] != "LOCKED":
        raise ValueError("Chỉ cố định được lô đã khoá")
    with conn:
        conn.execute("UPDATE events SET frozen=? WHERE id=?", (1 if frozen else 0, event_id))
