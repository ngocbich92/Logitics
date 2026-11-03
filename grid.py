"""Lưới dữ liệu tờ khai nhập: mọi cột (86 cột báo cáo + Đã trừ/Tồn) đều lọc, sắp xếp, xuất Excel được."""
import json
import re

from openpyxl import Workbook

import db
import report
from db import fold

_PG = [False]   # đang chạy trên PostgreSQL: cột số trong JSON phải ép kiểu số (SQLite tự có kiểu)

USED, REMAIN = "Đã trừ (hệ thống)", "Tồn (hệ thống)"
U_CO, U_XK, U_TH, U_MDK = "Trừ C/O", "Trừ xuất khẩu", "Trừ tiêu huỷ", "Trừ xuất khác"
STATUS = "Tình trạng tồn"
RATE, CIP, NOTE = "Tỷ giá USD/VNĐ (tuần)", "Đơn giá CIP (USD)", "Ghi chú (nhập tay)"
COMPUTED = {USED: "number", U_CO: "number", U_XK: "number", U_TH: "number", U_MDK: "number", REMAIN: "number", STATUS: "text", RATE: "number", CIP: "number", NOTE: "text"}
DEFAULT_VISIBLE = ["Số TK", "Ngày ĐK", "Mã loại hình", "Trạng thái", "Tên đối tác", "Số hóa đơn TM", "Mã hàng", "Tên hàng",
                   "Đơn vị tính", "Nước xuất xứ", "Mã HS", "Số lượng", U_CO, U_XK, U_TH, U_MDK, USED, REMAIN, STATUS,
                   "Ng.Tệ hóa đơn", "Tỷ giá VNĐ", "Đơn giá hóa đơn", "Đơn giá tính thuế", RATE, CIP, NOTE]

BASE = """FROM (SELECT i.id, i.raw, i.ghi_chu AS note, i.qty AS q, COALESCE(a.u,0) AS used, COALESCE(a.u_co,0) AS u_co, COALESCE(a.u_xk,0) AS u_xk,
              COALESCE(a.u_th,0) AS u_th, COALESCE(a.u_mdk,0) AS u_mdk
       FROM import_lines i LEFT JOIN (
           SELECT al.import_line_id, SUM(al.qty) u,
                  SUM(CASE WHEN e.kind='CO' THEN al.qty ELSE 0 END) u_co, SUM(CASE WHEN e.kind='XK' THEN al.qty ELSE 0 END) u_xk,
                  SUM(CASE WHEN e.kind='TH' THEN al.qty ELSE 0 END) u_th, SUM(CASE WHEN e.kind='MDK' THEN al.qty ELSE 0 END) u_mdk
           FROM allocations al JOIN events e ON e.id=al.event_id GROUP BY al.import_line_id) a
       ON a.import_line_id=i.id) t"""


def catalog(conn):
    _PG[0] = db.is_pg(conn)
    r = conn.execute("SELECT value FROM meta WHERE key='grid_cols'").fetchone()
    cols = json.loads(r["value"])["columns"] if r else []
    return [{"name": n, "type": t, "computed": True} for n, t in COMPUTED.items()] + cols


def _expr(name, ctype, params):
    if name == USED:
        return "t.used"
    if name == U_CO:
        return "t.u_co"
    if name == U_XK:
        return "t.u_xk"
    if name == U_TH:
        return "t.u_th"
    if name == U_MDK:
        return "t.u_mdk"
    if name == NOTE:
        return "COALESCE(t.note, '')"
    if name == RATE:
        params.append('$."Ngày ĐK"')
        return "usd_rate(json_extract(t.raw, ?))"
    if name == CIP:  # Đơn giá CIP (USD) = Đơn giá tính thuế (VNĐ) / tỷ giá USD/VNĐ của tuần chứa ngày đăng ký tờ khai
        params.extend(['$."Ngày ĐK"', '$."Đơn giá tính thuế"', '$."Ngày ĐK"'])
        jn = "json_num" if _PG[0] else "json_extract"
        return f"CASE WHEN usd_rate(json_extract(t.raw, ?)) > 0 THEN {jn}(t.raw, ?) / usd_rate(json_extract(t.raw, ?)) END"
    if name == STATUS:
        return "CASE WHEN t.q - t.used <= 0.000000001 THEN 'Hết hàng' WHEN t.used > 0 THEN 'Dùng một phần' ELSE 'Chưa dùng' END"
    if name == REMAIN:
        return "CASE WHEN t.q - t.used > 0 THEN t.q - t.used ELSE 0 END"
    params.append('$."%s"' % name.replace('"', "'"))
    return ("json_num" if (_PG[0] and ctype == "number") else "json_extract") + "(t.raw, ?)"


_NUM = r"-?\d+(?:[.,]\d+)?"


def _to_num(s):
    return float(s.replace(",", "."))


def _condition(name, ctype, text, params, expr=None):
    """Một ô lọc -> mệnh đề SQL. Trả None nếu ô trống."""
    text = (text or "").strip()
    if not text:
        return None
    e = []
    ex = (expr or _expr)(name, ctype, e)

    def add(sql, *vals):
        params.extend(e)
        params.extend(vals)
        return sql

    if ctype == "number":
        m = re.fullmatch(rf"({_NUM})\s*\.\.\s*({_NUM})", text)
        if m:
            return add(f"({ex}) BETWEEN ? AND ?", _to_num(m.group(1)), _to_num(m.group(2)))
        m = re.fullmatch(rf"(>=|<=|>|<|=)\s*({_NUM})", text)
        if m:
            return add(f"({ex}) {m.group(1)} ?", _to_num(m.group(2)))
        if re.fullmatch(_NUM, text):
            return add(f"({ex}) = ?", _to_num(text))
    if ctype == "date":
        def iso(s):
            m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s.strip())
            if m:
                return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
            m = re.fullmatch(r"(\d{1,2})/(\d{4})", s.strip())
            return f"{m.group(2)}-{int(m.group(1)):02d}" if m else s.strip()
        if ".." in text:
            a, b = (iso(x) for x in text.split("..", 1))
            sql = f"({ex}) >= ? AND substr({ex},1,{max(len(b), 1)}) <= ?"
            params.extend(e); params.append(a); params.extend(e); params.append(b)
            return "(" + sql + ")"
        m = re.fullmatch(r"(>=|<=|>|<|=)\s*(.+)", text)
        if m:
            return add(f"({ex}) {m.group(1)} ?", iso(m.group(2)))
        return add(f"instr({ex}, ?) > 0", iso(text))
    # text (và number/date không parse được -> tìm chứa)
    if text.startswith("="):
        terms = [fold(x) for x in text[1:].split("|")]
        parts = []
        for tm in terms:
            parts.append(f"fold(CAST({ex} AS TEXT)) = ?")
            params.extend(e); params.append(tm)
        return "(" + " OR ".join(parts) + ")"
    neg = text.startswith("!")
    terms = [fold(x) for x in (text[1:] if neg else text).split("|") if x.strip()]
    parts = []
    for tm in terms:
        parts.append(f"instr(fold(CAST({ex} AS TEXT)), ?) > 0")
        params.extend(e); params.append(tm)
    sql = "(" + " OR ".join(parts) + ")"
    return f"NOT {sql}" if neg else sql


def _prepare(conn):
    """Đăng ký hàm SQL usd_rate(ngày) = tỷ giá hải quan USD/VNĐ của tuần chứa ngày đó (xem report.Rates)."""
    _PG[0] = db.is_pg(conn)
    rates, cache = report.Rates(conn), {}
    if _PG[0]:  # PostgreSQL: bảng tạm ngày -> tỷ giá, hàm usd_rate() tra bảng này
        conn.execute("CREATE TEMP TABLE IF NOT EXISTS usd_day(d text PRIMARY KEY, usd double precision)")
        conn.execute("DELETE FROM usd_day")
        ds = [r[0] for r in conn.execute("SELECT DISTINCT json_extract(raw, ?) FROM import_lines", ('$."Ngày ĐK"',)) if r[0]]
        conn.executemany("INSERT INTO usd_day VALUES(?,?)", [(d, rates.get(d)[0]) for d in ds])
        return

    def usd_rate(d):
        if d not in cache:
            cache[d] = rates.get(d)[0]
        return cache[d]
    conn.create_function("usd_rate", 1, usd_rate)


def build_where(conn, filters, q=""):
    _prepare(conn)
    cols = {c["name"]: c["type"] for c in catalog(conn)}
    conds, params = [], []
    for name, text in (filters or {}).items():
        if name in cols:
            c = _condition(name, cols[name], text, params)
            if c:
                conds.append(c)
    if q and q.strip():
        conds.append("instr(fold(t.raw), ?) > 0")
        params.append(fold(q.strip()))
    return ("WHERE " + " AND ".join(conds)) if conds else "", params


def _order(conn, sort, direction):
    cols = {c["name"]: c["type"] for c in catalog(conn)}
    if sort in cols:
        p = []
        ex = _expr(sort, cols[sort], p)
        d = "DESC" if direction == "desc" else "ASC"
        return f"ORDER BY ({ex}) IS NULL, ({ex}) {d}, t.id", p + p
    return "ORDER BY t.id", []


def _select_cols(conn, visible):
    cat = {c["name"]: c for c in catalog(conn)}
    use = [c for c in (visible or DEFAULT_VISIBLE) if c in cat]
    return use or [c["name"] for c in list(cat.values())[:10]]


def query(conn, filters, q, sort, direction, page, size, visible):
    where, wparams = build_where(conn, filters, q)
    order, oparams = _order(conn, sort, direction)
    sm = conn.execute(f"SELECT COUNT(*) n, SUM(CASE WHEN t.q - t.used > 0.000000001 THEN 1 ELSE 0 END) con, "
                      f"SUM(CASE WHEN t.used > 0 THEN 1 ELSE 0 END) dung {BASE} {where}", wparams).fetchone()
    total = sm["n"]
    summary = {"con_ton": sm["con"] or 0, "het": total - (sm["con"] or 0), "da_dung": sm["dung"] or 0}
    use = _select_cols(conn, visible)
    cat = {c["name"]: c for c in catalog(conn)}
    sel, sparams = ["t.id AS _id"], []
    for i, name in enumerate(use):
        sel.append(f"{_expr(name, cat[name]['type'], sparams)} AS c{i}")
    sql = f"SELECT {', '.join(sel)} {BASE} {where} {order} LIMIT ? OFFSET ?"
    rows = conn.execute(sql, sparams + wparams + oparams + [size, (page - 1) * size]).fetchall()
    return {"total": total, "summary": summary, "columns": use,
            "rows": [[r[f"c{i}"] for i in range(len(use))] for r in rows], "ids": [r["_id"] for r in rows]}


def distinct(conn, name, filters, q, limit=300):
    cat = {c["name"]: c for c in catalog(conn)}
    if name not in cat:
        return []
    others = {k: v for k, v in (filters or {}).items() if k != name}
    where, wparams = build_where(conn, others, q)
    p = []
    ex = _expr(name, cat[name]["type"], p)
    sql = f"SELECT v, COUNT(*) n FROM (SELECT {ex} v {BASE} {where}) s WHERE v IS NOT NULL GROUP BY v ORDER BY n DESC LIMIT ?"
    return [{"v": r["v"], "n": r["n"]} for r in conn.execute(sql, p + wparams + [limit])]


def export_xlsx(conn, filters, q, sort, direction, visible, cap=300000):
    from io import BytesIO
    where, wparams = build_where(conn, filters, q)
    order, oparams = _order(conn, sort, direction)
    use = _select_cols(conn, visible)
    cat = {c["name"]: c for c in catalog(conn)}
    sel, sparams = [], []
    for i, name in enumerate(use):
        sel.append(f"{_expr(name, cat[name]['type'], sparams)} AS c{i}")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Du lieu loc")
    ws.append(use)
    for r in conn.execute(f"SELECT {', '.join(sel)} {BASE} {where} {order} LIMIT ?", sparams + wparams + oparams + [cap]):
        ws.append(list(r))
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
