"""Máy chủ web cục bộ (chỉ stdlib + openpyxl). Chạy: python server.py  ->  http://127.0.0.1:8765"""
import json
import os
import re
import sys
import tempfile
import threading
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import db
import engine
import exporter
import grid
import importer
import lots_import
import report
import rvc_form
import invoice_scan
import outbound
import outbound_forms
import boms

STATIC = Path(__file__).parent / "static"
PORT = int(os.environ.get("NVL_PORT", "8765"))
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOCK = threading.Lock()  # một người dùng tại một thời điểm ghi dữ liệu


def _grid_args(q):
    def js(k, default):
        try:
            return json.loads((q.get(k) or [""])[0]) or default
        except json.JSONDecodeError:
            return default
    return {"filters": js("filters", {}), "q": (q.get("q") or [""])[0], "sort": (q.get("sort") or [""])[0],
            "dir": (q.get("dir") or ["asc"])[0], "cols": js("cols", None),
            "page": max(int((q.get("page") or ["1"])[0]), 1), "size": min(int((q.get("size") or ["100"])[0]), 500)}


def rows(cur):
    return [dict(r) for r in cur]


def _ob_stats(conn):
    """Số liệu tổng quan cho trang Outbound: đếm theo Status và Forwarder trong List shipment, shipment mới nhất."""
    out = {"status": [], "forwarder": [], "recent": []}
    info = outbound.table_info(conn, "list")
    if info:
        letter = {c["name"]: c["letter"] for c in info["cols"]}
        for key, name in (("status", "Status"), ("forwarder", "Forwarder")):
            if name in letter:
                out[key] = rows(conn.execute(
                    f"SELECT v, COUNT(*) n FROM (SELECT json_extract(cells,'$.{letter[name]}') v FROM ob_rows WHERE tbl='list') s "
                    f"WHERE v IS NOT NULL GROUP BY v ORDER BY n DESC LIMIT 12"))
    src = outbound.table_info(conn, "source")
    if src:
        # cột theo chữ cái của file gốc: G Consignee, R Description of Packages and Goods, O Pick up date, AS Request pick up,
        # AB Type of container, AK Ter of delivery, AT/AF Status, AU FWD
        have = {c["letter"] for c in src["cols"]}
        j = lambda *ls: "COALESCE(" + ",".join(f"json_extract(cells,'$.{l}')" for l in ls if l in have) + ",NULL)"
        out["recent"] = rows(conn.execute(
            f"SELECT id rid, key id, {j('G')} consignee, {j('R')} goods, {j('O', 'AS')} pickup, {j('AB')} container, {j('AK')} term, "
            f"{j('AT', 'AF')} status, {j('AU')} fwd FROM ob_rows WHERE tbl='source' ORDER BY pos DESC LIMIT 8"))
    return out


def api_get(conn, path, q):
    st = db.get_settings(conn)
    if path == "/api/ob/report":
        sts = [s for s in (q.get("statuses") or [""])[0].split("|") if s]
        r = outbound.report(conn, sts or None, (q.get("from") or [""])[0] or None, (q.get("to") or [""])[0] or None)
        r.pop("shipments")
        return r
    if path == "/api/ob/goods-grid":
        a = _grid_args(q)
        return outbound.goods_query(conn, a["filters"], a["q"], a["sort"], a["dir"], a["page"], a["size"])
    m = re.fullmatch(r"/api/ob/goods-line/(\d+)", path)
    if m:
        return outbound.goods_line_get(conn, int(m.group(1)))
    if path == "/api/ob/invoices":
        return outbound.invoices_list(conn)
    if path == "/api/ob/search":
        return outbound.search(conn, (q.get("q") or [""])[0])
    if path == "/api/ob/overview":
        return {"tables": outbound.overview(conn), "stats": _ob_stats(conn)}
    if path == "/api/ob/grid":
        a = _grid_args(q)
        return outbound.query(conn, (q.get("tbl") or [""])[0], a["filters"], a["q"], a["sort"], a["dir"], a["page"], a["size"])
    if path == "/api/ob/distinct":
        a = _grid_args(q)
        return outbound.distinct(conn, (q.get("tbl") or [""])[0], (q.get("col") or [""])[0], a["filters"], a["q"])
    m = re.fullmatch(r"/api/ob/row/(\d+)", path)
    if m:
        return outbound.get_row(conn, int(m.group(1)))
    if path == "/api/ob/forms":
        return {"forms": [{"sheet": s, "title": t, "kind": k} for s, t, k in outbound_forms.FORMS],
                "ship_ids": outbound_forms.shipment_ids(conn), "inv_ids": outbound_forms.invoice_ids(conn)}
    if path == "/api/ob/form-html":
        sheet = (q.get("sheet") or [""])[0]
        if sheet not in outbound_forms.KIND:
            raise ValueError("Biểu mẫu không hợp lệ")
        ship = (q.get("ship") or [""])[0].strip()
        inv = (q.get("inv") or [""])[0].strip() or None
        if not ship and outbound_forms.KIND[sheet] == "ship":
            raise ValueError("Chọn Shipment ID")
        html_, errs = outbound_forms.preview_html(conn, ship, inv, sheet)
        return {"html": html_, "errors": errs}
    if path == "/api/summary":
        s = conn.execute("""SELECT COUNT(*) n, COALESCE(SUM(qty),0) nhap FROM import_lines""").fetchone()
        used = conn.execute("SELECT COALESCE(SUM(qty),0) FROM allocations").fetchone()[0]
        ev = rows(conn.execute("SELECT status, kind, COUNT(*) n FROM events GROUP BY status, kind"))
        tk = conn.execute("SELECT COUNT(DISTINCT so_tk) FROM import_lines").fetchone()[0]
        codes = conn.execute("SELECT COUNT(DISTINCT ma_key) FROM import_lines").fetchone()[0]
        return {"lines": s["n"], "tk": tk, "codes": codes, "nhap": s["nhap"], "xuat": used, "events": ev, "settings": st}
    if path == "/api/grid/columns":
        if (q.get("raw") or [""])[0] == "1":  # trang Database: chỉ cột dữ liệu gốc, không có cột trừ lùi / tồn kho
            cols = [c for c in grid.catalog(conn) if not c.get("computed")]
            return {"columns": cols, "default": [c["name"] for c in cols]}
        return {"columns": grid.catalog(conn), "default": grid.DEFAULT_VISIBLE}
    if path == "/api/grid":
        a = _grid_args(q)
        return grid.query(conn, a["filters"], a["q"], a["sort"], a["dir"], a["page"], a["size"], a["cols"])
    if path == "/api/grid/distinct":
        a = _grid_args(q)
        return grid.distinct(conn, (q.get("col") or [""])[0], a["filters"], a["q"])
    if path == "/api/import/log":
        return rows(conn.execute("SELECT * FROM import_log ORDER BY id DESC LIMIT 20"))
    if path == "/api/stock/by-code":
        return rows(conn.execute("""
            SELECT i.ma_key, MIN(i.ten_hang) ten_hang, MIN(i.dvt) dvt, COUNT(*) dong, SUM(i.qty) nhap,
                   COALESCE(SUM((SELECT SUM(a.qty) FROM allocations a WHERE a.import_line_id=i.id)),0) xuat
            FROM import_lines i GROUP BY i.ma_key ORDER BY i.ma_key"""))
    if path == "/api/ledger":
        kind = (q.get("kind") or ["CO"])[0]
        if kind not in report.KIND_LABEL:
            raise ValueError("Loại không hợp lệ")
        return report.ledger(conn, kind, st)
    if path == "/api/boms":
        return rows(conn.execute("""SELECT b.*, (SELECT COUNT(*) FROM bom_lines l WHERE l.bom_id=b.id) n_lines
                                    FROM boms b ORDER BY b.ma_tp"""))
    m = re.fullmatch(r"/api/boms/(\d+)", path)
    if m:
        return boms.detail(conn, int(m.group(1)), st)
    if path == "/api/fx":
        return report.Rates(conn).table()
    m = re.fullmatch(r"/api/rvc-html/(\d+)", path)
    if m:
        wb, d = rvc_form.build(conn, int(m.group(1)), st)
        return {"html": rvc_form.to_html(wb["RVC"]), "total_usd": d["total_usd"],
                "missing_rate": d["missing_rate"], "shortfall_rows": d["shortfall_rows"], "status": d["event"]["status"]}
    m = re.fullmatch(r"/api/rvc/(\d+)", path)
    if m:
        return report.rvc(conn, int(m.group(1)), st)
    if path == "/api/events":
        return rows(conn.execute("""
            SELECT e.*, (SELECT COUNT(*) FROM event_lines l WHERE l.event_id=e.id) n_lines FROM events e ORDER BY e.id DESC"""))
    m = re.fullmatch(r"/api/events/(\d+)", path)
    if m:
        eid = int(m.group(1))
        ev = conn.execute("SELECT * FROM events WHERE id=?", (eid,)).fetchone()
        if not ev:
            raise KeyError("Không có lô")
        plan = engine.locked_plan(conn, eid) if ev["status"] == "LOCKED" else engine.plan_event(conn, eid, st)
        return {"event": dict(ev), "plan": plan, "rows": report.lot_rows(conn, eid, st),
                "short": sum(1 for p in plan if p["shortfall"] > 1e-9)}
    raise KeyError(path)


STAGES = {}  # file đã đọc, đang chờ người dùng xác nhận (giữ trong RAM, tối đa 3 file)


STAGE_DIR = db.DB_PATH.parent / "staging"
FILES = {}  # token -> đường dẫn file đã tải lên, chờ chọn sheet
FILE_NAMES = {}  # token -> tên file gốc


def _save_upload(body, keep=2):
    """Lưu file đã tải lên để chọn sheet / xem trước mà không phải tải lại (file 50MB)."""
    STAGE_DIR.mkdir(parents=True, exist_ok=True)
    for old in list(FILES):
        if len(FILES) >= keep:  # chỉ giữ vài file gần nhất
            try:
                os.unlink(FILES.pop(old))
            except OSError:
                pass
    token = uuid.uuid4().hex[:12]
    p = STAGE_DIR / f"{token}.xlsx"
    p.write_bytes(body)
    FILES[token] = str(p)
    return token, str(p)


def _sheet_info(path):
    """Danh sách sheet; đánh dấu sheet trông giống báo cáo tờ khai (có cột 'Số TK', 'Mã hàng', 'Số lượng')."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        out = []
        for ws in wb.worksheets:
            ok, hdr_row = False, None
            try:
                for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
                    vals = {str(c).strip() for c in row if c is not None}
                    if {"Số TK", "Mã hàng", "Số lượng"} <= vals:
                        ok, hdr_row = True, i
                        break
            except Exception:  # noqa: BLE001 - sheet lỗi vẫn liệt kê
                pass
            out.append({"name": ws.title, "looks_like_data": ok, "header_row": hdr_row})
        default = next((s["name"] for s in out if s["name"] == "01_DATABASE" and s["looks_like_data"]), None) \
            or next((s["name"] for s in out if s["looks_like_data"]), out[0]["name"] if out else None)
        return out, default
    finally:
        wb.close()


def _stage_database(conn, path, sheet, filename, full):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError(f"Không có sheet '{sheet}' trong file")
        data = importer.parse_database_sheet(wb[sheet])
    finally:
        wb.close()
    decisions = importer.load_decisions(conn)
    groups = importer.duplicate_groups(data["parsed"], decisions)
    for p in data["parsed"]:
        del p["rawd"]  # chỉ cần để so sánh trùng
    sid = uuid.uuid4().hex[:12]
    while len(STAGES) >= 3:
        STAGES.pop(next(iter(STAGES)))
    STAGES[sid] = {"data": data, "filename": filename, "groups": {g["key"]: g for g in groups}}
    drop = [g["key"] for g in groups if g["prior"] == 1]
    return sid, groups, drop


LOT_STAGES = {}


def _tmp_xlsx(body):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(body)
        return f.name


def api_post(conn, path, q, body, ctype):
    st = db.get_settings(conn)
    if path == "/api/ob/file":
        token, p_ = _save_upload(body)
        cand, names = outbound.sheet_candidates(p_)
        return {"token": token, "candidates": cand, "sheets": names,
                "tables": [{"name": k, "title": v["title"], "default_sheet": v["sheet"]} for k, v in outbound.TABLES.items()]}
    if path == "/api/ob/import":
        d = json.loads(body or b"{}")
        p_ = FILES.get(d.get("token"))
        if not p_ or not os.path.exists(p_):
            raise ValueError("File đã hết hạn — hãy chọn lại file")
        sel = {k: v for k, v in (d.get("selection") or {}).items() if v}
        if not sel:
            raise ValueError("Chưa chọn bảng nào để nạp")
        return outbound.import_workbook(conn, p_, sel)
    m = re.fullmatch(r"/api/ob/row/(\d+)/(save|delete)", path)
    if m:
        d = json.loads(body or b"{}")
        if m.group(2) == "delete":
            outbound.delete_row(conn, int(m.group(1)))
            return {"deleted": True}
        return outbound.update_row(conn, int(m.group(1)), d.get("cells") or {})
    if path == "/api/ob/row":
        d = json.loads(body or b"{}")
        return {"id": outbound.add_row(conn, d.get("tbl"), d.get("cells") or {})}
    if path == "/api/ob/files":
        rid = int((q.get("row") or ["0"])[0])
        return outbound.add_file(conn, rid, (q.get("name") or ["file"])[0], body)
    m = re.fullmatch(r"/api/ob/files/(\d+)/delete", path)
    if m:
        outbound.delete_file(conn, int(m.group(1)))
        return {"deleted": True}
    m = re.fullmatch(r"/api/ob/goods-line/(\d+)/delete", path)
    if m:
        outbound.goods_line_delete(conn, int(m.group(1)))
        return {"deleted": True}
    if path == "/api/ob/invoice/scan":
        token = (q.get("token") or [""])[0]
        name = (q.get("name") or ["invoice.xlsx"])[0]
        if body:
            token, p_ = _save_upload(body, keep=12)
            FILE_NAMES[token] = name
        p_ = FILES.get(token)
        if not p_ or not os.path.exists(p_):
            raise ValueError("File đã hết hạn — hãy chọn lại file")
        res = invoice_scan.scan(p_, FILE_NAMES.get(token, name), (q.get("sheet") or [""])[0] or None)
        res["token"], res["file"] = token, FILE_NAMES.get(token, name)
        h = res["header"]
        res["known_ship"] = bool(h.get("polarium_ref") and conn.execute(
            "SELECT 1 FROM ob_rows WHERE tbl='source' AND key_lc=?", (h["polarium_ref"].lower(),)).fetchone())
        res["existing"] = bool(h.get("polarium_ref") and h.get("invoice_no") and conn.execute(
            "SELECT 1 FROM ob_invoices WHERE ship_lc=? AND invoice_no_lc=?", (h["polarium_ref"].lower(), h["invoice_no"].lower())).fetchone())
        return res
    if path == "/api/ob/invoice/apply":
        d = json.loads(body or b"{}")
        p_ = FILES.get(d.get("token"))
        return outbound.apply_invoice(conn, d.get("ship"), d.get("header") or {}, d.get("lines") or [], d.get("currency"), d.get("total"),
                                      d.get("mode") or "replace", p_ if p_ and os.path.exists(p_) else None, d.get("file") or FILE_NAMES.get(d.get("token")), d.get("source"))
    m = re.fullmatch(r"/api/ob/invoice/(\d+)/delete", path)
    if m:
        d = json.loads(body or b"{}")
        outbound.invoice_delete(conn, int(m.group(1)), bool(d.get("lines")))
        return {"deleted": True}
    if path == "/api/ob/goods-line":
        d = json.loads(body or b"{}")
        return {"id": outbound.goods_line_save(conn, d.get("id"), d.get("ship"), d.get("fields") or {})}
    if path == "/api/ob/goods/save":
        d = json.loads(body or b"{}")
        if not (d.get("ship") or "").strip():
            raise ValueError("Thiếu Shipment ID")
        return {"saved": outbound.goods_save(conn, d["ship"], d.get("lines") or [])}
    if path in ("/api/ob/goods/parse", "/api/ob/goods/import"):
        tmp = _tmp_xlsx(body)
        try:
            rows_, warn = outbound.goods_parse_file(tmp)
        finally:
            os.unlink(tmp)
        if path.endswith("/parse"):  # nạp cho 1 shipment đang mở: dòng không có Shipment ID thuộc shipment đó
            ship = (q.get("ship") or [""])[0].strip().lower()
            mine = [r for r in rows_ if not r["ship"] or r["ship"].lower() == ship]
            return {"lines": mine, "warnings": warn, "skipped_other_shipments": len(rows_) - len(mine)}
        res = outbound.goods_import_bulk(conn, rows_, dry_run=(q.get("dry") or ["1"])[0] == "1", mode=(q.get("mode") or ["replace"])[0],
                                         default_ship=((q.get("ship") or [""])[0].strip() or None))
        res["warnings"] = warn
        return res
    if path == "/api/boms/upload":
        tmp = _tmp_xlsx(body)
        try:
            data, sheets = boms.parse_any(tmp)
        finally:
            os.unlink(tmp)
        res = boms.apply_table(conn, data, dry_run=(q.get("dry") or ["1"])[0] == "1")
        res["sheets"] = sheets
        return res
    if path == "/api/export-material/preview":
        import openpyxl
        tmp = _tmp_xlsx(body)
        try:
            wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
            try:
                names = importer.sheets_named(wb, r"export\s*material")
                if not names:
                    raise ValueError("Không tìm thấy sheet 'Export Material' trong file")
                header, lines = importer.read_export_material(wb[names[0]])
            finally:
                wb.close()
        finally:
            os.unlink(tmp)
        if not lines:
            raise ValueError("Sheet 'Export Material' chưa có dòng NVL nào (cột Mã hàng / Qty từ dòng 11)")
        known = {r[0] for r in conn.execute("SELECT DISTINCT ma_key FROM import_lines")}
        for l in lines:
            l["in_stock_data"] = db.ma_key(l["ma_hang"]) in known
        dup = conn.execute("SELECT kind FROM events WHERE code=?", (header["so_inv"],)).fetchall() if header["so_inv"] else []
        return {"sheet": names[0], "header": header, "lines": lines, "existing_kinds": [r["kind"] for r in dup]}
    if path == "/api/fx/upload":
        tmp = _tmp_xlsx(body)
        try:
            return report.import_fx_file(conn, tmp)
        finally:
            os.unlink(tmp)
    m = re.fullmatch(r"/api/lines/(\d+)/note", path)
    if m:
        data = json.loads(body or b"{}")
        with conn:
            conn.execute("UPDATE import_lines SET ghi_chu=? WHERE id=?", ((data.get("note") or "").strip() or None, int(m.group(1))))
        return {"saved": True}
    if path in ("/api/boms", "/api/fx", "/api/fx/delete") or re.fullmatch(r"/api/boms/\d+/(save|delete)", path):
        data = json.loads(body or b"{}")
        if path == "/api/boms":
            bid = boms.create_bom(conn, data.get("ma_tp"), data if data.get("lines") else None)
            return {"id": bid}
        if path == "/api/fx":
            return report.save_fx(conn, data)
        if path == "/api/fx/delete":
            with conn:
                conn.execute("DELETE FROM fx_rates WHERE date_from=?", (importer.to_iso(data.get("date_from")),))
            return {"deleted": True}
        bid, act = re.fullmatch(r"/api/boms/(\d+)/(save|delete)", path).groups()
        if act == "delete":
            with conn:
                conn.execute("DELETE FROM boms WHERE id=?", (int(bid),))
            return {"deleted": True}
        boms.save_bom(conn, int(bid), data)
        return {"saved": True}
    if path == "/api/lots/preview":
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(body)
            tmp = f.name
        try:
            rows_ = lots_import.parse_file(tmp)
        finally:
            os.unlink(tmp)
        sid = uuid.uuid4().hex[:12]
        while len(LOT_STAGES) >= 3:
            LOT_STAGES.pop(next(iter(LOT_STAGES)))
        LOT_STAGES[sid] = rows_
        return {"stage": sid, "rows": len(rows_)}
    if path == "/api/lots/apply":
        d = json.loads(body or b"{}")
        rows_ = LOT_STAGES.get(d.get("stage"))
        if rows_ is None:
            raise ValueError("Phiên xem trước đã hết hạn — hãy tải lại file")
        res = lots_import.apply(conn, rows_, d.get("kind", "XK"), bool(d.get("replace")), bool(d.get("frozen", True)),
                                dry_run=bool(d.get("dry", True)), settings=st)
        if not d.get("dry", True):
            del LOT_STAGES[d["stage"]]
        return res
    if path == "/api/import/file":
        token, p_ = _save_upload(body)
        sheets, default = _sheet_info(p_)
        return {"token": token, "sheets": sheets, "default": default}
    if path == "/api/import/database/preview":
        full = (q.get("full") or ["1"])[0] == "1"
        path_ = FILES.get((q.get("token") or [""])[0])
        if not path_ or not os.path.exists(path_):
            raise ValueError("File đã hết hạn — hãy chọn lại file")
        sid, groups, drop = _stage_database(conn, path_, (q.get("sheet") or [""])[0], (q.get("filename") or [""])[0], full)
        diff = importer.apply_database_import(conn, STAGES[sid]["data"], full_sync=full, dedupe_keys=drop, dry_run=True)
        return {"stage": sid, "groups": groups, "drop": drop, "diff": diff}
    if path in ("/api/import/database/diff", "/api/import/database/commit"):
        d = json.loads(body or b"{}")
        s = STAGES.get(d.get("stage"))
        if not s:
            raise ValueError("Phiên xem trước đã hết hạn — hãy tải lại file")
        drop = [k for k in d.get("drop", []) if k in s["groups"]]
        full = bool(d.get("full", True))
        if path.endswith("/diff"):
            return importer.apply_database_import(conn, s["data"], full_sync=full, dedupe_keys=drop, dry_run=True)
        res = importer.apply_database_import(conn, s["data"], s["filename"], full, drop)
        with conn:
            for k, g in s["groups"].items():
                conn.execute("INSERT INTO dup_decisions(key,n_file,keep) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET "
                             "n_file=excluded.n_file, keep=excluded.keep, at=datetime('now','localtime')",
                             (k, g["n"], 1 if k in drop else g["n"]))
        del STAGES[d["stage"]]
        return res
    if path == "/api/import/legacy" or path == "/api/events/from-file":
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(body)
            tmp = f.name
        try:
            if path == "/api/import/legacy":
                return importer.import_legacy_workbook(conn, tmp, (q.get("filename") or [""])[0])
            items = importer.import_item_list_workbook(conn, tmp)
            kind = (q.get("kind") or ["XK"])[0]
            code = (q.get("code") or [""])[0].strip()
            if not code:
                raise ValueError("Thiếu mã lô / số INV")
            hdr = {k: (q.get(k) or [None])[0] for k in ("so_tk_xuat", "ngay_tk", "so_inv", "ngay_inv", "ma_tp", "ten_hang", "hs6", "qty_lo", "dvt_lo", "fob_usd", "nhap_tu", "nhap_den")}
            return {"id": importer.create_event(conn, kind, code, hdr, items, multiply_by_lot=(kind == "CO"))}
        finally:
            os.unlink(tmp)
    data = json.loads(body or b"{}")
    if path == "/api/events":
        kind = data.get("kind", "XK")
        items = [{"ma_hang": l["ma_hang"], "mo_ta": l.get("mo_ta"), "qty": float(l["qty"]), "dvt": l.get("dvt")}
                 for l in data.get("lines", []) if l.get("ma_hang")]
        if data.get("bom_id"):  # lô C/O lấy định mức từ BOM đã lưu: nhu cầu = định mức × số lượng lô
            if kind != "CO":
                raise ValueError("BOM chỉ dùng cho lô C/O")
            b = conn.execute("SELECT * FROM boms WHERE id=?", (int(data["bom_id"]),)).fetchone()
            if b is None:
                raise ValueError("Không tìm thấy BOM")
            items = boms.lines_for_event(conn, b["id"])
            data = {**data, "ma_tp": data.get("ma_tp") or b["ma_tp"], "ten_hang": data.get("ten_hang") or b["ten_hang"],
                    "hs6": data.get("hs6") or b["hs6"], "dvt_lo": data.get("dvt_lo") or b["dvt"]}
            if not items:
                raise ValueError("BOM chưa có dòng NVL nào")
        if not data.get("code") or not items:
            raise ValueError("Cần mã lô và ít nhất 1 dòng NVL")
        return {"id": importer.create_event(conn, kind, data["code"], data, items, multiply_by_lot=(kind == "CO"))}
    if path in ("/api/rebuild/preview", "/api/rebuild/commit"):
        return engine.rebuild(conn, st, dry_run=path.endswith("preview"))
    m = re.fullmatch(r"/api/events/(\d+)/(lock|unlock|delete|update|freeze)", path)
    if m:
        eid, act = int(m.group(1)), m.group(2)
        if act == "freeze":
            engine.set_frozen(conn, eid, bool(data.get("frozen", True)))
            return {"frozen": bool(data.get("frozen", True))}
        if act == "lock":
            plan = engine.lock_event(conn, eid, st, allow_shortfall=bool(data.get("allow_shortfall")),
                                     frozen=bool(data.get("frozen")))
            # lô vừa khoá có làm đổi thứ tự ưu tiên của các lô khác không? (vd thêm lô C/O sau khi đã có lô XK)
            pending = len(engine.rebuild(conn, st, dry_run=True)["changed"])
            return {"locked": True, "short": sum(1 for p in plan if p["shortfall"] > 1e-9), "rebuild_pending": pending}
        if act == "unlock":
            engine.unlock_event(conn, eid)
            return {"unlocked": True}
        if act == "delete":
            with conn:
                conn.execute("DELETE FROM events WHERE id=?", (eid,))
            return {"deleted": True}
        fields = ("so_tk_xuat", "ngay_tk", "so_inv", "ngay_inv", "ma_tp", "ten_hang", "hs6", "qty_lo", "dvt_lo", "fob_usd", "co_so", "co_ngay", "note", "nhap_tu", "nhap_den", "tieu_chi")
        sets = {k: (importer.to_iso(data[k]) if k in ("ngay_tk", "ngay_inv", "co_ngay", "nhap_tu", "nhap_den") else data[k])
                for k in fields if k in data}
        for k in ("nhap_tu", "nhap_den"):
            if sets.get(k) and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", sets[k]):
                raise ValueError("Ngày nhập không hợp lệ — dùng dd/mm/yyyy hoặc yyyy-mm-dd")
        locked = conn.execute("SELECT status FROM events WHERE id=?", (eid,)).fetchone()
        if locked and locked["status"] == "LOCKED" and ({"nhap_tu", "nhap_den"} & sets.keys()):
            sets.pop("nhap_tu", None)  # lô đã khoá: khoảng ngày nhập không đổi (huỷ khoá trước nếu muốn đổi)
            sets.pop("nhap_den", None)
        if sets:
            with conn:
                conn.execute(f"UPDATE events SET {','.join(k + '=?' for k in sets)} WHERE id=?", [*sets.values(), eid])
        return {"updated": True}
    if path == "/api/settings":
        with conn:
            for k, v in data.items():
                if k in db.DEFAULT_SETTINGS:
                    conn.execute("UPDATE settings SET value=? WHERE key=?", (str(v), k))
        return db.get_settings(conn)
    raise KeyError(path)


class Handler(BaseHTTPRequestHandler):
    server_version = "NVL/1.0"

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _run(self, method):
        u = urlparse(self.path)
        path, q = u.path, parse_qs(u.query)
        try:
            if method == "GET" and (path == "/" or path == "/index.html"):
                return self._send(200, (STATIC / "index.html").read_bytes(), "text/html; charset=utf-8")
            conn = db.connect()
            try:
                if method == "GET" and path.startswith("/files/"):
                    mf = re.fullmatch(r"/files/(\d+)(?:/.*)?", path)
                    if not mf:
                        raise KeyError(path)
                    meta, p_ = outbound.get_file(conn, int(mf.group(1)))
                    ext = Path(meta["name"]).suffix.lower()
                    from urllib.parse import quote
                    inline = ext in outbound.INLINE_EXT
                    fname = quote(meta["name"])
                    return self._send(200, p_.read_bytes(), meta["mime"] if inline else "application/octet-stream",
                                      {"Content-Disposition": f"{'inline' if inline else 'attachment'}; filename*=UTF-8''{fname}",
                                       "X-Content-Type-Options": "nosniff"})
                if method == "GET" and path.startswith("/export/"):
                    st = db.get_settings(conn)
                    mm = re.fullmatch(r"/export/(lot|kind)/(\w+)\.xlsx", path)
                    if mm and mm.group(1) == "lot":
                        data, name = exporter.build_lot_xlsx(conn, int(mm.group(2)), st), f"lo-{mm.group(2)}.xlsx"
                    elif mm:
                        data, name = exporter.build_kind_ledger_xlsx(conn, mm.group(2)), f"so-nvl-{mm.group(2)}.xlsx"
                    elif path == "/export/ob-form.xlsx":
                        ship = (q.get("ship") or [""])[0].strip()
                        inv = (q.get("inv") or [""])[0].strip() or None
                        sheets = [s for s in ((q.get("sheets") or [""])[0].split("|")) if s]
                        if not sheets:
                            raise ValueError("Chọn ít nhất một biểu mẫu")
                        wb_, _errs = outbound_forms.build(conn, ship, inv, sheets)
                        data = outbound_forms.to_bytes(wb_)
                        name = (f"{ship or inv}-{'+'.join(sheets) if len(sheets) < 3 else 'bieu-mau'}.xlsx").replace(" ", "_")
                    elif path == "/export/ob-full.xlsx":
                        a = _grid_args(q)
                        sh_ = (q.get("ship") or [""])[0].strip() or None
                        if (q.get("row") or [""])[0].isdigit():
                            sh_ = outbound.get_row(conn, int(q["row"][0])).get("ship") or sh_
                        data, name = outbound.full_export_xlsx(conn, a["filters"], a["q"], "list" if (q.get("tbl") or [""])[0] == "list" else "source", sh_), (f"{sh_}-day-du.xlsx" if sh_ else "outbound-day-du-shipment.xlsx")
                    elif path == "/export/ob-goods.xlsx":
                        a = _grid_args(q)
                        data, name = outbound.goods_export_xlsx(conn, a["filters"], a["q"], a["sort"], a["dir"]), "dong-hang-invoice.xlsx"
                    elif path == "/export/ob-report.xlsx":
                        sts = [s for s in (q.get("statuses") or [""])[0].split("|") if s]
                        data = outbound.report_xlsx(conn, sts or None, (q.get("from") or [""])[0] or None, (q.get("to") or [""])[0] or None)
                        name = "bao-cao-outbound.xlsx"
                    elif re.fullmatch(r"/export/ob/\w+\.xlsx", path):
                        a = _grid_args(q)
                        tbl_ = path.split("/")[-1][:-5]
                        data, name = outbound.export_xlsx(conn, tbl_, a["filters"], a["q"], a["sort"], a["dir"]), f"outbound-{tbl_}.xlsx"
                    elif path == "/export/bom-template.xlsx":
                        data, name = boms.template_xlsx(), "mau-bom.xlsx"
                    elif path == "/export/lots-template.xlsx":
                        data, name = lots_import.template_xlsx(), "mau-nap-lo-cu.xlsx"
                    elif path == "/export/grid.xlsx":
                        a = _grid_args(q)
                        data = grid.export_xlsx(conn, a["filters"], a["q"], a["sort"], a["dir"], a["cols"])
                        name = "to-khai-nhap-loc.xlsx"
                    elif path == "/export/stock.xlsx":
                        data, name = exporter.build_stock_xlsx(conn), "ton-kho-nvl.xlsx"
                    elif path == "/export/ledger.xlsx":
                        data, name = exporter.build_ledger_xlsx(conn), "nhat-ky-tru-lui.xlsx"
                    else:
                        m = re.fullmatch(r"/export/rvc/(\d+)\.xlsx", path)
                        if not m:
                            raise KeyError(path)
                        data = exporter.build_rvc_xlsx(conn, int(m.group(1)), st)
                        name = f"RVC-{m.group(1)}.xlsx"
                    return self._send(200, data, XLSX, {"Content-Disposition": f'attachment; filename="{name}"'})
                if method == "GET":
                    return self._send(200, api_get(conn, path, q))
                n = int(self.headers.get("Content-Length") or 0)
                body = self.rfile.read(n) if n else b""
                with LOCK:
                    return self._send(200, api_post(conn, path, q, body, self.headers.get("Content-Type", "")))
            finally:
                conn.close()
        except KeyError as e:
            self._send(404, {"error": f"Không tìm thấy: {e}"})
        except (ValueError, json.JSONDecodeError) as e:
            self._send(400, {"error": str(e)})
        except Exception as e:  # noqa: BLE001 - báo lỗi về giao diện thay vì sập server
            self._send(500, {"error": f"{type(e).__name__}: {e}"})

    def do_GET(self):
        self._run("GET")

    def do_POST(self):
        self._run("POST")


def main():
    db.connect().close()
    if STAGE_DIR.exists():  # dọn các file tải lên dở dang của lần chạy trước
        for f in STAGE_DIR.glob("*.xlsx"):
            try:
                f.unlink()
            except OSError:
                pass
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)  # chỉ nghe localhost
    url = f"http://127.0.0.1:{PORT}"
    print(f"Sổ theo dõi NVL đang chạy tại {url}  (Ctrl+C để dừng)")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
