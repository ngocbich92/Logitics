"""Xuất báo cáo Excel: tồn kho, nhật ký trừ lùi, bảng kê RVC (Phụ lục VIII)."""
import io
from statistics import mean

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="BBBBBB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _sheet(wb, title, headers, rows, widths=None, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for r in rows:
        ws.append(list(r))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(i, max(12, min(40, len(str(h)) + 4)))
    return ws


def stock_rows(conn):
    return conn.execute("""
        SELECT i.id, i.so_tk, i.ngay_dk, i.loai_hinh, i.ma_hang, i.ten_hang, i.dvt, i.nuoc_xx, i.ma_hs, i.don_gia_tt, i.qty,
               COALESCE((SELECT SUM(a.qty) FROM allocations a WHERE a.import_line_id=i.id),0) AS used
        FROM import_lines i ORDER BY i.ngay_dk DESC, i.seq""").fetchall()


def build_stock_xlsx(conn):
    wb = Workbook()
    rows = []
    for r in stock_rows(conn):
        remain = max(r["qty"] - r["used"], 0)
        rows.append([r["so_tk"], r["ngay_dk"], r["loai_hinh"], r["ma_hang"], r["ten_hang"], r["dvt"], r["nuoc_xx"], r["ma_hs"],
                     r["don_gia_tt"], r["qty"], r["used"], remain, remain * (r["don_gia_tt"] or 0),
                     "Hết hàng" if remain <= 1e-9 else "Còn hàng"])
    _sheet(wb, "Tồn kho theo tờ khai",
           ["Số TK", "Ngày ĐK", "Loại hình", "Mã hàng", "Tên hàng", "ĐVT", "Nước XX", "Mã HS", "Đơn giá tính thuế",
            "SL nhập", "Luỹ kế xuất", "SL tồn", "Giá trị tồn (VNĐ)", "Trạng thái"], rows, {5: 50}, first=True)
    summ = conn.execute("""
        SELECT i.ma_key, MIN(i.ten_hang) ten, MIN(i.dvt) dvt, SUM(i.qty) nhap,
               COALESCE(SUM((SELECT SUM(a.qty) FROM allocations a WHERE a.import_line_id=i.id)),0) xuat
        FROM import_lines i GROUP BY i.ma_key ORDER BY i.ma_key""").fetchall()
    _sheet(wb, "Tồn kho theo mã hàng", ["Mã hàng", "Tên hàng", "ĐVT", "Tổng nhập", "Tổng xuất", "Tồn"],
           [[s["ma_key"], s["ten"], s["dvt"], s["nhap"], s["xuat"], s["nhap"] - s["xuat"]] for s in summ], {2: 50})
    return _bytes(wb)


def build_ledger_xlsx(conn):
    """Nhật ký trừ lùi + ma trận (dòng nhập × lô) — thay cho các cột lô hàng của sheet 02_CO / 03."""
    wb = Workbook()
    ev = conn.execute("SELECT id,kind,code FROM events WHERE status='LOCKED' ORDER BY id").fetchall()
    log = conn.execute("""
        SELECT e.kind, e.code, e.so_tk_xuat, e.ngay_tk, el.ma_hang, a.qty, i.so_tk, i.ngay_dk, i.loai_hinh, i.nuoc_xx
        FROM allocations a JOIN events e ON e.id=a.event_id JOIN event_lines el ON el.id=a.event_line_id
        JOIN import_lines i ON i.id=a.import_line_id ORDER BY e.id, el.id, a.id""").fetchall()
    _sheet(wb, "Nhật ký trừ lùi",
           ["Loại", "Lô / INV", "TK xuất", "Ngày TK xuất", "Mã hàng", "SL trừ", "TK nhập bị trừ", "Ngày TK nhập", "Loại hình", "Nước XX"],
           [list(r) for r in log], first=True)
    lines = conn.execute("SELECT id, so_tk, ngay_dk, loai_hinh, ma_hang, qty FROM import_lines ORDER BY ngay_dk DESC, seq").fetchall()
    cell = {}
    for r in conn.execute("SELECT event_id, import_line_id, SUM(qty) q FROM allocations GROUP BY event_id, import_line_id"):
        cell[(r["import_line_id"], r["event_id"])] = r["q"]
    heads = ["Số TK", "Ngày ĐK", "Loại hình", "Mã hàng", "SL nhập"] + [f'{e["kind"]}-{e["code"]}' for e in ev] + ["Luỹ kế xuất", "Tồn"]
    rows = []
    for l in lines:
        per = [cell.get((l["id"], e["id"]), 0) for e in ev]
        rows.append([l["so_tk"], l["ngay_dk"], l["loai_hinh"], l["ma_hang"], l["qty"], *per, sum(per), l["qty"] - sum(per)])
    _sheet(wb, "Ma trận", heads, rows)
    return _bytes(wb)


LEDGER_COLS = [  # (tiêu đề, field, độ rộng)
    ("Lô / INV", "lo", 14), ("Trạng thái lô", "trang_thai", 18), ("Ngày TK xuất", "ngay_tk_xuat", 12), ("TK xuất", "so_tk_xuat", 15), ("Số C/O", "so_co", 12), ("Ngày C/O", "ngay_co", 11),
    ("Tờ khai nhập", "so_tk", 15), ("Ngày tờ khai", "ngay_tk_nhap", 12), ("Mã hàng", "ma_hang", 14), ("Tên hàng", "ten_hang", 50),
    ("ĐVT", "dvt", 7), ("Nước xuất xứ", "nuoc_xx", 14), ("Mã HS", "ma_hs", 11), ("Số lượng", "qty", 11),
    ("Đơn giá CIF (USD)", "don_gia_cif", 14), ("Tỷ giá USD/VNĐ", "ty_gia_usd", 12), ("Trị giá (USD)", "tri_gia_usd", 13)]


def build_kind_ledger_xlsx(conn, kind):
    import report
    wb = Workbook()
    rows = report.ledger(conn, kind)
    _sheet(wb, f"NVL {report.KIND_LABEL[kind]}".replace("/", "-")[:31], [c[0] for c in LEDGER_COLS],
           [[r.get(c[1]) for c in LEDGER_COLS] for r in rows], {i: c[2] for i, c in enumerate(LEDGER_COLS, start=1)}, first=True)
    return _bytes(wb)


LOT_COLS = [("STT", "stt", 6), ("Mã hàng", "ma_hang", 14), ("Nhu cầu", "need", 10), ("Tờ khai nhập", "so_tk", 15), ("Ngày tờ khai", "ngay_tk_nhap", 12),
            ("Tên hàng", "ten_hang", 50), ("Nước xuất xứ", "nuoc_xx", 14), ("Mã HS", "ma_hs", 11), ("SL trừ", "qty", 11),
            ("Đơn giá CIF (USD)", "don_gia_cif", 14), ("Tỷ giá USD/VNĐ", "ty_gia_usd", 12), ("Trị giá (USD)", "tri_gia_usd", 13)]


def build_lot_xlsx(conn, event_id, settings):
    import report
    ev = conn.execute("SELECT code FROM events WHERE id=?", (event_id,)).fetchone()
    rows = report.lot_rows(conn, event_id, settings)
    wb = Workbook()
    data = [[("THIẾU" if r.get("is_short") and c[1] in ("so_tk",) else r.get(c[1])) for c in LOT_COLS] for r in rows]
    _sheet(wb, (ev["code"] or "Lo")[:31], [c[0] for c in LOT_COLS], data, {i: c[2] for i, c in enumerate(LOT_COLS, start=1)}, first=True)
    return _bytes(wb)


def build_rvc_xlsx(conn, event_id, settings):
    """Phụ lục VIII — điền vào đúng mẫu sheet RVC của sổ Excel (xem rvc_form.py)."""
    import rvc_form
    wb, _ = rvc_form.build(conn, event_id, settings)
    return rvc_form.to_bytes(wb)

def _bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
