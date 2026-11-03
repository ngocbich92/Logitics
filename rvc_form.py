"""Form RVC (Phụ lục VIII) — điền vào ĐÚNG mẫu sheet 'RVC' trong sổ Excel của người dùng (templates/rvc_template.xlsx).

Mẫu giữ nguyên: font, cỡ chữ, ô gộp, đường viền, độ rộng cột, chiều cao dòng, lề in. Vùng dữ liệu của mẫu là dòng 17–118
(102 dòng); nếu nhiều/ít dòng hơn thì chèn/xoá dòng và dịch phần cuối (dòng cộng, kết luận, chữ ký) cho khớp.
"""
import copy
import datetime as dt
import html
import io
import re
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

import report

TEMPLATE = Path(__file__).parent / "templates" / "rvc_template.xlsx"
FIRST, N_TPL = 17, 102            # dòng dữ liệu đầu tiên của mẫu, số dòng dữ liệu của mẫu (17..118)
TOTAL_ROW = FIRST + N_TPL         # 119: dòng cộng (có viền, trống trong mẫu)
DATE_FMT = "dd/mm/yyyy"


def _date(v):
    if not v:
        return None
    try:
        return dt.date.fromisoformat(str(v)[:10])
    except ValueError:
        return v


def _name(x):
    nm = x.get("mo_ta") or x.get("ten_hang") or ""
    return nm if nm.startswith(x["ma_hang"] + "#&") else f"{x['ma_hang']}#&{nm}"


def _resize(ws, n):
    """Chèn / xoá dòng dữ liệu để vùng dữ liệu có đúng n dòng. Trả về độ lệch dòng (delta)."""
    delta = n - N_TPL
    if delta == 0:
        return 0
    tail_merges = [m for m in ws.merged_cells.ranges if m.min_row >= TOTAL_ROW]
    spans = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in tail_merges]
    for m in tail_merges:
        ws.unmerge_cells(str(m))
    heights = {r: ws.row_dimensions[r].height for r in range(TOTAL_ROW, ws.max_row + 1)}
    if delta > 0:
        ws.insert_rows(TOTAL_ROW, amount=delta)
        model = FIRST + N_TPL - 1
        for r in range(TOTAL_ROW, TOTAL_ROW + delta):
            for c in range(1, 14):
                src, dst = ws.cell(model, c), ws.cell(r, c)
                dst._style = copy.copy(src._style)
    else:
        ws.delete_rows(FIRST + n, -delta)
    for r in list(ws.row_dimensions.keys()):
        if r >= TOTAL_ROW - max(delta, 0):
            ws.row_dimensions[r].height = None
    for r, h in heights.items():
        ws.row_dimensions[r + delta].height = h
    for (r1, c1, r2, c2) in spans:
        ws.merge_cells(start_row=r1 + delta, start_column=c1, end_row=r2 + delta, end_column=c2)
    return delta


def build(conn, event_id, settings):
    """-> (openpyxl Workbook đã điền, dữ liệu rvc)"""
    d = report.rvc(conn, event_id, settings)
    ev, co, rows = d["event"], d["company"], d["rows"]
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["RVC"]
    n = max(len(rows), 1)
    # xoá dữ liệu mẫu trong vùng dữ liệu (kể cả số nhập tay như H17=41.76)
    for r in range(FIRST, FIRST + N_TPL):
        for c in range(1, 14):
            ws.cell(r, c).value = None
    delta = _resize(ws, n)

    # ---- phần đầu ----
    tieu_chi = ev.get("tieu_chi") or settings.get("tieu_chi") or ""
    ws["A5"] = f"Tên thương nhân: {co['name']}"
    ws["A6"] = f"Mã số thuế/ID tax: {co['mst']}"
    ws["K5"] = tieu_chi
    ws["I6"] = f"Tên hàng hóa: {ev.get('ten_hang') or ''}"
    ws["A7"] = f"Tờ khai hải quan xuất khẩu số: {ev.get('so_tk_xuat') or ''}"
    ws["D7"] = _date(ev.get("ngay_tk"))
    ws["K7"] = ev.get("hs6")
    ws["B8"] = ev.get("so_inv")
    ws["D8"] = _date(ev.get("ngay_inv"))
    ws["K8"] = ev.get("ma_tp")
    ws["K9"] = ev.get("qty_lo")
    ws["L9"] = ev.get("dvt_lo")
    ws["K10"] = ev.get("fob_usd")
    for ref in ("D7", "D8"):
        if isinstance(ws[ref].value, dt.date):
            ws[ref].number_format = DATE_FMT

    # ---- bảng nguyên liệu ----
    last = None
    for i, x in enumerate(rows):
        r = FIRST + i
        new = x["stt"] != last
        last = x["stt"]
        if x.get("is_short"):
            ws.cell(r, 1, x["stt"] if new else None)
            ws.cell(r, 2, _name(x))
            ws.cell(r, 4, x["dvt"])
            ws.cell(r, 5, x["dinh_muc"] if new else None)
            ws.cell(r, 9, f"THIẾU TỒN {x['qty']:g}")
            continue
        vals = [x["stt"] if new else None, _name(x), (x["ma_hs"] or "")[:6] or None, x["dvt"], x["dinh_muc"] if new else None,
                x["don_gia_cif"], None, None, x["nuoc_xx"], x["so_tk"], _date(x["ngay_tk_nhap"]), None, None]
        for c, v in enumerate(vals, start=1):
            ws.cell(r, c).value = v
        if len(_name(x)) > 75:
            ws.row_dimensions[r].height = 55
        elif not ws.row_dimensions[r].height:
            ws.row_dimensions[r].height = 40.95
        if ws.cell(r, 11).value and isinstance(ws.cell(r, 11).value, dt.date):
            ws.cell(r, 11).number_format = "[$-1010000]d/m/yyyy;@"

    # ---- dòng cộng, kết luận, chữ ký (đã dịch theo delta) ----
    t = TOTAL_ROW + delta
    ws.cell(t, 2, "Cộng chi phí nguyên liệu (USD)")
    # G, H (có / không có xuất xứ FTA) để trống cho người dùng điền tay; dòng cộng là công thức SUM nên tự cập nhật khi điền
    ws.cell(t, 7, f"=SUM(G{FIRST}:G{t - 1})")
    ws.cell(t, 8, f"=SUM(H{FIRST}:H{t - 1})")
    ws.cell(122 + delta, 2, f"Kết luận: Hàng hóa đáp ứng tiêu chí {tieu_chi}")
    today = dt.date.today()
    ws.cell(127 + delta, 8, f"{co['place']}, ngày {today.day} tháng  {today.month} năm {today.year}")
    return wb, d


def to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------- xem trước HTML (từ chính bảng đã điền) ---------------
def _fmt(c):
    v = c.value
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        v = v.date() if isinstance(v, dt.datetime) else v
        return f"{v.day}/{v.month}/{v.year}"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        f = c.number_format or ""
        if "." in f.split(";")[0]:
            dec = sum(1 for ch in f.split(";")[0].split(".")[1] if ch == "0")  # số chữ số thập phân theo định dạng ô
            return f"{v:,.{dec}f}"
        return f"{v:,.4f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
    return str(v)


def _hex(color):
    """Màu openpyxl -> #rrggbb (chỉ nhận màu RGB / một số màu index phổ biến); khác -> None."""
    if color is None:
        return None
    try:
        if color.type == "rgb" and isinstance(color.rgb, str) and len(color.rgb) == 8 and color.rgb != "00000000":
            return "#" + color.rgb[2:].lower()
        if color.type == "indexed" and color.indexed == 10:
            return "#ff0000"
    except AttributeError:
        pass
    return None


def print_bounds(ws, default_cols=13):
    """(min_col, min_row, max_col, max_row) theo vùng in của sheet; nếu không có thì toàn bộ vùng có dữ liệu."""
    pa = ws.print_area
    if pa:
        m = re.search(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)", pa if isinstance(pa, str) else str(pa))
        if m:
            return (column_index_from_string(m.group(1)), int(m.group(2)), column_index_from_string(m.group(3)), int(m.group(4)))
    return (1, 1, min(ws.max_column, default_cols) if default_cols else ws.max_column, ws.max_row)


def to_html(ws, ncols=13, bounds=None):
    min_c, min_r, max_c, max_r = bounds or (1, 1, ncols, ws.max_row)
    merges = {}
    covered = set()
    for m in ws.merged_cells.ranges:
        merges[(m.min_row, m.min_col)] = (m.max_row - m.min_row + 1, m.max_col - m.min_col + 1)
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                if (r, c) != (m.min_row, m.min_col):
                    covered.add((r, c))
    widths = []
    for c in range(min_c, max_c + 1):
        w = ws.column_dimensions[get_column_letter(c)].width or 8.43
        widths.append(int(w * 7 + 5))
    out = [f'<table style="border-collapse:collapse;table-layout:fixed;width:{sum(widths)}px;background:#fff;color:#000">',
           "<colgroup>" + "".join(f'<col style="width:{w}px">' for w in widths) + "</colgroup>"]
    def border(side):
        """Đường viền theo đúng kiểu và màu của ô gốc (vd viền đỏ)."""
        if not side or not side.style:
            return "none"
        w = {"thin": 1, "hair": 1, "medium": 2, "thick": 3, "double": 3}.get(side.style, 1)
        line = {"dashed": "dashed", "dotted": "dotted", "double": "double", "mediumDashed": "dashed", "hair": "dotted"}.get(side.style, "solid")
        return f"{w}px {line} {_hex(side.color) or '#000'}"
    for r in range(min_r, max_r + 1):
        h = ws.row_dimensions[r].height
        out.append(f'<tr style="height:{int((h or 15) * 1.34)}px">')
        for c in range(min_c, max_c + 1):
            if (r, c) in covered:
                continue
            cell = ws.cell(r, c)
            rs, cs = merges.get((r, c), (1, 1))
            f, a, b = cell.font, cell.alignment, cell.border
            if (rs, cs) == (1, 1) and isinstance(cell.value, str) and not a.wrap_text and a.horizontal in (None, "general", "left"):
                # văn bản không xuống dòng tràn sang các ô trống bên phải (giống Excel)
                while c + cs <= max_c and (r, c + cs) not in covered and ws.cell(r, c + cs).value is None and (r, c + cs) not in merges:
                    covered.add((r, c + cs))
                    cs += 1
            st = [f"font-family:'{f.name or 'Times New Roman'}',serif", f"font-size:{(f.sz or 11) * 1.33:.0f}px"]
            if f.b:
                st.append("font-weight:bold")
            if f.i:
                st.append("font-style:italic")
            st.append(f"text-align:{a.horizontal if a.horizontal in ('left', 'center', 'right') else 'left'}")
            st.append(f"vertical-align:{ {'center': 'middle', 'bottom': 'bottom'}.get(a.vertical, 'top') }")
            st.append("white-space:normal;word-break:break-word" if a.wrap_text else "white-space:nowrap;overflow:hidden")
            st.append(f"border-left:{border(b.left)};border-right:{border(b.right)};border-top:{border(b.top)};border-bottom:{border(b.bottom)}")
            bg = _hex(cell.fill.fgColor) if cell.fill and cell.fill.fill_type == "solid" else None
            if bg and bg not in ("#ffffff", "#000000"):
                st.append(f"background:{bg}")
            fc = _hex(f.color)
            if fc and fc != "#000000":
                st.append(f"color:{fc}")
            st.append("padding:1px 3px;overflow:hidden")
            val = html.escape(_fmt(cell)).replace("\n", "<br>")
            out.append(f'<td rowspan="{rs}" colspan="{cs}" style="{";".join(st)}">{val}</td>')
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)