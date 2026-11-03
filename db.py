"""SQLite storage for the NVL (nguyên vật liệu) tracking system."""
import os
import re
import sqlite3
import unicodedata
from pathlib import Path

import pgcompat

DB_PATH = Path(os.environ.get("NVL_DB") or Path(__file__).parent / "data" / "nvl.db")   # NVL_DB: đường dẫn DB khác (dùng khi thử nghiệm)

SCHEMA = """
CREATE TABLE IF NOT EXISTS import_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    seq INTEGER NOT NULL,              -- thứ tự dòng trong nguồn (tie-break khi cùng ngày)
    so_tk TEXT NOT NULL,
    ngay_dk TEXT,                      -- ISO yyyy-mm-dd
    loai_hinh TEXT,                    -- E11 / E15 / E13 / B11 ...
    doi_tac TEXT,
    van_don TEXT,
    so_hd TEXT,
    ngay_hd TEXT,
    ccy TEXT,                          -- Ng.Tệ hóa đơn
    ty_gia REAL,                       -- Tỷ giá VNĐ
    trang_thai TEXT,
    ma_hs TEXT,
    ma_hang TEXT NOT NULL,
    ma_key TEXT NOT NULL,              -- mã hàng chuẩn hoá để so khớp BOM
    ten_hang TEXT,
    dvt TEXT,
    qty REAL NOT NULL,
    don_gia_hd REAL,
    don_gia_tt REAL,                   -- đơn giá tính thuế (VNĐ)
    tri_gia_hd REAL,
    nuoc_xx TEXT,
    dedupe TEXT NOT NULL UNIQUE
);
CREATE INDEX IF NOT EXISTS ix_imp_key ON import_lines(ma_key, ngay_dk);

CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL CHECK (kind IN ('CO','XK','MDK','TH')),
    code TEXT NOT NULL,                -- VD VNOUT00732 / PXK0044
    so_tk_xuat TEXT,
    ngay_tk TEXT,
    so_inv TEXT,
    ngay_inv TEXT,
    ma_tp TEXT,                        -- mã thành phẩm
    ten_hang TEXT,
    hs6 TEXT,
    qty_lo REAL DEFAULT 1,
    dvt_lo TEXT,
    fob_usd REAL,
    co_so TEXT,
    co_ngay TEXT,
    status TEXT NOT NULL DEFAULT 'DRAFT' CHECK (status IN ('DRAFT','LOCKED')),
    note TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    locked_at TEXT
);

CREATE TABLE IF NOT EXISTS event_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    ma_hang TEXT NOT NULL,
    ma_key TEXT NOT NULL,
    mo_ta TEXT,
    dvt TEXT,
    dinh_muc REAL,                     -- Qty used in FG (CO)
    need REAL NOT NULL                 -- nhu cầu cần trừ
);
CREATE INDEX IF NOT EXISTS ix_el_event ON event_lines(event_id);

CREATE TABLE IF NOT EXISTS allocations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    event_line_id INTEGER NOT NULL REFERENCES event_lines(id) ON DELETE CASCADE,
    import_line_id INTEGER NOT NULL REFERENCES import_lines(id),
    qty REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_al_imp ON allocations(import_line_id);
CREATE INDEX IF NOT EXISTS ix_al_ev ON allocations(event_id);

CREATE TABLE IF NOT EXISTS boms (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_tp TEXT NOT NULL UNIQUE,        -- mã thành phẩm
    ten_hang TEXT, hs6 TEXT, dvt TEXT, note TEXT,
    updated_at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE TABLE IF NOT EXISTS bom_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bom_id INTEGER NOT NULL REFERENCES boms(id) ON DELETE CASCADE,
    ma_hang TEXT NOT NULL, ma_key TEXT NOT NULL, mo_ta TEXT, dvt TEXT,
    dinh_muc REAL NOT NULL            -- Qty used in FG: NVL cho 1 đơn vị thành phẩm
);
CREATE INDEX IF NOT EXISTS ix_bl_bom ON bom_lines(bom_id);
CREATE TABLE IF NOT EXISTS fx_rates (  -- tỷ giá hải quan USD/VNĐ theo tuần, nhập tay / từ website hải quan (ưu tiên hơn tỷ giá suy ra từ dữ liệu)
    date_from TEXT PRIMARY KEY, date_to TEXT NOT NULL, usd REAL NOT NULL, note TEXT
);
CREATE TABLE IF NOT EXISTS ob_tables (      -- hệ thống Outbound shipment: các bảng dữ liệu lấy từ file 'Database outbound shipping plan'
    name TEXT PRIMARY KEY, title TEXT, sheet TEXT, key_col TEXT,
    cols TEXT,                              -- JSON [{letter, name, type, tag}]
    updated_at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE TABLE IF NOT EXISTS ob_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tbl TEXT NOT NULL REFERENCES ob_tables(name) ON DELETE CASCADE,
    pos INTEGER NOT NULL,                   -- thứ tự dòng
    key TEXT, key_lc TEXT,                  -- khoá (vd Shipment ID) và dạng chữ thường để tra như VLOOKUP
    cells TEXT NOT NULL                     -- JSON {chữ cột: giá trị}; ngày lưu dạng ISO
);
CREATE TABLE IF NOT EXISTS ob_files (       -- file đính kèm (hình ảnh / pdf bằng chứng giao nhận...) của một dòng dữ liệu Outbound
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    row_id INTEGER REFERENCES ob_rows(id) ON DELETE CASCADE,
    name TEXT NOT NULL, stored TEXT NOT NULL, size INTEGER, mime TEXT,
    at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS ix_obf_row ON ob_files(row_id);
CREATE TABLE IF NOT EXISTS ob_goods (       -- Information of goods: các dòng hàng của một shipment (nạp từ file Excel khác)
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ship_key TEXT NOT NULL, ship_lc TEXT NOT NULL, pos INTEGER NOT NULL,
    ma_hang TEXT, ten_hang TEXT, qty REAL, dvt TEXT, don_gia REAL, ccy TEXT DEFAULT 'USD',
    xuat_xu TEXT, amount REAL, ter TEXT, net_w REAL, gross_w REAL
);
CREATE INDEX IF NOT EXISTS ix_obg_ship ON ob_goods(ship_lc, pos);
CREATE TABLE IF NOT EXISTS ob_invoices (    -- thông tin Commercial Invoice đã quét từ file Excel / PDF
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ship_key TEXT NOT NULL, ship_lc TEXT NOT NULL,
    invoice_no TEXT, invoice_no_lc TEXT, invoice_date TEXT, customer_po TEXT, so_to TEXT, so_to_lc TEXT, polarium_ref TEXT,
    incoterms TEXT, mode TEXT, carrier TEXT, shipped_on TEXT, hs_code TEXT, term_of_payment TEXT, remarks TEXT,
    sold_by TEXT, sold_to TEXT, shipper TEXT, deliver_to TEXT,
    ccy TEXT, total_value REAL, n_lines INTEGER, file_name TEXT, source TEXT,
    at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS ix_obi_ship ON ob_invoices(ship_lc);
CREATE INDEX IF NOT EXISTS ix_obi_inv ON ob_invoices(invoice_no_lc);
CREATE INDEX IF NOT EXISTS ix_obi_so ON ob_invoices(so_to_lc);
CREATE INDEX IF NOT EXISTS ix_obr_key ON ob_rows(tbl, key_lc);
CREATE INDEX IF NOT EXISTS ix_obr_pos ON ob_rows(tbl, pos);
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);   -- danh mục cột của lưới dữ liệu
CREATE TABLE IF NOT EXISTS dup_decisions (    -- quyết định của người dùng cho nhóm dòng trùng
    key TEXT PRIMARY KEY, n_file INTEGER, keep INTEGER, at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    at TEXT DEFAULT (datetime('now','localtime')),
    filename TEXT, mode TEXT, rows_in_file INTEGER,
    added INTEGER, updated INTEGER, removed INTEGER, kept INTEGER
);
"""

DEFAULT_SETTINGS = {
    "only_cleared": "1",             # 1 = chỉ trừ vào tờ khai 'Đã thông quan'
    "company": "CÔNG TY TNHH POLARIUM VIỆT NAM",   # tiêu đề form RVC
    "mst": "0202099739",
    "place": "Hải Phòng",
    "tieu_chi": "CTSH + ACU",         # tiêu chí RVC mặc định cho lô mới
    "priority": "CO_FIRST",          # CO_FIRST = lô C/O luôn được trừ trước (như Excel) | CHRONO = theo ngày TK xuất
    "order": "LIFO",                 # LIFO = nhập mới nhất trừ trước (như sổ Excel hiện tại)
    "types_co": "E11,E15",           # loại hình được trừ cho C/O
    "types_other": "E11,E15,E13",    # loại hình được trừ cho XK / MDK / tiêu huỷ
}

_CODE = re.compile(r"^\s*(\d{3}-\d{5})")


def fold(s):
    """Chữ thường, bỏ dấu tiếng Việt — để lọc 'thong quan' khớp 'Thông quan'."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s).lower().replace("đ", "d"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def ma_key(code):
    """Chuẩn hoá mã hàng: NVL dạng 'nnn-nnnnn[-xx]' lấy 9 ký tự đầu (khớp cột 'Mã hải quan' của sổ cũ)."""
    code = (code or "").strip().upper()
    m = _CODE.match(code)
    return m.group(1) if m else code


def is_pg(conn):
    return getattr(conn, "is_pg", False)


def connect(path=None):
    p = Path(path) if path else DB_PATH
    p.parent.mkdir(parents=True, exist_ok=True)
    test_dsn = os.environ.get("NVL_TEST_PG")   # chạy bộ test trên PostgreSQL: mỗi DB tạm = một schema riêng
    if test_dsn:
        conn = pgcompat.PGConn(test_dsn, pgcompat.schema_for(p))
    elif pgcompat.dsn() and not os.environ.get("NVL_DB") and path is None:
        conn = pgcompat.PGConn(pgcompat.dsn())
    else:
        conn = sqlite3.connect(p)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA journal_mode=WAL")
    conn.executescript(SCHEMA)
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(import_lines)")}
    ev_cols = {r["name"] for r in conn.execute("PRAGMA table_info(events)")}
    if "frozen" not in ev_cols:
        conn.execute("ALTER TABLE events ADD COLUMN frozen INTEGER NOT NULL DEFAULT 0")
    if "tieu_chi" not in ev_cols:
        conn.execute("ALTER TABLE events ADD COLUMN tieu_chi TEXT")
    for c in ("nhap_tu", "nhap_den"):  # khoảng ngày ĐK tờ khai nhập được phép trừ cho lô (ISO; trống = không giới hạn)
        if c not in ev_cols:
            conn.execute(f"ALTER TABLE events ADD COLUMN {c} TEXT")
    goods_cols = {r["name"] for r in conn.execute("PRAGMA table_info(ob_goods)")}
    if "inv_no" not in goods_cols:  # dòng hàng thuộc Invoice nào (để nạp lại đúng Invoice không xoá dòng của Invoice khác)
        conn.execute("ALTER TABLE ob_goods ADD COLUMN inv_no TEXT")
    if "ghi_chu" not in cols:  # ghi chú nhập tay của người dùng, không bị ghi đè khi nạp lại file
        conn.execute("ALTER TABLE import_lines ADD COLUMN ghi_chu TEXT")
    if "raw" not in cols:  # nâng cấp DB cũ
        conn.execute("ALTER TABLE import_lines ADD COLUMN raw TEXT")
    if not is_pg(conn):
        conn.create_function("fold", 1, fold, deterministic=True)
    for k, v in DEFAULT_SETTINGS.items():
        conn.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    conn.commit()
    return conn


def get_settings(conn):
    return {r["key"]: r["value"] for r in conn.execute("SELECT key,value FROM settings")}

