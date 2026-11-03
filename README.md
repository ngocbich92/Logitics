# Sổ theo dõi NVL — hệ thống trừ lùi nguyên vật liệu

Chạy: bấm đúp `run.bat` (cần Python 3.10+ và `pip install openpyxl`). Mở http://127.0.0.1:8765.
Cơ sở dữ liệu: nếu có file `data/db_config.json` (`{"backend":"postgres","dsn":"postgresql://user:pass@host:5432/nvl"}`) hoặc biến môi trường `NVL_PG` thì dùng **PostgreSQL** (cần `pip install "psycopg[binary]"`, extension `unaccent`); không có thì dùng SQLite `data/nvl.db`. Chuyển dữ liệu SQLite → PostgreSQL: `python tools/migrate_to_pg.py "<dsn>"`. Sao lưu PostgreSQL: `pg_dump`. Kiểm thử: `python test_engine.py` (SQLite) hoặc đặt `NVL_TEST_PG=<dsn>` để chạy trên PostgreSQL.

**Cập nhật dữ liệu hằng ngày:** tab *Nhập dữ liệu* → chọn file Excel mới (sheet 01_DATABASE) → *Tải lên & cập nhật*.
Chế độ "toàn bộ dữ liệu" đồng bộ: thêm dòng mới, sửa dòng đổi, xoá dòng không còn (giữ dòng đã bị trừ lùi).
**Lọc:** tab *Tồn kho* — mọi cột (86 cột + Đã trừ/Tồn) có ô lọc riêng: chứa · `=chính xác` · `!loại trừ` · `a|b` · số `>100`, `5..20` · ngày `2026-09`, `01/09/2026..30/09/2026`; lọc không phân biệt dấu.

**Hai hệ thống trên thanh điều hướng dọc:** 📦 *Hệ thống quản lý NVL* và 🚢 *Outbound shipment* (dữ liệu Source_Data / List shipment / Tracking + 12 biểu mẫu Excel gốc: SI Form, DGD, VGM, CMA, Manifest, ZIM, MSC, ONE, CVC, DGD air, PK, INV). Nạp dữ liệu Outbound ở *Outbound → Nạp dữ liệu*. Muốn đổi mẫu biểu: sửa file Excel gốc rồi chạy python tools/make_outbound_template.py <file>.

| File | Vai trò |
|---|---|
| db.py | schema, kết nối (SQLite hoặc PostgreSQL), chuẩn hoá mã hàng |
| pgcompat.py | lớp tương thích PostgreSQL (giữ giao diện sqlite3) |
| engine.py | thuật toán trừ lùi LIFO/FIFO, khoá / huỷ khoá lô |
| importer.py | nạp báo cáo tờ khai, BOM, sổ Excel cũ |
| exporter.py | Excel: tồn kho, nhật ký trừ lùi, bảng kê RVC |
| server.py, static/index.html | giao diện web cục bộ (chỉ nghe 127.0.0.1) |
| outbound.py, outbound_forms.py, xl_eval.py | Outbound: dữ liệu, biểu mẫu, bộ tính công thức Excel (VLOOKUP, IF, ...) |
| rvc_form.py, report.py, boms.py, grid.py | RVC, báo cáo CIP/tỷ giá, BOM, lưới lọc |
