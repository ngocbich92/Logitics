"""Kiểm thử engine trừ lùi:  python test_engine.py"""
import tempfile
import unittest
from pathlib import Path

import db
import engine
import importer


def mk(conn, tk, ngay, loai, ma, qty, seq, status="Đã thông quan"):
    conn.execute("""INSERT INTO import_lines(seq,so_tk,ngay_dk,loai_hinh,ma_hang,ma_key,qty,dedupe,trang_thai)
                    VALUES(?,?,?,?,?,?,?,?,?)""", (seq, tk, ngay, loai, ma, db.ma_key(ma), qty, f"{tk}{ma}{seq}", status))


class EngineTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.conn = db.connect(Path(self.tmp.name) / "t.db")
        self.st = db.get_settings(self.conn)
        c = self.conn
        mk(c, "A", "2026-08-01", "E11", "110-00001", 100, 1)   # cũ nhất
        mk(c, "B", "2026-08-10", "E15", "110-00001", 50, 2)
        mk(c, "C", "2026-08-20", "E11", "110-00001", 30, 3)    # mới nhất
        mk(c, "D", "2026-08-25", "B11", "110-00001", 999, 4)   # loại hình không hợp lệ
        mk(c, "E", "2026-08-05", "E13", "110-00001", 40, 5)    # chỉ hợp lệ cho XK
        c.commit()

    def tearDown(self):
        self.conn.close()
        self.tmp.cleanup()

    def event(self, kind, code, lines):
        return importer.create_event(self.conn, kind, code, {}, [{"ma_hang": m, "qty": q} for m, q in lines])

    def parts(self, plan, i=0):
        return [(p["so_tk"], p["qty"]) for p in plan[i]["parts"]]

    def test_lifo_order_and_type_filter(self):
        e = self.event("CO", "CO1", [("110-00001", 70)])
        plan = engine.plan_event(self.conn, e, self.st)
        self.assertEqual(self.parts(plan), [("C", 30), ("B", 40)])  # B11 và E13 bị loại cho C/O
        self.assertEqual(plan[0]["shortfall"], 0)

    def test_second_event_sees_remaining_stock(self):
        e1 = self.event("CO", "CO1", [("110-00001", 70)])
        engine.lock_event(self.conn, e1, self.st)
        e2 = self.event("CO", "CO2", [("110-00001", 40)])
        plan = engine.plan_event(self.conn, e2, self.st)
        self.assertEqual(self.parts(plan), [("B", 10), ("A", 30)])

    def test_xk_can_use_e13(self):
        e = self.event("XK", "XK1", [("110-00001", 200)])
        plan = engine.plan_event(self.conn, e, self.st)
        self.assertEqual(self.parts(plan), [("C", 30), ("B", 50), ("E", 40), ("A", 80)])

    def test_shortfall_reported_and_lock_refused(self):
        e = self.event("CO", "CO1", [("110-00001", 500)])
        plan = engine.plan_event(self.conn, e, self.st)
        self.assertEqual(plan[0]["allocated"], 180)
        self.assertEqual(plan[0]["shortfall"], 320)
        with self.assertRaises(ValueError):
            engine.lock_event(self.conn, e, self.st)
        engine.lock_event(self.conn, e, self.st, allow_shortfall=True)

    def test_unlock_returns_stock(self):
        e = self.event("CO", "CO1", [("110-00001", 180)])
        engine.lock_event(self.conn, e, self.st)
        e2 = self.event("CO", "CO2", [("110-00001", 1)])
        self.assertEqual(engine.plan_event(self.conn, e2, self.st)[0]["shortfall"], 1)
        engine.unlock_event(self.conn, e)
        self.assertEqual(engine.plan_event(self.conn, e2, self.st)[0]["shortfall"], 0)

    def test_duplicate_code_lines_in_same_event_do_not_double_spend(self):
        e = self.event("CO", "CO1", [("110-00001", 100), ("110-00001", 100)])
        plan = engine.plan_event(self.conn, e, self.st)
        self.assertEqual(plan[0]["allocated"] + plan[1]["allocated"], 180)

    def test_fifo(self):
        self.conn.execute("UPDATE settings SET value='FIFO' WHERE key='order'")
        e = self.event("CO", "CO1", [("110-00001", 110)])
        plan = engine.plan_event(self.conn, e, self.st | {"order": "FIFO"})
        self.assertEqual(self.parts(plan), [("A", 100), ("B", 10)])

    def test_rebuild_gives_co_priority_and_respects_frozen(self):
        # XK khoá trước chiếm hết tồn C/O-hợp-lệ; sau đó C/O đến sau -> tính lại cho C/O ưu tiên
        xk = self.event("XK", "XK1", [("110-00001", 180)])
        engine.lock_event(self.conn, xk, self.st)
        co = self.event("CO", "CO1", [("110-00001", 100)])
        self.assertEqual(engine.plan_event(self.conn, co, self.st)[0]["shortfall"], 60)
        engine.lock_event(self.conn, co, self.st, allow_shortfall=True)
        r = engine.rebuild(self.conn, self.st, dry_run=True)
        self.assertEqual({c["code"] for c in r["changed"]}, {"XK1", "CO1"})
        self.assertEqual(engine.locked_plan(self.conn, co)[0]["shortfall"], 60)  # dry-run: chưa đổi gì
        engine.rebuild(self.conn, self.st, dry_run=False)
        self.assertEqual(engine.locked_plan(self.conn, co)[0]["shortfall"], 0)       # C/O giờ đủ
        self.assertEqual(engine.locked_plan(self.conn, xk)[0]["allocated"], 120)     # tổng tồn 220 - 100 (C/O)
        self.assertGreater(engine.locked_plan(self.conn, xk)[0]["shortfall"], 0)
        # lô cố định không bị đổi
        engine.set_frozen(self.conn, co, True)
        engine.set_frozen(self.conn, xk, True)
        self.assertEqual(engine.rebuild(self.conn, self.st)["lots_checked"], 0)

    def test_only_cleared_skips_phan_luong(self):
        mk(self.conn, "F", "2026-08-30", "E11", "110-00001", 500, 6, status="Đã phân luồng")
        e = self.event("CO", "CO1", [("110-00001", 10)])
        self.assertEqual(self.parts(engine.plan_event(self.conn, e, self.st)), [("C", 10)])   # mặc định bỏ qua F
        off = dict(self.st, only_cleared="0")
        self.assertEqual(self.parts(engine.plan_event(self.conn, e, off)), [("F", 10)])

    def test_import_date_range_per_event(self):
        def lot(tu, den):
            e = importer.create_event(self.conn, "CO", f"L{tu}{den}", {"nhap_tu": tu, "nhap_den": den}, [{"ma_hang": "110-00001", "qty": 200}])
            return engine.plan_event(self.conn, e, self.st)
        # A 1/8, B 10/8, C 20/8
        self.assertEqual(self.parts(lot(None, "12/08/2026")), [("B", 50), ("A", 100)])         # đến 12/8: bỏ C
        self.assertEqual(self.parts(lot("2026-08-05", None)), [("C", 30), ("B", 50)])          # từ 5/8: bỏ A
        self.assertEqual(self.parts(lot("2026-08-05", "2026-08-15")), [("B", 50)])             # khoảng 5/8–15/8
        self.assertEqual(self.parts(lot(None, None)), [("C", 30), ("B", 50), ("A", 100)])      # trống = không giới hạn

    def test_ma_key_normalisation(self):
        self.assertEqual(db.ma_key(" 825-00058-05 "), "825-00058")
        self.assertEqual(db.ma_key("MLH_H11"), "MLH_H11")


class SyncImportTest(unittest.TestCase):
    """Nạp lại file đầy đủ: thêm / cập nhật / xoá, không xoá dòng đã bị trừ lùi; lưới lọc theo kiểu cột."""

    HEAD = ["STT", "Số TK", "Ngày ĐK", "Mã loại hình", "Trạng thái", "Mã hàng", "Số lượng", "Đơn giá hóa đơn", "Tên hàng"]

    def book(self, rows):
        import datetime as dt
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["tiêu đề"])
        ws.append(self.HEAD)
        for r in rows:
            ws.append([1, r[0], dt.datetime.fromisoformat(r[1]), r[2], r[3], r[4], r[5], 1.0, r[6]])
        return ws

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.conn = db.connect(Path(self.tmp.name) / "s.db")
        self.st = db.get_settings(self.conn)

    def tearDown(self):
        self.conn.close()
        self.tmp.cleanup()

    def test_full_sync(self):
        import grid
        v1 = [(111, "2026-09-02", "E11", "Đã thông quan", "110-00001", 100, "Pin lithium"),
              (222, "2026-09-01", "E15", "Đã phân luồng", "110-00001", 50, "Pin lithium"),
              (333, "2026-08-01", "E11", "Đã thông quan", "110-00002", 10, "Tụ điện")]
        r = importer.import_database_sheet(self.conn, self.book(v1), "v1")
        self.assertEqual((r["added"], r["removed"]), (3, 0))
        e = importer.create_event(self.conn, "CO", "L1", {}, [{"ma_hang": "110-00002", "qty": 4}])
        engine.lock_event(self.conn, e, self.st)
        # v2: TK 222 đổi trạng thái, TK 333 biến mất (đã bị trừ → giữ lại), thêm TK 444
        v2 = [(444, "2026-09-05", "E11", "Đã thông quan", "110-00003", 7, "Dây"),
              (111, "2026-09-02", "E11", "Đã thông quan", "110-00001", 100, "Pin lithium"),
              (222, "2026-09-01", "E15", "Đã thông quan", "110-00001", 50, "Pin lithium")]
        r = importer.import_database_sheet(self.conn, self.book(v2), "v2")
        self.assertEqual((r["added"], r["updated"], r["removed"], r["kept_because_already_allocated"]), (1, 1, 0, 1))
        r = importer.import_database_sheet(self.conn, self.book(v2), "v2")
        self.assertEqual((r["added"], r["updated"], r["removed"]), (0, 0, 0))
        # lọc: bỏ dấu, số, ngày
        self.assertEqual(grid.query(self.conn, {"Trạng thái": "thong quan"}, "", "", "asc", 1, 50, None)["total"], 4)
        self.assertEqual(grid.query(self.conn, {"Số lượng": ">=50"}, "", "", "asc", 1, 50, None)["total"], 2)
        self.assertEqual(grid.query(self.conn, {"Ngày ĐK": "2026-09"}, "", "", "asc", 1, 50, None)["total"], 3)
        self.assertEqual(grid.query(self.conn, {"Tên hàng": "pin|day"}, "", "", "asc", 1, 50, None)["total"], 3)
        self.assertEqual(grid.query(self.conn, {"Tồn (hệ thống)": "0"}, "", "", "asc", 1, 50, None)["total"], 0)
        # dòng không còn trong file và chưa bị trừ thì bị xoá
        r = importer.import_database_sheet(self.conn, self.book(v2[1:]), "v3")
        self.assertEqual(r["removed"], 1)


class DuplicateReviewTest(SyncImportTest):
    def test_duplicates_reported_then_dropped_on_confirm(self):
        rows = [(111, "2026-09-02", "E11", "Đã thông quan", "110-00001", 100, "Pin"),
                (111, "2026-09-02", "E11", "Đã thông quan", "110-00001", 100, "Pin"),      # giống hệt
                (222, "2026-09-02", "E11", "Đã thông quan", "110-00002", 5, "Tụ"),
                (222, "2026-09-02", "E11", "Đã thông quan", "110-00002", 5, "Tụ khác")]     # khác Tên hàng
        data = importer.parse_database_sheet(self.book(rows))
        groups = importer.duplicate_groups(data["parsed"])
        self.assertEqual([(g["so_tk"], g["kind"], g["n"]) for g in groups], [("111", "identical", 2), ("222", "differs", 2)])
        self.assertIn("Tên hàng", groups[1]["diff_cols"])
        # xem trước không ghi gì
        r = importer.apply_database_import(self.conn, data, dry_run=True)
        self.assertEqual(r["added"], 4)
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM import_lines").fetchone()[0], 0)
        # xác nhận bỏ thừa nhóm 111
        key = groups[0]["key"]
        r = importer.apply_database_import(self.conn, data, "f", dedupe_keys=[key])
        self.assertEqual((r["added"], r["duplicates_dropped"]), (3, 1))
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM import_lines").fetchone()[0], 3)
        # nạp lại không đổi gì; giữ cả 2 dòng 111 thì thêm lại dòng thứ hai
        r = importer.apply_database_import(self.conn, data, "f", dedupe_keys=[key])
        self.assertEqual((r["added"], r["updated"], r["removed"]), (0, 0, 0))
        r = importer.apply_database_import(self.conn, data, "f")
        self.assertEqual(r["added"], 1)

    test_full_sync = None  # kế thừa chỉ để dùng book()/setUp


class LotsImportTest(unittest.TestCase):
    def setUp(self):
        import openpyxl
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.conn = db.connect(self.dir / "l.db")
        self.st = db.get_settings(self.conn)
        mk(self.conn, "A", "2026-08-01", "E11", "110-00001", 100, 1)
        mk(self.conn, "B", "2026-08-10", "E15", "110-00001", 50, 2)
        mk(self.conn, "B", "2026-08-10", "E15", "110-00002", 10, 3)
        self.conn.commit()
        self.openpyxl = openpyxl

    def tearDown(self):
        self.conn.close()
        self.tmp.cleanup()

    def rows(self, data):
        wb = self.openpyxl.Workbook()
        ws = wb.active
        ws.append(["Mã lô / INV", "Loại", "Số TK nhập", "Mã hàng", "Số lượng", "Ngày TK xuất"])
        for r in data:
            ws.append(list(r))
        p = self.dir / "lots.xlsx"
        wb.save(p)
        import lots_import
        return lots_import.parse_file(p)

    def test_explicit_allocation_validation_and_replace(self):
        import lots_import as L
        rows = self.rows([("L1", "CO", "A", "110-00001", 30, "15/07/2026"),
                          ("L1", "CO", "B", "110-00001", 20, None),
                          ("L1", "CO", "Z", "110-00001", 1, None),          # TK không tồn tại
                          ("L1", "CO", "B", "110-00002", 99, None),         # vượt tồn
                          ("L2", "", "A", "110-00001", 70, None)])          # loại mặc định XK; đủ 70 còn lại của A
        r = L.apply(self.conn, rows, "XK", dry_run=True)
        self.assertEqual((r["ok_rows"], r["n_errors"]), (3, 2))
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM events").fetchone()[0], 0)      # xem trước không ghi
        L.apply(self.conn, rows, "XK", dry_run=False)
        ev = self.conn.execute("SELECT status, frozen, ngay_tk FROM events WHERE code='L1'").fetchone()
        self.assertEqual((ev["status"], ev["frozen"], ev["ngay_tk"]), ("LOCKED", 1, "2026-07-15"))
        used = dict(self.conn.execute("SELECT i.so_tk, SUM(a.qty) FROM allocations a JOIN import_lines i ON i.id=a.import_line_id "
                                      "WHERE i.ma_key='110-00001' GROUP BY i.so_tk").fetchall())
        self.assertEqual(used, {"A": 100.0, "B": 20.0})
        # tồn còn lại được lô LIFO sau đó nhìn thấy: chỉ còn B = 30
        e = importer.create_event(self.conn, "CO", "NEW", {}, [{"ma_hang": "110-00001", "qty": 100}])
        plan = engine.plan_event(self.conn, e, self.st)
        self.assertEqual(plan[0]["allocated"], 30)
        # nạp lại cùng lô: bị từ chối trừ khi chọn thay thế
        r = L.apply(self.conn, rows, "XK", dry_run=True)
        self.assertEqual(r["ok_rows"], 0)
        r = L.apply(self.conn, rows, "XK", replace=True, dry_run=False)
        self.assertEqual(r["ok_rows"], 3)

    def test_split_across_duplicate_lines_of_same_tk(self):
        import lots_import as L
        mk(self.conn, "C", "2026-08-12", "E11", "110-00003", 5, 4)
        mk(self.conn, "C", "2026-08-12", "E11", "110-00003", 5, 5)   # 2 dòng cùng TK + mã
        self.conn.commit()
        r = L.apply(self.conn, self.rows([("L9", "XK", "C", "110-00003", 8, None)]), dry_run=False)
        self.assertEqual(r["ok_rows"], 1)
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM allocations").fetchone()[0], 2)

    def test_header_detection_error(self):
        import lots_import as L
        wb = self.openpyxl.Workbook()
        wb.active.append(["a", "b"])
        wb.save(self.dir / "bad.xlsx")
        with self.assertRaises(ValueError):
            L.parse_file(self.dir / "bad.xlsx")


class ReportTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.conn = db.connect(Path(self.tmp.name) / "r.db")

    def tearDown(self):
        self.conn.close()
        self.tmp.cleanup()

    def add(self, tk, ngay, ccy, rate, tt, seq, nuoc="CHINA"):
        self.conn.execute("""INSERT INTO import_lines(seq,so_tk,ngay_dk,loai_hinh,ma_hang,ma_key,qty,dedupe,trang_thai,ccy,ty_gia,don_gia_tt,nuoc_xx,ma_hs,ten_hang)
                             VALUES(?,?,?,'E11','110-00001','110-00001',100,?,?,?,?,?,?,'85076090','Pin')""",
                          (seq, tk, ngay, tk, "Đã thông quan", ccy, rate, tt, nuoc))

    def test_weekly_customs_rate_and_cif(self):
        import report
        # tuần 10/8–16/8/2026 (thứ Hai–Chủ nhật): tỷ giá USD 25500 (xuất hiện 2 lần) — dòng EUR ngày 14/8 dùng tỷ giá USD của tuần
        self.add("T1", "2026-08-10", "USD", 25500, 50000, 1)
        self.add("T2", "2026-08-12", "USD", 25500, 52000, 2)
        self.add("T3", "2026-08-14", "EUR", 29000, 60000, 3)
        self.add("T4", "2026-08-20", "EUR", 29000, 70000, 4)     # tuần 17–23/8 không có dòng USD
        self.conn.commit()
        rt = report.Rates(self.conn)
        self.assertEqual(rt.get("2026-08-16")[0], 25500)                  # Chủ nhật vẫn thuộc tuần đó
        self.assertEqual(rt.get("2026-08-17")[0], 25500)                  # tuần sau: dùng tuần gần nhất trước đó + ghi chú
        self.assertTrue(rt.get("2026-08-17")[2])
        self.assertIsNone(rt.get("2026-08-02")[0])
        # tỷ giá nhập tay (website hải quan) ưu tiên hơn
        report.save_fx(self.conn, {"date_from": "17/08/2026", "date_to": "23/08/2026", "usd": 25800})
        rt = report.Rates(self.conn)
        self.assertEqual(rt.get("2026-08-20"), (25800, "2026-08-17→2026-08-23", ""))
        e = importer.create_event(self.conn, "XK", "X1", {}, [{"ma_hang": "110-00001", "qty": 250}])
        rows = report.lot_rows(self.conn, e, db.get_settings(self.conn), rt)
        self.assertEqual(sum(r["qty"] for r in rows), 250)
        r4 = next(r for r in rows if r["so_tk"] == "T4")
        self.assertAlmostEqual(r4["don_gia_cif"], 70000 / 25800, places=5)        # đơn giá tính thuế VNĐ / tỷ giá USD
        self.assertEqual((r4["nuoc_xx"], r4["ma_hs"], r4["ten_hang"]), ("CHINA", "85076090", "Pin"))
        with self.assertRaises(ValueError):
            report.save_fx(self.conn, {"date_from": "2026-08-10", "date_to": "2026-08-01", "usd": 25000})

    def test_ledger_includes_draft_and_locked(self):
        import report
        self.add("T1", "2026-08-10", "USD", 25500, 50000, 1)
        self.conn.commit()
        st = db.get_settings(self.conn)
        e1 = importer.create_event(self.conn, "XK", "A", {}, [{"ma_hang": "110-00001", "qty": 10}])
        engine.lock_event(self.conn, e1, st)
        importer.create_event(self.conn, "XK", "B", {}, [{"ma_hang": "110-00001", "qty": 5}])
        rows = report.ledger(self.conn, "XK", st)
        self.assertEqual({(r["lo"], r["trang_thai"]) for r in rows}, {("A", "Đã ghi sổ"), ("B", "Nháp — dự kiến")})


class SheetPickingTest(unittest.TestCase):
    """Tự tìm sheet 'BOM' / 'Export Material' và cột CIP + ghi chú trong lưới tồn kho."""

    def wb(self):
        import openpyxl
        wb = openpyxl.Workbook()
        wb.active.title = "Khac"
        b = wb.create_sheet("BOM")
        for r, (k, v) in enumerate([("Mã thành phẩm (Mã sản xuất hàng)", "FG9"), ("Tên hàng hóa", "Pin 9"), (None, None), ("Số lượng lô hàng", 1),
                                    ("ĐVT lô hàng", "EA"), ("Mã HS hàng hóa (06 số)", "850760")], start=4):
            b.cell(r, 1, k)
            b.cell(r, 2, v)
        for c, v in enumerate(["Mã hàng (NVL)", "Item description", "Qty used in FG", "Unit"], start=1):
            b.cell(16, c, v)
        for r, row in enumerate([["110-00001", "Cell", 15, "EACH"], ["410-00002", "Keo", 0.44, "ML"]], start=17):
            for c, v in enumerate(row, start=1):
                b.cell(r, c, v)
        e = wb.create_sheet("export material")
        for r, (k, v) in enumerate([("Số TK xuất khẩu", "3087"), ("Ngày ĐK (TK xuất)", "15/07/2026"), ("Số INV", "VNOUT801"), ("Ngày INV", "01/07/2026")], start=4):
            e.cell(r, 1, k)
            e.cell(r, 2, v)
        for c, v in enumerate(["STT", "Mã hàng", "Item description", "Qty", "Unit"], start=1):
            e.cell(10, c, v)
        for r, row in enumerate([[1, "110-00001", "Cell", 30, "PCE"], [2, "410-00002", "Keo", 0, "ML"]], start=11):   # Qty 0: bỏ qua
            for c, v in enumerate(row, start=1):
                e.cell(r, c, v)
        return wb

    def test_find_and_read_sheets(self):
        import boms
        import openpyxl
        wb = self.wb()
        self.assertEqual(importer.sheets_named(wb, r"bom(\s*\(\d+\))?"), ["BOM"])
        self.assertEqual(importer.sheets_named(wb, r"export\s*material"), ["export material"])
        h, lines = importer.read_export_material(wb["export material"])
        self.assertEqual((h["so_inv"], h["so_tk_xuat"], h["ngay_tk"]), ("VNOUT801", "3087", "2026-07-15"))
        self.assertEqual([(l["ma_hang"], l["qty"]) for l in lines], [("110-00001", 30.0)])
        h, lines = importer.read_bom_sheet(wb["BOM"])
        self.assertEqual((h["ma_tp"], h["hs6"], h["dvt"]), ("FG9", "850760", "EA"))
        self.assertEqual(len(lines), 2)
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "f.xlsx"
            wb.save(p)
            rows, sheets = boms.parse_any(p)
            self.assertEqual((sheets, len(rows), rows[0]["ma_tp"], rows[0]["dvt_tp"]), (["BOM"], 2, "FG9", "EA"))

    def test_grid_cip_and_note(self):
        import grid
        import openpyxl
        with tempfile.TemporaryDirectory() as tmp:
            conn = db.connect(Path(tmp) / "g.db")
            ws = openpyxl.Workbook().active
            ws.append(["t"])
            ws.append(["Số TK", "Ngày ĐK", "Mã loại hình", "Trạng thái", "Mã hàng", "Số lượng", "Đơn giá hóa đơn", "Ng.Tệ hóa đơn", "Tỷ giá VNĐ", "Đơn giá tính thuế"])
            import datetime as dt
            ws.append(["T1", dt.datetime(2026, 8, 12), "E11", "Đã thông quan", "110-00001", 10, 2.0, "USD", 25000, 50000])
            ws.append(["T2", dt.datetime(2026, 8, 20), "E11", "Đã thông quan", "110-00001", 10, 2.0, "EUR", 29000, 60000])   # tuần không có USD
            importer.import_database_sheet(conn, ws)
            r = grid.query(conn, {}, "", "", "asc", 1, 50, ["Số TK", grid.RATE, grid.CIP, grid.NOTE])
            self.assertEqual(r["rows"][0][1], 25000)
            self.assertAlmostEqual(r["rows"][0][2], 2.0)
            self.assertAlmostEqual(r["rows"][1][2], 60000 / 25000)           # tuần 17–23/8 dùng tuần gần nhất trước
            conn.execute("UPDATE import_lines SET ghi_chu='kiểm tra lại' WHERE so_tk='T1'")
            conn.commit()
            r = grid.query(conn, {grid.NOTE: "kiem tra"}, "", "", "asc", 1, 50, ["Số TK", grid.NOTE])
            self.assertEqual([x[0] for x in r["rows"]], ["T1"])
            r = grid.query(conn, {grid.CIP: ">=2.3"}, "", "", "asc", 1, 50, ["Số TK"])
            self.assertEqual([x[0] for x in r["rows"]], ["T2"])
            # nạp lại file không làm mất ghi chú
            importer.import_database_sheet(conn, ws)
            self.assertEqual(conn.execute("SELECT ghi_chu FROM import_lines WHERE so_tk='T1'").fetchone()[0], "kiểm tra lại")
            conn.close()


class XlEvalTest(unittest.TestCase):
    class Ctx:
        def __init__(self):
            self.cells = {("S", 1, 1): 5.0, ("S", 2, 1): "abc", ("S", 3, 1): None}
            self.rows = {"k1": {"B": "k1", "C": 12.0, "D": None, "E": "text"}}

        def cell_value(self, sheet, c, r):
            return self.cells.get((sheet, c, r))

        def lookup_row(self, sheet, key):
            return self.rows.get(key.lower()) if sheet == "Data" else None

        def today(self):
            import datetime as dt
            return dt.date(2026, 10, 13)

    def ev(self, f):
        import xl_eval
        return xl_eval.Evaluator(self.Ctx()).ev(xl_eval.parse(f), "S")

    def test_excel_semantics(self):
        import xl_eval
        self.assertEqual(self.ev('=IF(A1>3,"big","small")'), "big")
        self.assertEqual(self.ev('=A1&" CNTS="&B1'), "5 CNTS=abc")                         # số nguyên không có .0
        self.assertEqual(self.ev('=IF(C1="","empty","filled")'), "empty")                  # ô trống = ""
        self.assertEqual(self.ev('=VLOOKUP("K1",Data!B:E,2,0)'), 12.0)                      # không phân biệt hoa thường
        self.assertEqual(self.ev('=IF(VLOOKUP("k1",Data!B:E,3,0)="","","x")'), "")         # ô trống trong VLOOKUP
        self.assertEqual(self.ev('=IFERROR(VLOOKUP("zzz",Data!B:E,2,0),"none")'), "none")
        self.assertEqual(self.ev('=IFERROR(VLOOKUP("k1",#REF!,2,0),"bad")'), "bad")
        self.assertEqual(self.ev("=ROUND(2.5,0)+ROUNDUP(1.01,1)"), 3.0 + 1.1)
        self.assertEqual(self.ev('=LEFT("Polarium",8)&MID("abcdef",2,3)'), "Polariumbcd")
        self.assertEqual(self.ev('=TEXT(TODAY(),"yyyy/mm/dd")'), "2026/10/13")
        self.assertEqual(self.ev('=AND(A1=5,B1="ABC")'), True)
        with self.assertRaises(xl_eval.XlError):
            self.ev('=B1+1')                                                                # chữ + số -> #VALUE!
        with self.assertRaises(xl_eval.XlError):
            self.ev('="" + 1')


class OutboundTest(unittest.TestCase):
    def test_import_edit_and_form(self):
        import datetime as dt
        import openpyxl
        import outbound
        import outbound_forms
        with tempfile.TemporaryDirectory() as tmp:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Source_Data"
            heads = ["No.", "Shipment ID", "SO No.", "Items number", "Booking No.", "Shipper / Exporter", "Consignee"]
            for i, h in enumerate(heads, start=1):
                ws.cell(1, i, h)
                ws.cell(2, i, "Logs")
                ws.cell(3, i, f"-{i}")
            ws.append([1, "VNOUT00900", "SO-1", "800-00182", "BK123", "Polarium VN", "ACME Corp"])
            ws.append([2, "VNOUT00901", "SO-2", "800-00183", None, "Polarium VN", "Beta"])
            ws2 = wb.create_sheet("List shipment")
            ws2.append([])
            ws2.append(["No.", "Separate ID#", "Qty", "Pick up"])
            ws2.append([1, "VNOUT00900-1", 10, dt.datetime(2026, 9, 1)])
            p = Path(tmp) / "ob.xlsx"
            wb.save(p)
            conn = db.connect(Path(tmp) / "o.db")
            cand, _ = outbound.sheet_candidates(p)
            self.assertEqual(set(cand), {"source", "list"})
            r = outbound.import_workbook(conn, p, cand)
            self.assertEqual((r["source"]["after"], r["list"]["after"]), (2, 1))
            self.assertEqual(outbound.find_row(conn, "source", "vnout00900")["E"], "BK123")            # tra như VLOOKUP, không phân biệt hoa thường
            self.assertEqual(outbound.find_row(conn, "list", "VNOUT00900-1")["D"].date(), dt.date(2026, 9, 1))
            # lưới: lọc / sắp xếp / sửa / thêm / xoá
            g = outbound.query(conn, "source", {"Consignee": "acme"}, "", "", "asc", 1, 50)
            self.assertEqual((g["total"], g["rows"][0][1]), (1, "VNOUT00900"))
            rid = g["ids"][0]
            outbound.update_row(conn, rid, {"E": "BK999", "B": "VNOUT00900X"})
            self.assertIsNone(outbound.find_row(conn, "source", "VNOUT00900"))
            self.assertEqual(outbound.find_row(conn, "source", "vnout00900x")["E"], "BK999")
            nid = outbound.add_row(conn, "source", {"B": "VNOUT00902", "C": "SO-9"})
            self.assertEqual(outbound.get_row(conn, nid)["key"], "VNOUT00902")
            outbound.delete_row(conn, nid)
            self.assertEqual(outbound.overview(conn)[0]["rows"], 2)
            # nạp lại thay toàn bộ
            r = outbound.import_workbook(conn, p, {"source": "Source_Data"})
            self.assertEqual((r["source"]["before"], r["source"]["after"]), (2, 2))
            # biểu mẫu: SI Form!E3 = VLOOKUP(..., Source_Data!B:BH, 4) = Booking No.; B3 (cột 5) = Shipper
            wbf, errs = outbound_forms.build(conn, "VNOUT00900", None, ["SI Form"])
            self.assertEqual(wbf.sheetnames, ["SI Form"])
            self.assertEqual(wbf["SI Form"]["A2"].value, "VNOUT00900")
            self.assertEqual(wbf["SI Form"]["E3"].value, "BK123")
            self.assertEqual(wbf["SI Form"]["B3"].value, "Polarium VN")
            conn.close()


class ObExtrasTest(unittest.TestCase):
    """Cột bổ sung (Số tờ khai/Invoice...), file đính kèm, Information of goods, khung viền đỏ DGD — và biểu mẫu không bị ảnh hưởng."""

    def setUp(self):
        import openpyxl
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Source_Data"
        for i, h in enumerate(["No.", "Shipment ID", "SO No.", "Items number", "Booking No.", "Shipper / Exporter"], start=1):
            ws.cell(1, i, h)
            ws.cell(3, i, f"-{i}")
        ws.append([1, "VNOUT00900", "SO-1", "800-00182", "BK123", "Polarium VN"])
        ws.append([2, "VNOUT00901", "SO-2", "800-00183", "BK456", "Polarium VN"])
        t = wb.create_sheet("Tracking")
        for i, h in enumerate(["STT", None, "Shipment ID", "Note"], start=1):
            t.cell(3, i, h)
        t.append([1, None, "VNOUT00900", "trip 1"])
        t.append([2, None, "VNOUT00900", "trip 2"])
        self.src = self.dir / "ob.xlsx"
        wb.save(self.src)
        import os
        self._old = os.environ.get("NVL_DB")
        os.environ["NVL_DB"] = str(self.dir / "x.db")
        import importlib
        importlib.reload(db)
        self.conn = db.connect()

    def tearDown(self):
        import os
        import importlib
        self.conn.close()
        if self._old is None:
            os.environ.pop("NVL_DB", None)
        else:
            os.environ["NVL_DB"] = self._old
        importlib.reload(db)
        self.tmp.cleanup()

    def test_extras_files_goods_and_forms(self):
        import outbound
        import outbound_forms
        sel = {"source": "Source_Data", "tracking": "Tracking"}
        outbound.import_workbook(self.conn, self.src, sel)
        info = outbound.table_info(self.conn, "source")
        self.assertEqual([c["name"] for c in info["cols"][-6:]], ["Số tờ khai", "Ngày tờ khai", "Số Invoice", "Type of BL", "UN", "Actual pick up"])
        g = outbound.query(self.conn, "source", {}, "", "", "asc", 1, 10)
        rid = g["ids"][0]
        form_before, _ = outbound_forms.build(self.conn, "VNOUT00900", None, ["SI Form"])
        before = {c.coordinate: c.value for row in form_before["SI Form"].iter_rows() for c in row if c.value is not None}
        outbound.update_row(self.conn, rid, {"+SoTK": "108123", "+NgayTK": "15/09/2026", "+Invoice": "INV-9", "+TypeBL": "Telex", "+UN": "UN3480"})
        row = outbound.get_row(self.conn, rid)
        self.assertEqual((row["cells"]["+SoTK"], row["cells"]["+NgayTK"], row["cells"]["+TypeBL"]), ("108123", "2026-09-15", "Telex"))
        self.assertEqual(outbound.query(self.conn, "source", {"Số Invoice": "inv-9"}, "", "", "asc", 1, 10)["total"], 1)
        # biểu mẫu không đổi khi có thêm trường bổ sung
        form_after, _ = outbound_forms.build(self.conn, "VNOUT00900", None, ["SI Form"])
        after = {c.coordinate: c.value for row in form_after["SI Form"].iter_rows() for c in row if c.value is not None}
        self.assertEqual(before.keys(), after.keys())
        self.assertEqual([k for k in before if before[k] != after[k] and not hasattr(before[k], "year")], [])
        # file đính kèm: chỉ cho phép hình ảnh/pdf..., tính ở cột File, tách theo dòng Tracking
        tg = outbound.query(self.conn, "tracking", {}, "", "", "asc", 1, 10)
        trid = tg["ids"][1]
        f = outbound.add_file(self.conn, trid, "poD.jpg", b"\xff\xd8\xff-fake")
        with self.assertRaises(ValueError):
            outbound.add_file(self.conn, trid, "evil.html", b"<script>")
        with self.assertRaises(ValueError):
            outbound.add_file(self.conn, trid, "evil.exe", b"MZ")
        tg = outbound.query(self.conn, "tracking", {}, "", "", "asc", 1, 10)
        self.assertEqual([r[-1] for r in tg["rows"]], [None, 1])
        # goods: lưu, tự tính Amount, nạp file Excel
        n = outbound.goods_save(self.conn, "VNOUT00900", [{"ma_hang": "800-1", "ten_hang": "Pin", "qty": "10", "don_gia": "2,5", "ccy": "eur"},
                                                          {"ma_hang": "", "ten_hang": ""}])
        self.assertEqual(n, 1)
        gl = outbound.goods_get(self.conn, "vnout00900")
        self.assertEqual((gl[0]["amount"], gl[0]["ccy"]), (25.0, "EUR"))
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Shipment ID", "Item", "Description", "Qty", "Unit", "Unit price", "Currency", "Origin", "Net weight", "Gross weight"])
        ws.append(["VNOUT00901", "800-2", "Cell", 4, "PCE", 3, "SEK", "China", 1.2, 1.5])
        ws.append(["VNOUT00901", "800-3", "Cap", 2, "PCE", 1.5, "USD", "Vietnam", 0.2, 0.3])
        gp = self.dir / "goods.xlsx"
        wb.save(gp)
        rows, warn = outbound.goods_parse_file(gp)
        self.assertEqual((len(rows), rows[0]["ship"], rows[0]["ma_hang"], rows[0]["ccy"]), (2, "VNOUT00901", "800-2", "SEK"))
        res = outbound.goods_import_bulk(self.conn, rows, dry_run=False)
        self.assertEqual(res["ships"][0]["lines"], 2)
        self.assertEqual(sum(g["amount"] for g in outbound.goods_get(self.conn, "VNOUT00901")), 15.0)
        # nạp lại file Excel: cột bổ sung + file đính kèm đi theo dòng; hàng của shipment không đổi
        outbound.import_workbook(self.conn, self.src, sel)
        g2 = outbound.query(self.conn, "source", {}, "", "", "asc", 1, 10)
        self.assertEqual(outbound.get_row(self.conn, g2["ids"][0])["cells"]["+SoTK"], "108123")
        tg2 = outbound.query(self.conn, "tracking", {}, "", "", "asc", 1, 10)
        self.assertEqual([r[-1] for r in tg2["rows"]], [None, 1])
        self.assertEqual(len(outbound.goods_get(self.conn, "VNOUT00901")), 2)
        # xoá dòng xoá luôn file trên đĩa
        stored = [p.name for p in outbound.files_dir().iterdir()]
        self.assertEqual(len(stored), 1)
        outbound.delete_row(self.conn, tg2["ids"][1])
        self.assertEqual(list(outbound.files_dir().iterdir()), [])

    def test_overview_report_from_source_data(self):
        import datetime as dt
        import openpyxl
        import outbound
        from openpyxl.utils import column_index_from_string as ci
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Source_Data"
        for c in range(1, ci("AU") + 1):
            ws.cell(1, c, f"col{c}")
        #   id, status(AT), movement(Q), pickup(O), fwd(AU), ctns(S), container(Z)
        rows = [("S1", "Done", "FCL/FCL", "2023-01-05", "dhl", 10, "ECMU9896298\nCMAU2018976"),
                ("S1-1", "Done", "FCL/FCL", None, None, 5, "TRHU2733726"),                 # dòng container phụ gộp vào S1
                ("S2", "Processing", "Air/LCL", "2023-02-01", "dhl", 3, "Air/LCL"),          # 'Air/LCL' không phải số container
                ("S2-DG", "Processing", "Air/LCL", "2023-02-01", "dhl", 3, None),            # bản DG gộp vào S2
                ("S3", "DOne", "AIR", "2023-02-09", "DHL ", 1, None),
                ("S4", "Done/Roll the BCF", "TRUCK", "2023-03-01", None, 2, "29C-23025"),
                ("S5", "Combine/Delete", "FCL/FCL", None, None, 7, None),
                ("S6", "Done", "LCL", None, None, 1, None), ("S6-2", "Processing", "LCL", None, None, 1, None),
                ("V-23-00008", "Done", None, "2023-03-05", "x", 1, None)]
        for i, r in enumerate(rows, start=4):
            ws.cell(i, 2, r[0])
            ws.cell(i, ci("AT"), r[1])
            ws.cell(i, ci("Q"), r[2])
            ws.cell(i, ci("O"), r[3])
            ws.cell(i, ci("AU"), r[4])
            ws.cell(i, ci("S"), r[5])
            ws.cell(i, ci("Z"), r[6])
        p = self.dir / "rep.xlsx"
        wb.save(p)
        outbound.import_workbook(self.conn, p, {"source": "Source_Data"})
        r = outbound.report(self.conn)
        # S1, S2, S3, S4, S6, V-23-00008 = 6 shipment (S5 đã huỷ); S6 lẫn Done/Processing = Processing
        self.assertEqual(r["total"], 6)
        self.assertEqual({m["mode"]: m["count"] for m in r["modes"]}, {"AIR": 2, "SEA": 2, "TRUCK": 1, "OTHER": 1})
        self.assertEqual((r["progress"]["done"], r["progress"]["active"], r["progress"]["cancelled"]), (4, 6, 1))
        self.assertEqual(r["containers"], {"count": 4, "shipments": 2})                       # ECMU…, CMAU…, TRHU… (S1) + 29C-23025 (S4)
        self.assertEqual(r["qty"], 15 + 6 + 1 + 2 + 2 + 1)
        only = outbound.report(self.conn, ["Done"])
        self.assertEqual(only["total"], 4)
        self.assertEqual([(x["v"], x["n"]) for x in only["forwarders"]], [("DHL", 2), ("X", 1)])      # gộp "dhl" / "DHL "
        feb = outbound.report(self.conn, None, "2023-02-01", "2023-02-28")
        self.assertEqual((feb["total"], feb["progress"]["active"]), (2, 2))
        self.assertEqual([m["month"] for m in r["months"]], ["2023-01", "2023-02", "2023-03", "Chưa có ngày"])
        self.assertGreater(len(outbound.report_xlsx(self.conn, ["Done"])), 2000)
    def test_list_shipment_invoice_lines_shared_by_shipment(self):
        import openpyxl
        import outbound
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "List shipment"
        heads = ["No.", "Separate ID#", "Orders", "Produced", "Quantity", "Row updated", "Shipment#", "Item", "Description"]
        for i, h in enumerate(heads, start=1):
            ws.cell(2, i, h)
        ws.append([1, "VNOUT00004-1", "SO1", None, 5, None, "VNOUT00004", "800-1", "Cell"])
        ws.append([2, "VNOUT00004-2", "SO1", None, 3, None, "VNOUT00004", "800-2", "Cap"])
        ws.append([3, "VNOUT00006-1", "SO2", None, 1, None, "VNOUT00006", "800-3", "Pack"])
        p = self.dir / "ls.xlsx"
        wb.save(p)
        outbound.import_workbook(self.conn, p, {"list": "List shipment"})
        g = outbound.query(self.conn, "list", {}, "", "", "asc", 1, 10)
        rid1, rid2, rid3 = g["ids"]
        outbound.goods_save(self.conn, outbound.get_row(self.conn, rid1)["ship"],
                            [{"ma_hang": "800-1", "ten_hang": "Cell", "qty": 5, "dvt": "PCE", "don_gia": 2, "ccy": "SEK", "xuat_xu": "VN"}])
        self.assertEqual(len(outbound.get_row(self.conn, rid2)["goods"]), 1)           # dòng cùng Shipment# dùng chung dòng hàng
        self.assertEqual(outbound.get_row(self.conn, rid3)["goods"], [])
        g = outbound.query(self.conn, "list", {}, "", "", "asc", 1, 10)
        self.assertEqual(g["columns"][-1], "Dòng hàng (Invoice)")
        self.assertEqual([r[-1] for r in g["rows"]], [1, 1, None])

    def test_invoice_lines_page_import_modes_filters_and_link(self):
        import openpyxl
        import outbound
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "List shipment"
        for i, h in enumerate(["No.", "Separate ID#", "Orders", "Produced", "Quantity", "Row updated", "Shipment#"], start=1):
            ws.cell(2, i, h)
        ws.append([1, "VNOUT00004-1", "SO1", None, 5, None, "VNOUT00004"])
        ws.append([2, "VNOUT00004-2", "SO1", None, 3, None, "VNOUT00004"])
        p = self.dir / "ls2.xlsx"
        wb.save(p)
        outbound.import_workbook(self.conn, p, {"list": "List shipment"})
        rows = [{"row": 2, "ship": "VNOUT00004", "ma_hang": "A", "ten_hang": "Cell", "qty": 10, "don_gia": 2.5, "ccy": "usd"},
                {"row": 3, "ship": None, "ma_hang": "B", "ten_hang": "Cap", "qty": 4, "don_gia": 5, "ccy": "SEK"},
                {"row": 4, "ship": "NOPE-1", "ma_hang": "C", "qty": 1}]
        dry = outbound.goods_import_bulk(self.conn, rows, dry_run=True)
        self.assertEqual(dry["n_errors"], 1)                                                    # dòng không có Shipment ID và không có mặc định
        self.assertEqual(len(outbound.goods_query(self.conn, {}, "", "", "asc", 1, 50)["rows"]), 0)    # xem trước không ghi
        outbound.goods_import_bulk(self.conn, rows, dry_run=False, default_ship="VNOUT00004")
        outbound.goods_import_bulk(self.conn, [{"row": 2, "ship": "VNOUT00004", "ma_hang": "D", "qty": 1, "don_gia": 1}], dry_run=False, mode="append")
        g = outbound.goods_query(self.conn, {}, "", "Amount", "desc", 1, 50)
        self.assertEqual((g["total"], g["summary"]["ships"]), (4, 2))
        self.assertEqual([r[1] for r in g["rows"]][:2], ["A", "B"])                              # sắp theo Amount: 25 USD, 20 SEK
        self.assertEqual(g["list_lines"][0], 2)                                                  # link sang 2 dòng List shipment cùng Shipment#
        self.assertEqual(outbound.goods_query(self.conn, {"Currency": "sek"}, "", "", "asc", 1, 50)["total"], 1)
        self.assertEqual(outbound.goods_query(self.conn, {"Amount": ">=20"}, "", "", "asc", 1, 50)["total"], 2)
        self.assertEqual(outbound.goods_query(self.conn, {}, "cap", "", "asc", 1, 50)["total"], 1)
        gid = g["ids"][0]
        outbound.goods_line_save(self.conn, gid, "VNOUT00004", {"ma_hang": "A2", "qty": "20", "don_gia": "2"})
        self.assertEqual(outbound.goods_line_get(self.conn, gid)["amount"], 40.0)
        outbound.goods_line_delete(self.conn, gid)
        self.assertEqual(outbound.goods_query(self.conn, {}, "", "", "asc", 1, 50)["total"], 3)

    def test_invoice_scan_apply_and_inv_form(self):
        import datetime as dt
        import openpyxl
        import invoice_scan
        import outbound
        import outbound_forms
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Invoice (Polarium AB)"
        ws["C2"] = "COMMERCIAL INVOICE"
        ws["E5"], ws["H5"] = "Invoice Number", "Invoice Date"
        ws["E6"], ws["H6"] = "TO101009999", dt.datetime(2026, 10, 27)
        ws["E9"], ws["H9"] = "Polarium SO/TO number:", "TO101009999"
        ws["E10"], ws["H10"] = "Polarium Ref (Shipment ID):", "VNPLINT00099"
        ws["E13"], ws["H13"] = "INCOTERMS:", "FCA"
        ws["E15"] = "Mode of Shipment: AIR"
        ws["B13"] = "Deliver to:"
        ws["B14"] = "ACME 3PL\nStockholm"
        for c, h in zip("BCDEFGHI", ["Item NO", "ITEM", "DESCRIPTION", "QTY", "UNIT", "Country of origin", "Unit Price (SEK)", "Total Amount (SEK)"]):
            ws[f"{c}25"] = h
        for r, row in enumerate([(1, "880-00022", "Stacking\xa0Material", 100, "EA", "VN", 17.118, 1711.8), (2, "800-00182", "Cell", 5, "EA", "VN", 2, 10)], start=26):
            for c, v in zip("BCDEFGHI", row):
                ws[f"{c}{r}"] = v
        ws["D50"], ws["H50"], ws["I50"] = "Total Value of goods:", "SEK", 1721.8
        p = self.dir / "inv.xlsx"
        wb.save(p)
        res = invoice_scan.scan(p, "inv.xlsx")
        h = res["header"]
        self.assertEqual((h["invoice_no"], h["invoice_date"], h["so_to"], h["polarium_ref"], h["incoterms"], h["mode"]),
                         ("TO101009999", "2026-10-27", "TO101009999", "VNPLINT00099", "FCA", "AIR"))
        self.assertIn("ACME 3PL", h["deliver_to"])
        self.assertEqual((res["currency"], res["total"], res["sum_lines"], res["total_match"], len(res["lines"])), ("SEK", 1721.8, 1721.8, True, 2))
        self.assertEqual(res["lines"][0]["ten_hang"], "Stacking Material")
        # lưu: dòng hàng gắn theo Invoice; quét lại cùng Invoice thay dòng cũ, Invoice khác giữ nguyên
        outbound.goods_save(self.conn, "VNPLINT00099", [{"ma_hang": "MANUAL", "ten_hang": "Nhập tay", "qty": 1}])
        r1 = outbound.apply_invoice(self.conn, "VNPLINT00099", h, res["lines"], res["currency"], res["total"], "replace", p, "inv.xlsx", "xlsx")
        r2 = outbound.apply_invoice(self.conn, "VNPLINT00099", h, res["lines"], res["currency"], res["total"], "replace", p, "inv.xlsx", "xlsx")
        self.assertEqual((r1["lines"], r2["id"] == r1["id"]), (2, True))
        goods = outbound.goods_get(self.conn, "VNPLINT00099")
        self.assertEqual(sorted(g["ma_hang"] for g in goods), ["800-00182", "880-00022", "MANUAL"])
        self.assertEqual({g["inv_no"] for g in goods if g["ma_hang"] != "MANUAL"}, {"TO101009999"})
        self.assertEqual(outbound.find_invoice(self.conn, "to101009999")["polarium_ref"], "VNPLINT00099")
        self.assertEqual(outbound.search(self.conn, "TO101009999")["results"][0]["ship"], "VNPLINT00099")
        # biểu mẫu INV: số Invoice / ngày / Incoterms lấy từ Invoice đã quét (trước đây #REF!)
        wbf, _ = outbound_forms.build(self.conn, "VNPLINT00099", "TO101009999", ["INV"])
        sh = wbf["INV"]
        self.assertEqual((sh["E6"].value, sh["H12"].value), ("TO101009999", "FCA"))
        self.assertEqual(sh["I6"].value.date(), dt.date(2026, 10, 27))

    def test_full_export_one_file_all_shipments_with_lines(self):
        import io
        import openpyxl
        import outbound
        outbound.import_workbook(self.conn, self.src, {"source": "Source_Data", "tracking": "Tracking"})
        outbound.goods_save(self.conn, "VNOUT00900", [{"ma_hang": "A-1", "ten_hang": "Cell", "qty": 10, "don_gia": 2.5, "ccy": "EUR", "xuat_xu": "VN"},
                                                      {"ma_hang": "A-2", "ten_hang": "Cap", "qty": 4, "don_gia": 5, "ccy": "EUR"}])
        wb = openpyxl.load_workbook(io.BytesIO(outbound.full_export_xlsx(self.conn)))
        self.assertEqual(wb.sheetnames, ["Chi tiết shipment", "Shipment", "Dòng hàng", "Invoice", "Tracking"])
        rows = list(wb["Chi tiết shipment"].iter_rows(values_only=True))
        head = rows[0]
        self.assertEqual(head[:6], ("Shipment ID", "Dòng #", "Nguồn dòng hàng", "Item (mã hàng)", "Description", "Qty"))
        by_ship = {}
        for r in rows[1:]:
            by_ship.setdefault(r[0], []).append(r)
        self.assertEqual(len(by_ship["VNOUT00900"]), 2)                       # 2 dòng hàng -> 2 dòng, lặp thông tin shipment
        self.assertEqual(len(by_ship["VNOUT00901"]), 1)                       # shipment chưa có dòng hàng vẫn xuất hiện
        r1 = by_ship["VNOUT00900"][0]
        self.assertEqual((r1[3], r1[5], r1[8], r1[9], r1[10]), ("A-1", 10, 2.5, "EUR", 25.0))
        self.assertIn("Booking No.", head)
        self.assertEqual(r1[head.index("Booking No.")], "BK123")
        flt = list(outbound.full_export_xlsx(self.conn, {"Shipment ID": "=VNOUT00901"}) and [1])
        wb2 = openpyxl.load_workbook(io.BytesIO(outbound.full_export_xlsx(self.conn, {"Shipment ID": "=VNOUT00901"})))
        self.assertEqual({r[0] for r in list(wb2["Chi tiết shipment"].iter_rows(values_only=True))[1:]}, {"VNOUT00901"})
        self.assertEqual(len(list(wb2["Tracking"].iter_rows(values_only=True))), 1)   # chỉ còn dòng tiêu đề: Tracking của shipment khác bị lọc

    def test_full_export_from_list_shipment_uses_list_items_as_fallback(self):
        import io
        import openpyxl
        import outbound
        wbk = openpyxl.Workbook()
        ws = wbk.active
        ws.title = "List shipment"
        for i, h in enumerate(["No.", "Separate ID#", "Orders", "Produced", "Quantity", "Row updated", "Shipment#", "Item", "Description"], start=1):
            ws.cell(2, i, h)
        ws.append([1, "VNOUT00900-1", "SO1", None, 5, None, "VNOUT00900", "800-1", "Cell"])
        ws.append([2, "VNOUT00900-2", "SO1", None, 3, None, "VNOUT00900", "800-2", "Cap"])
        ws.append([3, "VNOUT00901-1", "SO2", None, 7, None, "VNOUT00901", "800-3", "Pack"])
        ws.append([4, "VNPLINT00099-1", "SO3", None, 2, None, "VNPLINT00099", "800-4", "Only in list"])
        p = self.dir / "ls3.xlsx"
        wbk.save(p)
        outbound.import_workbook(self.conn, self.src, {"source": "Source_Data"})
        outbound.import_workbook(self.conn, p, {"list": "List shipment"})
        outbound.goods_save(self.conn, "VNOUT00901", [{"ma_hang": "INV-1", "ten_hang": "From invoice", "qty": 7, "don_gia": 3, "ccy": "SEK"}])
        out = openpyxl.load_workbook(io.BytesIO(outbound.full_export_xlsx(self.conn, {"Shipment#": "=VNOUT00900|VNOUT00901|VNPLINT00099"}, "", "list")))
        rows = list(out["Chi tiết shipment"].iter_rows(values_only=True))
        got = [(r[0], r[2], r[3]) for r in rows[1:]]
        self.assertEqual(got, [("VNOUT00900", "List shipment", "800-1"), ("VNOUT00900", "List shipment", "800-2"), ("VNOUT00901", "Invoice", "INV-1"),
                               ("VNPLINT00099", "List shipment", "800-4")])                  # shipment chỉ có ở List shipment vẫn xuất; có Invoice thì ưu tiên Invoice
        self.assertEqual(len(list(out["List shipment"].iter_rows(values_only=True))) - 1, 4)
        # xuất từ chi tiết 1 dòng List shipment: theo ship của dòng đó, không cần bộ lọc
        lid = self.conn.execute("SELECT id FROM ob_rows WHERE tbl='list' AND json_extract(cells,'$.G')='VNOUT00901'").fetchone()
        if lid:
            sh = outbound.get_row(self.conn, lid["id"])["ship"]
            one = openpyxl.load_workbook(io.BytesIO(outbound.full_export_xlsx(self.conn, None, "", "list", sh)))
            r1 = list(one["Chi tiết shipment"].iter_rows(values_only=True))[1:]
            self.assertEqual({r[0] for r in r1}, {"VNOUT00901"})
            self.assertEqual(r1[0][3], "INV-1")                                                      # dòng hàng Invoice có trong file
            self.assertEqual(len(list(one["List shipment"].iter_rows(values_only=True))) - 1, 1)

    def test_dgd_red_frame_drawing_is_reattached(self):
        import io
        import zipfile
        import outbound
        import outbound_forms
        import xml.etree.ElementTree as ET
        outbound.import_workbook(self.conn, self.src, {"source": "Source_Data"})
        wb, _ = outbound_forms.build(self.conn, "VNOUT00900", None, ["DGD WH"])
        z = zipfile.ZipFile(io.BytesIO(outbound_forms.to_bytes(wb)))
        self.assertIn("xl/drawings/drawing_orig1.xml", z.namelist())
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf8")
        self.assertIn('<drawing r:id="rIdOrigDrw"/>', sheet)
        ET.fromstring(sheet)
        self.assertIn("drawing_orig1", z.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf8"))
        # các biểu mẫu khác (không có hình vẽ) không bị ghép gì thêm
        wb2, _ = outbound_forms.build(self.conn, "VNOUT00900", None, ["VGM"])
        self.assertNotIn("xl/drawings/drawing_orig1.xml", zipfile.ZipFile(io.BytesIO(outbound_forms.to_bytes(wb2))).namelist())


class RvcFormTest(unittest.TestCase):
    def build(self, n_codes):
        import rvc_form
        with tempfile.TemporaryDirectory() as tmp:
            conn = db.connect(Path(tmp) / "v.db")
            for i in range(n_codes):
                code = f"110-{i:05d}"
                conn.execute("""INSERT INTO import_lines(seq,so_tk,ngay_dk,loai_hinh,ma_hang,ma_key,qty,dedupe,trang_thai,ccy,ty_gia,don_gia_tt,nuoc_xx,ma_hs,ten_hang)
                                VALUES(?,?,?,'E11',?,?,100,?,?,'USD',25000,50000,'CHINA','85076090','Pin')""",
                             (i, f"T{i}", "2026-08-10", code, code, f"d{i}", "Đã thông quan"))
            conn.commit()
            lines = [{"ma_hang": f"110-{i:05d}", "qty": 2} for i in range(n_codes)]
            e = importer.create_event(conn, "CO", "L1", {"so_tk_xuat": "3089", "ngay_tk": "01/10/2026", "ma_tp": "FG1", "qty_lo": 1}, lines)
            wb, d = rvc_form.build(conn, e, db.get_settings(conn))
            ws = wb["RVC"]
            res = (ws.max_row, ws.cell(rvc_form.TOTAL_ROW + n_codes - 102, 7).value, ws["G17"].value, ws["H17"].value, ws["F17"].value,
                   sorted(str(m) for m in ws.merged_cells.ranges if m.min_row >= 20), ws["A7"].value)
            html_ = rvc_form.to_html(ws)
            conn.close()
            return res, html_

    def test_template_resizes_and_fta_columns_are_left_blank(self):
        for n in (3, 102, 105):
            (max_row, total_g, g17, h17, f17, tail, a7), html_ = self.build(n)
            delta = n - 102
            self.assertEqual(max_row, 129 + delta)
            self.assertEqual(tail, sorted([f"H{127 + delta}:M{127 + delta}", f"H{128 + delta}:M{128 + delta}", f"H{129 + delta}:M{129 + delta}"]))
            self.assertIsNone(g17)
            self.assertIsNone(h17)                              # cột có / không có xuất xứ FTA: điền tay
            self.assertAlmostEqual(f17, 50000 / 25000)          # đơn giá CIF = giá tính thuế / tỷ giá
            self.assertEqual(total_g, f"=SUM(G17:G{rvc_form_total(n) - 1})")
            self.assertIn("3089", a7)
            self.assertIn("<table", html_)


def rvc_form_total(n):
    import rvc_form
    return rvc_form.TOTAL_ROW + n - 102


class BomTest(unittest.TestCase):
    def test_bom_upload_create_lot_and_stock_hint(self):
        import boms
        import openpyxl
        with tempfile.TemporaryDirectory() as tmp:
            conn = db.connect(Path(tmp) / "b.db")
            mk(conn, "A", "2026-08-01", "E11", "110-00001", 100, 1)
            conn.commit()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Mã thành phẩm", "Tên hàng hóa", "Mã HS 6 số", "Mã hàng (NVL)", "Item description", "Qty used in FG", "Unit"])
            ws.append(["FG1", "Pin", "850760", "110-00001", "Cell", 15, "EACH"])
            ws.append(["FG1", None, None, "999-00009", "Không có tồn", 2, "EACH"])
            ws.append(["FG1", None, None, "110-00002", "Sai", -1, "EACH"])
            p = Path(tmp) / "bom.xlsx"
            wb.save(p)
            rows = boms.parse_table(p)
            r = boms.apply_table(conn, rows, dry_run=True)
            self.assertEqual((r["n_errors"], r["n_warnings"]), (1, 1))
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM boms").fetchone()[0], 0)
            boms.apply_table(conn, rows, dry_run=False)
            bid = conn.execute("SELECT id FROM boms WHERE ma_tp='FG1'").fetchone()[0]
            d = boms.detail(conn, bid, db.get_settings(conn))
            self.assertEqual(len(d["lines"]), 2)
            self.assertEqual(d["lines"][0]["lo_toi_da"], 6)                   # 100 / 15
            self.assertEqual(d["max_lots"], 0)                                 # mã 999-00009 không có tồn
            e = importer.create_event(conn, "CO", "L1", {"qty_lo": 3}, boms.lines_for_event(conn, bid), multiply_by_lot=True)
            need = {r["ma_hang"]: r["need"] for r in conn.execute("SELECT * FROM event_lines WHERE event_id=?", (e,))}
            self.assertEqual(need, {"110-00001": 45.0, "999-00009": 6.0})
            conn.close()

if __name__ == "__main__":
    unittest.main()
