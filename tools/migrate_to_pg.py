"""Chuyển toàn bộ dữ liệu từ SQLite (data/nvl.db) sang PostgreSQL.

Chạy:  python tools/migrate_to_pg.py "postgresql://user:pass@localhost:5432/nvl"  [đường_dẫn_nvl.db]
Dữ liệu cũ trong PostgreSQL (nếu có) bị thay thế. File SQLite không bị sửa."""
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pgcompat  # noqa: E402

ORDER = ["settings", "meta", "fx_rates", "dup_decisions", "import_log", "import_lines", "events", "event_lines", "allocations",
         "boms", "bom_lines", "ob_tables", "ob_rows", "ob_files", "ob_goods", "ob_invoices"]


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    dsn = sys.argv[1]
    src_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent.parent / "data" / "nvl.db"
    import db
    src = sqlite3.connect(f"file:{src_path}?mode=ro", uri=True)
    src.row_factory = sqlite3.Row
    pg = pgcompat.PGConn(dsn)
    pg.executescript(db.SCHEMA)
    for t in reversed(ORDER):
        pg.execute(f"DELETE FROM {t}")
    pg.commit()
    for t in ORDER:
        cols = [r[1] for r in src.execute(f"PRAGMA table_info({t})")]
        if not cols:
            continue
        pcols = {r["name"] for r in pg.execute("PRAGMA table_info(%s)" % t)}
        use = [c for c in cols if c in pcols]
        ptype = {r["column_name"]: r["data_type"] for r in pg.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name=%s AND table_schema=current_schema()".replace("%s", "?"), (t,))}
        nontext = [ptype[c] not in ("text", "character varying") for c in use]
        n = 0
        with pg._pg.cursor() as cur:
            with cur.copy(f"COPY {t} ({','.join(use)}) FROM STDIN") as cp:
                for r in src.execute(f"SELECT {','.join(use)} FROM {t}"):
                    cp.write_row(tuple(None if (nt and v == "") else v for v, nt in zip(r, nontext)))
                    n += 1
        if "id" in use:
            pg.execute(f"SELECT setval(pg_get_serial_sequence('{t}','id'), COALESCE((SELECT MAX(id) FROM {t}),0)+1, false)")
        pg.commit()
        print(f"{t}: {n} dòng")
    pg.close()


if __name__ == "__main__":
    main()
