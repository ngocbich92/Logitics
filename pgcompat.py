"""Lớp tương thích PostgreSQL: giữ nguyên giao diện của sqlite3 (execute / ? / lastrowid / Row...) để toàn bộ code còn lại chạy được trên PostgreSQL.

Cấu hình kết nối: biến môi trường NVL_PG (DSN) hoặc file data/db_config.json {"backend": "postgres", "dsn": "postgresql://user:pass@host:5432/nvl"}.
Không có cấu hình -> hệ thống dùng SQLite như cũ."""
import hashlib
import json
import os
import re
import threading
from pathlib import Path

CONFIG = Path(__file__).parent / "data" / "db_config.json"

ID_TABLES = {"import_lines", "events", "event_lines", "allocations", "boms", "bom_lines", "ob_rows", "ob_files", "ob_goods", "ob_invoices", "import_log"}


def dsn():
    """DSN PostgreSQL nếu được cấu hình, ngược lại None."""
    if os.environ.get("NVL_PG"):
        return os.environ["NVL_PG"]
    try:
        c = json.loads(CONFIG.read_text(encoding="utf-8-sig"))
        if c.get("backend") == "postgres" and c.get("dsn"):
            return c["dsn"]
    except (OSError, ValueError):
        pass
    return None


class Row:
    """Giống sqlite3.Row: r[0], r['ten'], dict(r), r.keys()."""
    __slots__ = ("_v", "_k")

    def __init__(self, keys, values):
        self._k, self._v = keys, values

    def __getitem__(self, i):
        if isinstance(i, str):
            try:
                return self._v[self._k.index(i)]
            except ValueError:
                raise IndexError("No item with that key")
        return self._v[i]

    def keys(self):
        return list(self._k)

    def __iter__(self):
        return iter(self._v)

    def __len__(self):
        return len(self._v)

    def __repr__(self):
        return f"Row({dict(zip(self._k, self._v))})"


def _row_factory(cursor):
    cols = [d.name for d in (cursor.description or [])]
    return lambda values: Row(cols, tuple(values))


def translate(sql):
    """Đổi cú pháp SQLite -> PostgreSQL."""
    out, i, n = [], 0, len(sql)
    while i < n:
        c = sql[i]
        if c == "'":  # chuỗi: giữ nguyên, chỉ nhân đôi dấu %
            j = i + 1
            while j < n:
                if sql[j] == "'" and j + 1 < n and sql[j + 1] == "'":
                    j += 2
                    continue
                if sql[j] == "'":
                    break
                j += 1
            out.append(sql[i:j + 1].replace("%", "%%"))
            i = j + 1
        elif c == "?":
            out.append("%s")
            i += 1
        elif c == "%":
            out.append("%%")
            i += 1
        else:
            out.append(c)
            i += 1
    s = "".join(out)
    s = re.sub(r"datetime\('now',\s*'localtime'\)", "to_char(now(),'YYYY-MM-DD HH24:MI:SS')", s)
    s = re.sub(r"\bLIKE\b", "ILIKE", s)
    s = re.sub(r"\bIS NOT %s", "IS DISTINCT FROM %s", s)
    m = re.match(r"\s*INSERT OR IGNORE INTO", s, re.I)
    if m:
        s = re.sub(r"INSERT OR IGNORE INTO", "INSERT INTO", s, count=1, flags=re.I).rstrip().rstrip(";") + " ON CONFLICT DO NOTHING"
    return s


def pg_schema(schema_sql):
    s = re.sub(r"INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY", schema_sql)
    s = re.sub(r"\bREAL\b", "DOUBLE PRECISION", s)
    s = re.sub(r"DEFAULT \(datetime\('now','localtime'\)\)", "DEFAULT to_char(now(),'YYYY-MM-DD HH24:MI:SS')", s)
    return s


FUNCS = """
CREATE EXTENSION IF NOT EXISTS unaccent;
CREATE OR REPLACE FUNCTION fold(t text) RETURNS text LANGUAGE sql STABLE AS
  $$ SELECT lower(unaccent(replace(replace(COALESCE(t,''),'đ','d'),'Đ','D'))) $$;
CREATE OR REPLACE FUNCTION instr(a text, b text) RETURNS integer LANGUAGE sql IMMUTABLE AS $$ SELECT strpos(a,b) $$;
CREATE OR REPLACE FUNCTION round(x double precision, n integer) RETURNS double precision LANGUAGE sql IMMUTABLE AS $$ SELECT round(x::numeric, n)::double precision $$;
CREATE OR REPLACE FUNCTION json_extract(j text, p text) RETURNS text LANGUAGE sql IMMUTABLE AS
  $$ SELECT CASE WHEN j IS NULL OR j='' THEN NULL ELSE (j::jsonb) ->> regexp_replace(substr(p,3),'^"(.*)"$','\\1') END $$;
CREATE OR REPLACE FUNCTION usd_rate(dd text) RETURNS double precision LANGUAGE plpgsql STABLE AS $$ DECLARE r double precision; BEGIN SELECT usd INTO r FROM pg_temp.usd_day WHERE pg_temp.usd_day.d = dd; RETURN r; END $$;
CREATE OR REPLACE FUNCTION json_num(j text, p text) RETURNS double precision LANGUAGE sql IMMUTABLE AS
  $$ SELECT CASE WHEN json_extract(j,p) ~ '^-?[0-9]+([.][0-9]+)?([eE][-+]?[0-9]+)?$' THEN json_extract(j,p)::double precision END $$;
"""


class Cursor:
    def __init__(self, cur, lastrowid=None):
        self._c, self.lastrowid = cur, lastrowid

    @property
    def rowcount(self):
        return self._c.rowcount

    @property
    def description(self):
        return self._c.description

    def fetchone(self):
        return self._c.fetchone() if self._c.description else None

    def fetchall(self):
        return self._c.fetchall() if self._c.description else []

    def __iter__(self):
        return iter(self.fetchall())


class _Null:
    lastrowid = None
    rowcount = 0
    description = None

    def fetchone(self):
        return None

    def fetchall(self):
        return []

    def __iter__(self):
        return iter(())


_ready = set()
_lock = threading.Lock()


class PGConn:
    def __init__(self, dsn_, schema=None, schema_sql=None):
        import psycopg
        self._pg = psycopg.connect(dsn_, autocommit=False)
        self.schema = schema or "public"
        if self.schema != "public":
            self._raw(f"CREATE SCHEMA IF NOT EXISTS {self.schema}")
        self._raw(f"SET search_path TO {self.schema}, public")
        key = (dsn_, self.schema)
        with _lock:
            if key not in _ready:
                self._pg.execute("SELECT pg_advisory_xact_lock(7654321)")
                for stmt in FUNCS.split(";\n"):
                    if stmt.strip():
                        self._pg.execute(stmt)
                self._pg.commit()
                _ready.add(key)
        self.is_pg = True

    def _raw(self, sql):
        self._pg.execute(sql)
        self._pg.commit()

    # --- giao diện sqlite3
    def execute(self, sql, params=None):
        st = sql.strip()
        up = st.upper()
        if up.startswith("PRAGMA"):
            m = re.match(r"PRAGMA\s+table_info\((\w+)\)", st, re.I)
            if m:
                cur = self._pg.cursor(row_factory=_row_factory)
                cur.execute("SELECT column_name AS name FROM information_schema.columns WHERE table_name=%s AND table_schema=current_schema() ORDER BY ordinal_position", (m.group(1),))
                return Cursor(cur)
            return _Null()
        q = translate(st)
        rid = False
        m = re.match(r"INSERT\s+INTO\s+(\w+)", st, re.I)
        if m and m.group(1).lower() in ID_TABLES and "RETURNING" not in up:
            q += " RETURNING id"
            rid = True
        cur = self._pg.cursor(row_factory=_row_factory)
        cur.execute(q, tuple(params) if params is not None else ())
        if rid:
            r = cur.fetchone()
            return Cursor(cur, r[0] if r else None)
        return Cursor(cur)

    def executemany(self, sql, seq):
        cur = self._pg.cursor(row_factory=_row_factory)
        cur.executemany(translate(sql.strip()), [tuple(p) for p in seq])
        return Cursor(cur)

    def executescript(self, script):
        for stmt in re.sub(r"--[^\n]*", "", pg_schema(script)).split(";"):
            if stmt.strip():
                self._pg.execute(stmt)
        self._pg.commit()

    def create_function(self, *a, **k):  # hàm đã có sẵn trong PostgreSQL
        pass

    def commit(self):
        self._pg.commit()

    def rollback(self):
        self._pg.rollback()

    def close(self):
        try:
            self._pg.close()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, et, ev, tb):
        if et is None:
            self._pg.commit()
        else:
            self._pg.rollback()
        return False


def schema_for(path):
    """Test: mỗi đường dẫn DB tạm -> một schema riêng."""
    return "t_" + hashlib.md5(str(path).encode()).hexdigest()[:12]
