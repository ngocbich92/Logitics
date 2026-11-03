"""Bộ tính công thức Excel tối thiểu (đủ cho các biểu mẫu outbound): VLOOKUP, IF, IFERROR, AND, ROUND, LEFT, MID, TEXT, ...

Dùng để điền các biểu mẫu Excel gốc (SI Form, DGD, VGM, CMA...) từ dữ liệu trong hệ thống mà không cần mở Excel.
Ô trống được biểu diễn bằng None (so sánh bằng "" và 0 như Excel).
"""
import datetime as dt
import math
import re
from decimal import ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP, ROUND_UP, Decimal

from openpyxl.utils import column_index_from_string, get_column_letter


class XlError(Exception):
    def __init__(self, code="#VALUE!"):
        super().__init__(code)
        self.code = code


# ------------------------------------------------------------------------------------------ tokenizer
_TOKEN = re.compile(r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<err>\#(?:REF|N/A|VALUE|NAME|DIV/0|NUM|NULL)[!?]?)
  | (?P<ref>(?:(?:'[^']+'|[A-Za-z_][\w\.]*)!)?(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}))
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<func>[A-Za-z][A-Za-z0-9\.]*(?=\())
  | (?P<bool>TRUE|FALSE)
  | (?P<op><=|>=|<>|[=<>&+\-*/^%(),])
""", re.X)


def tokenize(src):
    pos, out = 0, []
    while pos < len(src):
        m = _TOKEN.match(src, pos)
        if not m:
            raise XlError("#NAME?")
        pos = m.end()
        kind = m.lastgroup
        if kind != "ws":
            out.append((kind, m.group()))
    return out


# --------------------------------------------------------------------------------------------- parser
class Parser:
    """cmp (= <> < > <= >=) < concat (&) < additive < multiplicative < power < unary < percent."""

    def __init__(self, src):
        self.t, self.i = tokenize(src), 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, v):
        if self.take()[1] != v:
            raise XlError("#NAME?")

    def parse(self):
        node = self.cmp()
        if self.i != len(self.t):
            raise XlError("#NAME?")
        return node

    def cmp(self):
        n = self.concat()
        while self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            n = ("bin", op, n, self.concat())
        return n

    def concat(self):
        n = self.add()
        while self.peek()[1] == "&":
            self.take()
            n = ("bin", "&", n, self.add())
        return n

    def add(self):
        n = self.mul()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            n = ("bin", op, n, self.mul())
        return n

    def mul(self):
        n = self.pow()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            n = ("bin", op, n, self.pow())
        return n

    def pow(self):
        n = self.unary()
        while self.peek()[1] == "^":
            self.take()
            n = ("bin", "^", n, self.unary())
        return n

    def unary(self):
        if self.peek()[1] in ("-", "+"):
            op = self.take()[1]
            n = self.unary()
            return ("neg", n) if op == "-" else n
        return self.postfix()

    def postfix(self):
        n = self.atom()
        while self.peek()[1] == "%":
            self.take()
            n = ("bin", "/", n, ("num", 100.0))
        return n

    def atom(self):
        kind, v = self.take()
        if kind == "num":
            return ("num", float(v))
        if kind == "str":
            return ("str", v[1:-1].replace('""', '"'))
        if kind == "bool":
            return ("bool", v == "TRUE")
        if kind == "err":
            return ("err", v)
        if kind == "ref":
            return parse_ref(v)
        if kind == "func":
            self.expect("(")
            args = []
            if self.peek()[1] != ")":
                while True:
                    args.append(self.cmp() if self.peek()[1] not in (",", ")") else ("empty",))
                    if self.peek()[1] == ",":
                        self.take()
                        continue
                    break
            self.expect(")")
            return ("func", v.upper(), args)
        if v == "(":
            n = self.cmp()
            self.expect(")")
            return n
        raise XlError("#NAME?")


def parse_ref(text):
    sheet = None
    if "!" in text:
        sheet, text = text.rsplit("!", 1)
        sheet = sheet.strip("'")
    text = text.replace("$", "")
    a, _, b = text.partition(":")
    pa = re.fullmatch(r"([A-Z]{1,3})(\d*)", a)
    pb = re.fullmatch(r"([A-Z]{1,3})(\d*)", b) if b else None
    c1, r1 = column_index_from_string(pa.group(1)), (int(pa.group(2)) if pa.group(2) else None)
    if not b:
        return ("cell", sheet, c1, r1)
    c2, r2 = column_index_from_string(pb.group(1)), (int(pb.group(2)) if pb.group(2) else None)
    return ("range", sheet, c1, r1, c2, r2)


_cache = {}


def parse(formula):
    if formula not in _cache:
        _cache[formula] = Parser(formula[1:] if formula.startswith("=") else formula).parse()
    return _cache[formula]


# ------------------------------------------------------------------------------------------ evaluator
def to_num(v):
    if v is None:
        return 0.0
    if v == "":
        raise XlError("#VALUE!")  # chuỗi rỗng trong phép tính -> #VALUE! (khác ô trống = 0)
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, dt.datetime):
        return (v - dt.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(v, dt.date):
        return float((v - dt.date(1899, 12, 30)).days)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        raise XlError("#VALUE!")


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return f"{v:.15g}"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute) else v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


def truthy(v):
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        raise XlError("#VALUE!")
    return bool(to_num(v)) if v is not None else False


def _cmp(a, b):
    """So sánh kiểu Excel: ô trống = "" = 0; chữ không phân biệt hoa thường; số < chữ."""
    if a is None and b is None:
        return 0
    if a is None:
        a = "" if isinstance(b, str) else 0.0
    if b is None:
        b = "" if isinstance(a, str) else 0.0
    if isinstance(a, (dt.date, dt.datetime)):
        a = to_num(a)
    if isinstance(b, (dt.date, dt.datetime)):
        b = to_num(b)
    if isinstance(a, bool):
        a = float(a)
    if isinstance(b, bool):
        b = float(b)
    if isinstance(a, str) and isinstance(b, str):
        a, b = a.lower(), b.lower()
    elif isinstance(a, str):
        return 1
    elif isinstance(b, str):
        return -1
    return (a > b) - (a < b)


def _round(x, n, mode):
    q = Decimal(1).scaleb(-int(n))
    d = Decimal(repr(float(x)))
    return float(d.quantize(q, rounding=mode)) if n >= 0 else float((d / Decimal(10) ** -int(n)).quantize(Decimal(1), rounding=mode) * Decimal(10) ** -int(n))


def _fmt_text(v, fmt):
    if isinstance(v, (dt.date, dt.datetime)) or (isinstance(v, float) and "y" in fmt.lower()):
        d = v if isinstance(v, (dt.date, dt.datetime)) else dt.datetime(1899, 12, 30) + dt.timedelta(days=v)
        out = fmt
        for a, b in (("yyyy", "%Y"), ("yy", "%y"), ("mmmm", "%B"), ("mmm", "%b"), ("mm", "%m"), ("dd", "%d"), ("hh", "%H"), ("ss", "%S")):
            out = re.sub(a, b, out, flags=re.I) if a not in ("mm",) else out.replace("mm", "%m")
        return d.strftime(out)
    n = to_num(v)
    decimals = len(fmt.split(".")[1].rstrip("%")) if "." in fmt else 0
    s = f"{n:,.{decimals}f}" if "," in fmt else f"{n:.{decimals}f}"
    return s


class Evaluator:
    """ctx: đối tượng cung cấp
         ctx.cell_value(sheet, col, row) -> giá trị ô (đã tính công thức nếu là công thức)
         ctx.lookup_row(sheet, key)      -> dict {chữ_cột: giá_trị} của dòng có khoá (cột B) = key, hoặc None nếu không phải bảng dữ liệu
         ctx.today()                      -> ngày hôm nay"""

    def __init__(self, ctx):
        self.ctx = ctx

    def ev(self, node, sheet):
        k = node[0]
        if k == "num" or k == "str" or k == "bool":
            return node[1]
        if k == "empty":
            return None
        if k == "err":
            raise XlError(node[1])
        if k == "cell":
            _, sh, c, r = node
            return self.ctx.cell_value(sh or sheet, c, r)
        if k == "range":
            raise XlError("#VALUE!")
        if k == "neg":
            return -to_num(self.ev(node[1], sheet))
        if k == "bin":
            return self.binop(node, sheet)
        if k == "func":
            return self.call(node[1], node[2], sheet)
        raise XlError("#NAME?")

    def binop(self, node, sheet):
        _, op, a, b = node
        x, y = self.ev(a, sheet), self.ev(b, sheet)
        if op == "&":
            return to_text(x) + to_text(y)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            c = _cmp(x, y)
            return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]
        x, y = to_num(x), to_num(y)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise XlError("#DIV/0!")
            return x / y
        if op == "^":
            return x ** y
        raise XlError("#NAME?")

    def values_of(self, node, sheet):
        """Giá trị phẳng của một đối số (ô / vùng / biểu thức) — dùng cho SUM, MAX..."""
        if node[0] == "range":
            _, sh, c1, r1, c2, r2 = node
            if r1 is None or r2 is None:
                raise XlError("#VALUE!")
            return [self.ctx.cell_value(sh or sheet, c, r) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        return [self.ev(node, sheet)]

    def call(self, name, args, sheet):
        ev = lambda i: self.ev(args[i], sheet)
        n = len(args)
        if name == "IF":
            return ev(1) if truthy(ev(0)) else (ev(2) if n > 2 else False)
        if name == "IFERROR":
            try:
                v = ev(0)
                return v
            except XlError:
                return ev(1)
        if name == "AND":
            return all(truthy(ev(i)) for i in range(n))
        if name == "OR":
            return any(truthy(ev(i)) for i in range(n))
        if name == "NOT":
            return not truthy(ev(0))
        if name == "VLOOKUP":
            key = ev(0)
            rng = args[1]
            if rng[0] != "range":
                raise XlError("#REF!")
            _, sh, c1, _r1, c2, _r2 = rng
            idx = int(to_num(ev(2)))
            if key is None or key == "":
                raise XlError("#N/A")
            row = self.ctx.lookup_row(sh or sheet, to_text(key))
            if row is None:
                raise XlError("#N/A")
            col = c1 + idx - 1
            if idx < 1 or col > c2:
                raise XlError("#REF!")
            return row.get(get_column_letter(col))
        if name == "MATCH":
            key = ev(0)
            rng = args[1]
            if rng[0] != "range":
                raise XlError("#REF!")
            vals = self.values_of(rng, sheet)
            for i, v in enumerate(vals, start=1):
                if _cmp(v, key) == 0:
                    return float(i)
            raise XlError("#N/A")
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            v = ev(0)
            x, d = to_num(v), int(to_num(ev(1))) if n > 1 else 0
            mode = {"ROUND": ROUND_HALF_UP, "ROUNDUP": ROUND_UP, "ROUNDDOWN": ROUND_DOWN}[name]
            return _round(x, d, mode)
        if name == "LEFT":
            return to_text(ev(0))[: int(to_num(ev(1))) if n > 1 else 1]
        if name == "RIGHT":
            s, k = to_text(ev(0)), int(to_num(ev(1))) if n > 1 else 1
            return s[-k:] if k else ""
        if name == "MID":
            s, st, ln = to_text(ev(0)), int(to_num(ev(1))), int(to_num(ev(2)))
            return s[st - 1: st - 1 + ln]
        if name == "LEN":
            return float(len(to_text(ev(0))))
        if name in ("UPPER", "LOWER", "TRIM"):
            s = to_text(ev(0))
            return {"UPPER": s.upper(), "LOWER": s.lower(), "TRIM": " ".join(s.split())}[name]
        if name == "CONCATENATE":
            return "".join(to_text(ev(i)) for i in range(n))
        if name == "TEXT":
            return _fmt_text(ev(0), to_text(ev(1)))
        if name == "TODAY":
            return dt.datetime.combine(self.ctx.today(), dt.time())
        if name == "NOW":
            return dt.datetime.combine(self.ctx.today(), dt.time())
        if name == "CHAR":
            return chr(int(to_num(ev(0))))
        if name == "VALUE":
            return to_num(ev(0))
        if name == "ABS":
            return abs(to_num(ev(0)))
        if name == "ISBLANK":
            return ev(0) is None
        if name == "ISNUMBER":
            return isinstance(ev(0), (int, float)) and not isinstance(ev(0), bool)
        if name in ("SUM", "MAX", "MIN", "AVERAGE"):
            nums = []
            for a in args:
                for v in self.values_of(a, sheet):
                    if isinstance(v, bool) or v is None or isinstance(v, str):
                        continue
                    nums.append(to_num(v))
            if name == "SUM":
                return sum(nums)
            if not nums:
                return 0.0
            return {"MAX": max, "MIN": min, "AVERAGE": lambda x: sum(x) / len(x)}[name](nums)
        raise XlError("#NAME?")
