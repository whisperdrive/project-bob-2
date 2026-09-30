"""Runtime for Python overlays compiled from Excel (xlcompile.py): Excel values, operators and functions.

A compiled overlay is a module of row functions, one per line item: DCF_7(c) returns the value of DCF!r7 in
column c. They run against a Book, which holds:
  inputs      the overlay's constants (its assumptions), from the workbook
  overrides   values a person sets for a scenario (on constants or on formula cells)
  feed / ext  values read from outside the overlay: the client model's sheets in the same workbook, or
              another workbook through an external link ([1]Sheet!A1)
  cached      the workbook's saved values, used to validate and as the fallback for a circular reference
Values follow Excel: floats (dates are serial numbers), str, bool, None for blank, XLError, 2-D lists for
arrays and Rng for references (evaluated lazily, so INDEX(range, MATCH(...)) computes one cell, not the range).
"""
import math
import statistics
import re
from bisect import bisect_right
from collections import Counter
from datetime import date, datetime, timedelta
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal

RUNTIME = 2  # bump when a function's result changes (2: TEXT dates, half-up rounding): validations made with an
             # earlier runtime are shown as out of date until the overlay is built again
EPOCH = date(1899, 12, 30)
_MISS = object()


class XLError:
    __slots__ = ("code",)

    def __init__(self, code):
        self.code = code

    def __repr__(self):
        return self.code

    def __eq__(self, other):
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)


ERR = {c: XLError(c) for c in ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA",
                              "#SPILL!", "#CALC!")}
NA, VALUE, DIV0, NUM, REF, NAME = (ERR[c] for c in ("#N/A", "#VALUE!", "#DIV/0!", "#NUM!", "#REF!", "#NAME?"))


class Missing:
    """An empty argument, as in IF(a,,b) or OFFSET(A1,-1,)."""
    def __repr__(self):
        return "MISSING"


MISSING = Missing()
_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$")
_TIME = re.compile(r"^(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?$")  # a time-formatted number, as calamine writes it


def from_db(v):
    """A value as stored in model.db -> an Excel value: ISO dates -> serials, '#N/A' -> XLError."""
    if isinstance(v, str):
        m = _DATE.match(v)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                s = (date(y, mo, d) - EPOCH).days
            except ValueError:
                return v
            if m.group(4):
                s += (int(m.group(4)) * 3600 + int(m.group(5)) * 60 + int(m.group(6) or 0)) / 86400
            return float(s)
        m = _TIME.match(v)
        if m:
            return (int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))) / 86400
        if v in ERR:
            return ERR[v]
        return v
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return float(v)
    return v


def serial(d: date) -> float:
    return float((d - EPOCH).days)


def to_date(s) -> date:
    return EPOCH + timedelta(days=int(math.floor(s)))


# ---- references and arrays ----------------------------------------------------------------------------------

class Rng:
    """A reference to a block of cells: src "" = this workbook, or an external link index."""
    __slots__ = ("book", "src", "sheet", "r1", "c1", "r2", "c2")

    def __init__(self, book, src, sheet, r1, c1, r2, c2):
        if r2 < r1:
            r1, r2 = r2, r1
        if c2 < c1:
            c1, c2 = c2, c1
        self.book, self.src, self.sheet, self.r1, self.c1, self.r2, self.c2 = book, src, sheet, r1, c1, r2, c2

    @property
    def shape(self):
        return self.r2 - self.r1 + 1, self.c2 - self.c1 + 1

    def cell(self, i: int, j: int):
        return self.book.get(self.src, self.sheet, self.r1 + i, self.c1 + j)

    def values(self):
        key = (self.src, self.sheet, self.r1, self.c1, self.r2, self.c2)
        cache = self.book.range_cache
        v = cache.get(key)
        if v is None or self.book.rec is not None:  # while a cell's reads are recorded, read each cell again
            get = self.book.get
            v = [[get(self.src, self.sheet, r, c) for c in range(self.c1, self.c2 + 1)] for r in range(self.r1, self.r2 + 1)]
            cache[key] = v
        return v

    def __repr__(self):
        return f"Rng({self.src}{self.sheet}!R{self.r1}C{self.c1}:R{self.r2}C{self.c2})"


def grid(x):
    """Any value -> 2-D list."""
    if isinstance(x, Rng):
        return x.values()
    if isinstance(x, list):
        return x
    return [[x]]


def flat(x):
    for row in grid(x):
        yield from row


def first(x):
    """Rng / array -> its top-left value (Excel's single-value use of a reference)."""
    if isinstance(x, Rng):
        return x.cell(0, 0)
    if isinstance(x, list):
        return x[0][0] if x and x[0] else None
    return x


def is_array(x):
    return isinstance(x, (Rng, list))


def broadcast(fn, a, b):
    """Elementwise fn over arrays, Excel-style: a 1-row or 1-column side is repeated, missing cells are #N/A."""
    ga, gb = grid(a), grid(b)
    ra, ca, rb, cb = len(ga), len(ga[0]), len(gb), len(gb[0])
    rows, cols = max(ra, rb), max(ca, cb)

    def at(g, r, c, nr, nc):
        if nr == 1:
            r = 0
        if nc == 1:
            c = 0
        return g[r][c] if r < nr and c < nc else NA
    return [[fn(at(ga, r, c, ra, ca), at(gb, r, c, rb, cb)) for c in range(cols)] for r in range(rows)]


def elementwise(fn, a):
    return [[fn(v) for v in row] for row in grid(a)]


# ---- coercion -----------------------------------------------------------------------------------------------

def num(v):
    """Arithmetic operand: blank 0, TRUE 1, numeric text as a number, other text #VALUE!."""
    if v is None or isinstance(v, Missing):
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, float):
        return v
    if isinstance(v, int):
        return float(v)
    if isinstance(v, XLError):
        return v
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return VALUE
        t = s.replace(",", "")
        try:
            return float(t[:-1]) / 100.0 if t.endswith("%") else float(t)
        except ValueError:
            return from_db(s) if _DATE.match(s) else VALUE
    return VALUE


def text(v) -> str:
    if v is None or isinstance(v, Missing):
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return f"{v:.15g}"
    return str(v)


def truth(v):
    if isinstance(v, bool):
        return v
    if v is None or isinstance(v, Missing):
        return False
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, XLError):
        return v
    if isinstance(v, str):
        u = v.strip().upper()
        if u in ("TRUE", "FALSE"):
            return u == "TRUE"
        return VALUE
    return VALUE


def _rank(v):
    """Excel's ordering across types: numbers < text < logicals; blank compares as 0 or ''."""
    if isinstance(v, bool):
        return 2, v
    if isinstance(v, (int, float)):
        return 0, v
    if isinstance(v, str):
        return 1, v.lower()
    return 0, 0.0


def _cmp(a, b):
    if a is None or isinstance(a, Missing):
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None or isinstance(b, Missing):
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)
    ra, rb = _rank(a), _rank(b)
    return (ra > rb) - (ra < rb)


# ---- operators ----------------------------------------------------------------------------------------------

def _arith(op):
    def scalar(a, b):
        if isinstance(a, XLError):
            return a
        if isinstance(b, XLError):
            return b
        x, y = num(a), num(b)
        if isinstance(x, XLError):
            return x
        if isinstance(y, XLError):
            return y
        return op(x, y)

    def fn(a, b):
        if is_array(a) or is_array(b):
            return broadcast(scalar, a, b)
        return scalar(a, b)
    return fn


def _div(x, y):
    return DIV0 if y == 0 else x / y


def _pow(x, y):
    try:
        if x == 0 and y < 0:
            return DIV0
        r = x ** y
        return NUM if isinstance(r, complex) else float(r)
    except (OverflowError, ZeroDivisionError):
        return NUM


def _compare(test):
    def scalar(a, b):
        if isinstance(a, XLError):
            return a
        if isinstance(b, XLError):
            return b
        return test(_cmp(a, b))

    def fn(a, b):
        if is_array(a) or is_array(b):
            return broadcast(scalar, a, b)
        return scalar(a, b)
    return fn


def _concat_scalar(a, b):
    if isinstance(a, XLError):
        return a
    if isinstance(b, XLError):
        return b
    return text(a) + text(b)


_DATE_FMT = r'(?i)(?:"[^"]*"|\\.|[dmy]+|[\s\-/.,:\'()])+'  # d, m, y codes, separators and literal text only
_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _date_text(d: date, f: str) -> str:
    """A date in an Excel number format of day, month and year codes ("dd mmm yyyy", "d/mm/yy", "mmmm yyyy")."""
    out = []
    for tok in re.findall(r'"[^"]*"|\\.|y+|m+|d+|[^"\\ymd]+', f, re.I):
        t = tok.lower()
        if tok.startswith('"'):
            out.append(tok[1:-1])
        elif tok.startswith("\\"):
            out.append(tok[1:])
        elif t[0] == "y":
            out.append(f"{d.year:04d}" if len(t) > 2 else f"{d.year % 100:02d}")
        elif t[0] == "m":
            name = _MONTHS[d.month - 1]
            out.append({1: str(d.month), 2: f"{d.month:02d}", 3: name[:3], 4: name}.get(len(t), name[0]))
        elif t[0] == "d":
            day = _DAYS[d.weekday()]
            out.append({1: str(d.day), 2: f"{d.day:02d}", 3: day[:3]}.get(len(t), day))
        else:
            out.append(tok)
    return "".join(out)


class _XL:
    """Excel operators and functions, used by compiled overlays as xl.add(...), xl.SUM(...)."""
    ERR = ERR
    MISSING = MISSING
    quirks = Counter()  # Excel behaviours reproduced deliberately (see XIRR)

    add = staticmethod(_arith(lambda x, y: x + y))
    sub = staticmethod(_arith(lambda x, y: x - y))
    mul = staticmethod(_arith(lambda x, y: x * y))
    div = staticmethod(_arith(_div))
    pow = staticmethod(_arith(_pow))
    eq = staticmethod(_compare(lambda c: c == 0))
    ne = staticmethod(_compare(lambda c: c != 0))
    lt = staticmethod(_compare(lambda c: c < 0))
    gt = staticmethod(_compare(lambda c: c > 0))
    le = staticmethod(_compare(lambda c: c <= 0))
    ge = staticmethod(_compare(lambda c: c >= 0))

    @staticmethod
    def concat(a, b):
        if is_array(a) or is_array(b):
            return broadcast(_concat_scalar, a, b)
        return _concat_scalar(a, b)

    @staticmethod
    def neg(a):
        if is_array(a):
            return elementwise(lambda v: _XL.neg(v), a)
        x = num(a)
        return x if isinstance(x, XLError) else -x

    @staticmethod
    def pct(a):
        return _XL.div(a, 100.0)

    @staticmethod
    def array(rows):
        return rows

    # ---- logic --------------------------------------------------------------------------------------------
    @staticmethod
    def IF(cond, yes, no=None):
        if is_array(cond):
            return elementwise(lambda c: _pick(truth(c), yes, no), cond)
        return _pick(truth(cond), yes, no)

    @staticmethod
    def IFERROR(value, alt):
        v = value()
        if is_array(v):
            return elementwise(lambda x: alt() if isinstance(x, XLError) else x, v)
        return alt() if isinstance(v, XLError) else _unmiss(v)

    @staticmethod
    def IFNA(value, alt):
        v = value()
        if is_array(v):
            return elementwise(lambda x: alt() if x == NA else x, v)
        return alt() if v == NA else _unmiss(v)

    @staticmethod
    def CHOOSE(k, *options):
        if is_array(k):
            return elementwise(lambda x: _XL.CHOOSE(x, *options), k)
        i = num(k)
        if isinstance(i, XLError):
            return i
        i = int(i)
        if i < 1 or i > len(options):
            return VALUE
        return _unmiss(options[i - 1]())

    @staticmethod
    def AND(*args):
        return _logical(args, all)

    @staticmethod
    def OR(*args):
        return _logical(args, any)

    @staticmethod
    def XOR(*args):
        return _logical(args, lambda xs: sum(bool(x) for x in xs) % 2 == 1)

    @staticmethod
    def NOT(a):
        if is_array(a):
            return elementwise(lambda v: _XL.NOT(v), a)
        t = truth(a)
        return t if isinstance(t, XLError) else (not t)

    @staticmethod
    def TRUE():
        return True

    @staticmethod
    def FALSE():
        return False

    # ---- information --------------------------------------------------------------------------------------
    @staticmethod
    def ISBLANK(v):
        return _info(v, lambda x: x is None)

    @staticmethod
    def ISNUMBER(v):
        return _info(v, lambda x: isinstance(x, float) and not isinstance(x, bool))

    @staticmethod
    def ISTEXT(v):
        return _info(v, lambda x: isinstance(x, str))

    @staticmethod
    def ISNONTEXT(v):
        return _info(v, lambda x: not isinstance(x, str))

    @staticmethod
    def ISLOGICAL(v):
        return _info(v, lambda x: isinstance(x, bool))

    @staticmethod
    def ISERROR(v):
        return _info(v, lambda x: isinstance(x, XLError))

    @staticmethod
    def ISERR(v):
        return _info(v, lambda x: isinstance(x, XLError) and x != NA)

    @staticmethod
    def ISNA(v):
        return _info(v, lambda x: x == NA)

    @staticmethod
    def NA():
        return NA

    @staticmethod
    def N(v):
        v = first(v)
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, XLError):
            return v
        return 0.0

    # ---- aggregates ---------------------------------------------------------------------------------------
    @staticmethod
    def SUM(*args):
        return _agg(args, lambda xs: math.fsum(xs), empty=0.0)

    @staticmethod
    def PRODUCT(*args):
        return _agg(args, lambda xs: math.prod(xs), empty=0.0)

    @staticmethod
    def MAX(*args):
        return _agg(args, max, empty=0.0)

    @staticmethod
    def MIN(*args):
        return _agg(args, min, empty=0.0)

    @staticmethod
    def AVERAGE(*args):
        return _agg(args, lambda xs: math.fsum(xs) / len(xs), empty=DIV0)

    @staticmethod
    def COUNT(*args):
        n = 0
        for a in args:
            if is_array(a):
                n += sum(1 for v in flat(a) if isinstance(v, float) and not isinstance(v, bool))
            elif not isinstance(num(a), XLError) and not isinstance(a, Missing):
                n += 1
        return float(n)

    @staticmethod
    def COUNTA(*args):
        n = 0
        for a in args:
            n += sum(1 for v in flat(a) if v is not None) if is_array(a) else (0 if isinstance(a, Missing) else 1)
        return float(n)

    @staticmethod
    def COUNTBLANK(rng):
        return float(sum(1 for v in flat(rng) if v is None or v == ""))

    @staticmethod
    def SUMPRODUCT(*arrays):
        gs = [grid(a) for a in arrays]
        shape = (len(gs[0]), len(gs[0][0]))
        if any((len(g), len(g[0])) != shape for g in gs):
            return VALUE
        total = []
        for r in range(shape[0]):
            for c in range(shape[1]):
                p = 1.0
                for g in gs:
                    v = g[r][c]
                    if isinstance(v, XLError):
                        return v
                    p *= v if isinstance(v, float) and not isinstance(v, bool) else 0.0
                total.append(p)
        return math.fsum(total)

    @staticmethod
    def SUMSQ(*args):
        return _agg(args, lambda xs: math.fsum(x * x for x in xs), empty=0.0)

    @staticmethod
    def MMULT(a, b):
        ga, gb = grid(a), grid(b)
        if len(ga[0]) != len(gb):
            return VALUE
        out = []
        for row in ga:
            r = []
            for j in range(len(gb[0])):
                s = []
                for k, x in enumerate(row):
                    y = gb[k][j]
                    if isinstance(x, XLError):
                        return x
                    if isinstance(y, XLError):
                        return y
                    if not isinstance(x, float) or not isinstance(y, float) or isinstance(x, bool) or isinstance(y, bool):
                        return VALUE
                    s.append(x * y)
                r.append(math.fsum(s))
            out.append(r)
        return out

    # ---- conditional aggregates ---------------------------------------------------------------------------
    @staticmethod
    def SUMIFS(sum_rng, *pairs):
        return _ifs(sum_rng, pairs, lambda xs: math.fsum(xs), 0.0)

    @staticmethod
    def SUMIF(rng, crit, sum_rng=None):
        return _ifs(sum_rng if sum_rng is not None and not isinstance(sum_rng, Missing) else rng, (rng, crit),
                    lambda xs: math.fsum(xs), 0.0)

    @staticmethod
    def AVERAGEIFS(avg_rng, *pairs):
        return _ifs(avg_rng, pairs, lambda xs: math.fsum(xs) / len(xs), DIV0)

    @staticmethod
    def AVERAGEIF(rng, crit, avg_rng=None):
        return _ifs(avg_rng if avg_rng is not None and not isinstance(avg_rng, Missing) else rng, (rng, crit),
                    lambda xs: math.fsum(xs) / len(xs), DIV0)

    @staticmethod
    def COUNTIFS(*pairs):
        g = grid(pairs[0])
        mask = _mask(pairs)
        if isinstance(mask, XLError):
            return mask
        return float(sum(1 for r in range(len(g)) for c in range(len(g[0])) if mask[r][c]))

    @staticmethod
    def COUNTIF(rng, crit):
        return _XL.COUNTIFS(rng, crit)

    @staticmethod
    def MAXIFS(rng, *pairs):
        return _ifs(rng, pairs, max, 0.0)

    @staticmethod
    def MINIFS(rng, *pairs):
        return _ifs(rng, pairs, min, 0.0)

    # ---- lookup and reference -----------------------------------------------------------------------------
    @staticmethod
    def INDEX(arr, r=MISSING, c=MISSING):
        res = _index(arr, r, c)
        if isinstance(res, Rng):
            return res if res.shape != (1, 1) else res.cell(0, 0)
        return res

    @staticmethod
    def INDEX_REF(arr, r=MISSING, c=MISSING):
        return _index(arr, r, c)

    @staticmethod
    def MATCH(value, arr, kind=1.0):
        value = first(value) if is_array(value) else value
        if isinstance(value, XLError):
            return value
        vals = [v for v in flat(arr)]
        k = num(kind) if not isinstance(kind, Missing) else 1.0
        if isinstance(k, XLError):
            return k
        if k == 0:
            if isinstance(value, str) and re.search(r"[*?~]", value):
                pat = _wild(value)
                for i, v in enumerate(vals):
                    if isinstance(v, str) and pat.fullmatch(v):
                        return float(i + 1)
                return NA
            if value is None:
                return NA
            want = _rank(value)
            for i, v in enumerate(vals):
                if v is not None and not isinstance(v, XLError) and _rank(v)[0] == want[0] and _rank(v)[1] == want[1]:
                    return float(i + 1)
            return NA
        return _approx(vals, value, descending=k < 0)

    @staticmethod
    def VLOOKUP(value, table, col, approx=True):
        g = grid(table)
        col = int(num(col)) if not isinstance(num(col), XLError) else 0
        if col < 1 or col > len(g[0]):
            return REF
        keys = [row[0] for row in g]
        i = _lookup_pos(value, keys, approx)
        return i if isinstance(i, XLError) else _unmiss(g[i][col - 1])

    @staticmethod
    def HLOOKUP(value, table, row, approx=True):
        g = grid(table)
        row = int(num(row)) if not isinstance(num(row), XLError) else 0
        if row < 1 or row > len(g):
            return REF
        i = _lookup_pos(value, g[0], approx)
        return i if isinstance(i, XLError) else _unmiss(g[row - 1][i])

    @staticmethod
    def LOOKUP(value, lookup, result=MISSING):
        keys = list(flat(lookup))
        i = _approx(keys, value)
        if isinstance(i, XLError):
            return i
        res = list(flat(result)) if not isinstance(result, Missing) else keys
        return _unmiss(res[int(i) - 1]) if int(i) - 1 < len(res) else NA

    @staticmethod
    def OFFSET(ref, rows, cols, height=MISSING, width=MISSING):
        if not isinstance(ref, Rng):
            return VALUE
        vals = [0.0 if isinstance(x, Missing) else num(first(x)) for x in (rows, cols)]
        if any(isinstance(v, XLError) for v in vals):
            return next(v for v in vals if isinstance(v, XLError))
        h = ref.shape[0] if isinstance(height, Missing) else num(first(height))
        w = ref.shape[1] if isinstance(width, Missing) else num(first(width))
        if isinstance(h, XLError) or isinstance(w, XLError):
            return VALUE
        r1, c1 = ref.r1 + int(vals[0]), ref.c1 + int(vals[1])
        if r1 < 1 or c1 < 1 or int(h) < 1 or int(w) < 1:
            return REF
        return Rng(ref.book, ref.src, ref.sheet, r1, c1, r1 + int(h) - 1, c1 + int(w) - 1)

    @staticmethod
    def RANGE(a, b):
        """a:b between two references (e.g. INDEX(...):INDEX(...))."""
        if not isinstance(a, Rng) or not isinstance(b, Rng) or (a.src, a.sheet) != (b.src, b.sheet):
            return REF
        return Rng(a.book, a.src, a.sheet, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))

    @staticmethod
    def ISECT(a, b):
        if not isinstance(a, Rng) or not isinstance(b, Rng):
            return VALUE
        r1, c1, r2, c2 = max(a.r1, b.r1), max(a.c1, b.c1), min(a.r2, b.r2), min(a.c2, b.c2)
        if r1 > r2 or c1 > c2 or (a.src, a.sheet) != (b.src, b.sheet):
            return ERR["#NULL!"]
        return Rng(a.book, a.src, a.sheet, r1, c1, r2, c2)

    @staticmethod
    def ROWS(x):
        return float(x.shape[0] if isinstance(x, Rng) else len(grid(x)))

    @staticmethod
    def COLUMNS(x):
        return float(x.shape[1] if isinstance(x, Rng) else len(grid(x)[0]))

    @staticmethod
    def ROW(x):
        return float(x.r1) if isinstance(x, Rng) else VALUE

    @staticmethod
    def COLUMN(x):
        return float(x.c1) if isinstance(x, Rng) else VALUE

    @staticmethod
    def TRANSPOSE(x):
        g = grid(x)
        return [list(r) for r in zip(*g)]

    # ---- maths --------------------------------------------------------------------------------------------
    @staticmethod
    def ABS(x):
        return _math1(x, abs)

    @staticmethod
    def INT(x):
        return _math1(x, lambda v: float(math.floor(v)))

    @staticmethod
    def TRUNC(x, d=0.0):
        return _round(x, d, ROUND_DOWN)

    @staticmethod
    def SIGN(x):
        return _math1(x, lambda v: float((v > 0) - (v < 0)))

    @staticmethod
    def SQRT(x):
        return _math1(x, lambda v: math.sqrt(v) if v >= 0 else NUM)

    @staticmethod
    def EXP(x):
        return _math1(x, lambda v: math.exp(v))

    @staticmethod
    def LN(x):
        return _math1(x, lambda v: math.log(v) if v > 0 else NUM)

    @staticmethod
    def LOG10(x):
        return _math1(x, lambda v: math.log10(v) if v > 0 else NUM)

    @staticmethod
    def LOG(x, base=10.0):
        return _math2(x, base, lambda v, b: math.log(v, b) if v > 0 and b > 0 and b != 1 else NUM)

    @staticmethod
    def POWER(x, y):
        return _XL.pow(x, y)

    @staticmethod
    def MOD(x, y):
        return _math2(x, y, lambda a, b: DIV0 if b == 0 else a - b * math.floor(a / b))

    @staticmethod
    def ROUND(x, d=0.0):
        return _round(x, d, ROUND_HALF_UP)

    @staticmethod
    def ROUNDUP(x, d=0.0):
        return _round(x, d, ROUND_UP)

    @staticmethod
    def ROUNDDOWN(x, d=0.0):
        return _round(x, d, ROUND_DOWN)

    @staticmethod
    def CEILING(x, sig=1.0):
        return _math2(x, sig, lambda v, s: 0.0 if s == 0 else math.ceil(v / s) * s)

    @staticmethod
    def FLOOR(x, sig=1.0):
        return _math2(x, sig, lambda v, s: DIV0 if s == 0 else math.floor(v / s) * s)

    @staticmethod
    def PI():
        return math.pi

    # ---- dates --------------------------------------------------------------------------------------------
    @staticmethod
    def DATE(y, m, d):
        vals = [num(first(v)) for v in (y, m, d)]
        if any(isinstance(v, XLError) for v in vals):
            return next(v for v in vals if isinstance(v, XLError))
        y, m, d = (int(v) for v in vals)
        if y < 1900:
            y += 1900
        y += (m - 1) // 12
        m = (m - 1) % 12 + 1
        return serial(date(y, m, 1)) + d - 1

    @staticmethod
    def YEAR(s):
        return _date1(s, lambda d: float(d.year))

    @staticmethod
    def MONTH(s):
        return _date1(s, lambda d: float(d.month))

    @staticmethod
    def DAY(s):
        return _date1(s, lambda d: float(d.day))

    @staticmethod
    def EOMONTH(s, months):
        return _date2(s, months, lambda d, m: serial(_add_months(d, int(m), end=True)))

    @staticmethod
    def EDATE(s, months):
        return _date2(s, months, lambda d, m: serial(_add_months(d, int(m))))

    @staticmethod
    def YEARFRAC(a, b, basis=0.0):
        vals = [num(first(v)) for v in (a, b, 0.0 if isinstance(basis, Missing) else basis)]
        if any(isinstance(v, XLError) for v in vals):
            return next(v for v in vals if isinstance(v, XLError))
        return _yearfrac(to_date(vals[0]), to_date(vals[1]), int(vals[2]))

    @staticmethod
    def DAYS(end, start):
        return _math2(end, start, lambda e, s: float(math.floor(e) - math.floor(s)))

    @staticmethod
    def TODAY():
        return serial(date.today())

    @staticmethod
    def NOW():
        n = datetime.now()
        return serial(n.date()) + (n.hour * 3600 + n.minute * 60 + n.second) / 86400

    # ---- finance ------------------------------------------------------------------------------------------
    @staticmethod
    def NPV(rate, *values):
        r = num(first(rate))
        if isinstance(r, XLError):
            return r
        xs = []
        for a in values:
            xs += [v for v in flat(a) if isinstance(v, float) and not isinstance(v, bool)] if is_array(a) else [num(a)]
        if any(isinstance(x, XLError) for x in xs):
            return next(x for x in xs if isinstance(x, XLError))
        return math.fsum(x / (1 + r) ** (i + 1) for i, x in enumerate(xs))

    @staticmethod
    def XNPV(rate, values, dates):
        r = num(first(rate))
        vs, ds = list(flat(values)), list(flat(dates))
        if isinstance(r, XLError):
            return r
        if len(vs) != len(ds) or not vs:
            return NUM
        for x in vs + ds:
            if isinstance(x, XLError):
                return x
        if any(not isinstance(x, float) for x in vs + ds):
            return VALUE
        d0 = math.floor(ds[0])
        if any(math.floor(d) < d0 for d in ds):
            return NUM
        return math.fsum(v / (1 + r) ** ((math.floor(d) - d0) / 365.0) for v, d in zip(vs, ds))

    @staticmethod
    def XIRR(values, dates, guess=0.1):
        vs, ds = list(flat(values)), list(flat(dates))
        for x in vs + ds:
            if isinstance(x, XLError):
                return x
        if len(vs) != len(ds) or any(not isinstance(x, float) for x in vs + ds):
            return VALUE
        if not (any(v > 0 for v in vs) and any(v < 0 for v in vs)):
            return NUM
        if vs[0] == 0:
            # Excel quirk, reproduced on purpose: with a zero first cash flow XIRR stops at once and returns
            # 2.98E-09 (2^-25 / 10), not the rate. Workbooks carry that value, so the overlay must too.
            _XL.quirks["XIRR with a zero first cash flow returns 2.98E-09 in Excel"] += 1
            return 2.9802322387695314e-09
        d0 = math.floor(ds[0])
        t = [(math.floor(d) - d0) / 365.0 for d in ds]
        g = num(first(guess)) if not isinstance(guess, Missing) else 0.1
        return _irr(vs, t, 0.1 if isinstance(g, XLError) else g)

    @staticmethod
    def IRR(values, guess=0.1):
        vs = [v for v in flat(values) if isinstance(v, float) and not isinstance(v, bool)]
        if not (any(v > 0 for v in vs) and any(v < 0 for v in vs)):
            return NUM
        g = num(first(guess)) if not isinstance(guess, Missing) else 0.1
        return _irr(vs, [float(i) for i in range(len(vs))], 0.1 if isinstance(g, XLError) else g)

    # ---- text ---------------------------------------------------------------------------------------------
    @staticmethod
    def LEN(s):
        return _text1(s, lambda t: float(len(t)))

    @staticmethod
    def LEFT(s, n=1.0):
        return _text2(s, n, lambda t, k: t[:int(k)])

    @staticmethod
    def RIGHT(s, n=1.0):
        return _text2(s, n, lambda t, k: t[len(t) - int(k):] if int(k) else "")

    @staticmethod
    def MID(s, start, n):
        s, a, k = first(s), num(first(start)), num(first(n))
        if isinstance(s, XLError):
            return s
        if isinstance(a, XLError) or isinstance(k, XLError):
            return VALUE
        t = text(s)
        return t[int(a) - 1:int(a) - 1 + int(k)]

    @staticmethod
    def FIND(needle, hay, start=1.0):
        return _find(needle, hay, start, False)

    @staticmethod
    def SEARCH(needle, hay, start=1.0):
        return _find(needle, hay, start, True)

    @staticmethod
    def UPPER(s):
        return _text1(s, str.upper)

    @staticmethod
    def LOWER(s):
        return _text1(s, str.lower)

    @staticmethod
    def TRIM(s):
        return _text1(s, lambda t: re.sub(r" +", " ", t.strip(" ")))

    @staticmethod
    def CONCATENATE(*args):
        out = ""
        for a in args:
            v = first(a) if is_array(a) else a
            if isinstance(v, XLError):
                return v
            out += text(v)
        return out

    @staticmethod
    def CONCAT(*args):
        out = ""
        for a in args:
            for v in (flat(a) if is_array(a) else [a]):
                if isinstance(v, XLError):
                    return v
                out += text(v)
        return out

    @staticmethod
    def SUBSTITUTE(s, old, new, which=MISSING):
        s, old, new = (first(x) for x in (s, old, new))
        for v in (s, old, new):
            if isinstance(v, XLError):
                return v
        s, old, new = text(s), text(old), text(new)
        if isinstance(which, Missing):
            return s.replace(old, new) if old else s
        k, pos = int(num(which)), -1
        for _ in range(k):
            pos = s.find(old, pos + 1)
            if pos < 0:
                return s
        return s[:pos] + new + s[pos + len(old):]

    @staticmethod
    def REPT(s, n):
        return _text2(s, n, lambda t, k: t * int(k))

    @staticmethod
    def VALUE(s):
        v = first(s)
        return v if isinstance(v, (float, XLError)) else num(text(v))

    @staticmethod
    def TEXT(v, fmt):
        v, f = first(v), text(first(fmt))
        if isinstance(v, XLError):
            return v
        x = num(v)
        if isinstance(x, XLError):
            return text(v)
        m = re.fullmatch(r"[#,]*0(?:\.(0+))?(%?)", f.replace(",", "") if f.count(",") <= 1 else f)
        if m:
            d = len(m.group(1) or "")
            q = (Decimal(repr(x)) * (100 if m.group(2) else 1)).quantize(Decimal(1).scaleb(-d), ROUND_HALF_UP)  # 0.5 up
            if m.group(2):
                return f"{q:.{d}f}%"
            return f"{q:,.{d}f}" if "," in f else f"{q:.{d}f}"
        if re.fullmatch(_DATE_FMT, f) and re.search(r"[dmy]", f, re.I):  # a date format (no time: m would be minutes)
            return _date_text(to_date(x), f)
        return text(v)

    @staticmethod
    def CELL(info, ref=MISSING):
        k = text(first(info)).lower()
        if isinstance(ref, Rng):
            if k == "row":
                return float(ref.r1)
            if k in ("col", "column"):
                return float(ref.c1)
            if k == "contents":
                return ref.cell(0, 0)
            if k == "filename":  # models use it for the sheet's own name (text after "]"); the path isn't known
                return f"[workbook]{ref.sheet}"
        return VALUE  # "address", "format" etc. depend on Excel's window and formatting state



    # ---- newer and less common functions (Excel 365 ones arrive without their _xlfn. prefix) -------------------
    @staticmethod
    def XMATCH(value, arr, mode=0.0, search=1.0):
        value = first(value) if is_array(value) else value
        if isinstance(value, XLError):
            return value
        vals = list(flat(arr))
        m = 0 if isinstance(mode, Missing) else int(num(mode))
        back = not isinstance(search, Missing) and num(search) < 0
        order = list(range(len(vals)))[::-1] if back else list(range(len(vals)))
        if m == 2:  # wildcards
            pat = _wild(str(value))
            hit = next((i for i in order if isinstance(vals[i], str) and pat.fullmatch(vals[i])), None)
            return float(hit + 1) if hit is not None else NA
        want = _rank(value)
        exact = next((i for i in order if vals[i] is not None and not isinstance(vals[i], XLError)
                      and _rank(vals[i]) == want), None)
        if exact is not None or m == 0:
            return float(exact + 1) if exact is not None else NA
        best = None  # -1: the largest value below it; 1: the smallest above it
        for i in order:
            v = vals[i]
            if v is None or isinstance(v, XLError) or _rank(v)[0] != want[0]:
                continue
            c = _cmp(v, value)
            if (m == -1 and c < 0 and (best is None or _cmp(v, vals[best]) > 0)) or \
               (m == 1 and c > 0 and (best is None or _cmp(v, vals[best]) < 0)):
                best = i
        return float(best + 1) if best is not None else NA

    @staticmethod
    def XLOOKUP(value, lookup, result, missing=MISSING, mode=0.0, search=1.0):
        i = _XL.XMATCH(value, lookup, mode, search)
        if isinstance(i, XLError):
            return i if isinstance(missing, Missing) else missing
        i = int(i) - 1
        g, r = grid(lookup), grid(result)
        if len(g) > 1 or len(g[0]) == 1:  # a column: the matching row of the results
            if i >= len(r):
                return REF
            row = r[i]
            return _unmiss(row[0]) if len(row) == 1 else [list(row)]
        if i >= len(r[0]):
            return REF
        col = [[x[i]] for x in r]
        return _unmiss(col[0][0]) if len(col) == 1 else col

    @staticmethod
    def LARGE(arr, k):
        xs = sorted((v for v in flat(arr) if isinstance(v, float) and not isinstance(v, bool)), reverse=True)
        k = num(k)
        return k if isinstance(k, XLError) else (xs[int(k) - 1] if 1 <= int(k) <= len(xs) else NUM)

    @staticmethod
    def SMALL(arr, k):
        xs = sorted(v for v in flat(arr) if isinstance(v, float) and not isinstance(v, bool))
        k = num(k)
        return k if isinstance(k, XLError) else (xs[int(k) - 1] if 1 <= int(k) <= len(xs) else NUM)

    @staticmethod
    def MEDIAN(*args):
        return _agg(args, lambda xs: sorted(xs)[len(xs) // 2] if len(xs) % 2 else
                    (sorted(xs)[len(xs) // 2 - 1] + sorted(xs)[len(xs) // 2]) / 2, empty=NUM)

    @staticmethod
    def RANK(value, arr, order=0.0):
        v = num(value)
        if isinstance(v, XLError):
            return v
        xs = [x for x in flat(arr) if isinstance(x, float) and not isinstance(x, bool)]
        if v not in xs:
            return NA
        asc = not isinstance(order, Missing) and num(order) != 0
        return float(1 + sum(1 for x in xs if (x < v if asc else x > v)))

    RANK_EQ = RANK

    @staticmethod
    def STDEV(*args):
        return _agg(args, lambda xs: statistics.stdev(xs) if len(xs) > 1 else DIV0, empty=DIV0)

    STDEV_S = STDEV

    @staticmethod
    def STDEV_P(*args):
        return _agg(args, lambda xs: statistics.pstdev(xs), empty=DIV0)

    @staticmethod
    def VAR(*args):
        return _agg(args, lambda xs: statistics.variance(xs) if len(xs) > 1 else DIV0, empty=DIV0)

    VAR_S = VAR

    @staticmethod
    def TEXTJOIN(delim, skip_empty, *args):
        skip = truth(skip_empty) is True
        parts = [text(v) for a in args for v in (flat(a) if is_array(a) else [a]) if not isinstance(v, Missing)]
        bad = next((v for a in args for v in (flat(a) if is_array(a) else [a]) if isinstance(v, XLError)), None)
        if bad is not None:
            return bad
        return text(delim).join(x for x in parts if x != "" or not skip)

    @staticmethod
    def _annuity(rate, nper, pmt, pv, fv, when):
        r, n, pm, p, f, w = (num(x) if not isinstance(x, Missing) else 0.0 for x in (rate, nper, pmt, pv, fv, when))
        for x in (r, n, pm, p, f, w):
            if isinstance(x, XLError):
                raise _Err(x)
        return r, n, pm, p, f, 1.0 if w else 0.0

    @staticmethod
    def PMT(rate, nper, pv, fv=MISSING, when=MISSING):
        try:
            r, n, _, p, f, w = _XL._annuity(rate, nper, 0.0, pv, fv, when)
        except _Err as e:
            return e.err
        if n == 0:
            return NUM
        if r == 0:
            return -(p + f) / n
        g = (1 + r) ** n
        return -(r * (p * g + f)) / ((1 + r * w) * (g - 1))

    @staticmethod
    def PV(rate, nper, pmt, fv=MISSING, when=MISSING):
        try:
            r, n, pm, _, f, w = _XL._annuity(rate, nper, pmt, 0.0, fv, when)
        except _Err as e:
            return e.err
        if r == 0:
            return -(f + pm * n)
        g = (1 + r) ** n
        return -(f + pm * (1 + r * w) * (g - 1) / r) / g

    @staticmethod
    def FV(rate, nper, pmt, pv=MISSING, when=MISSING):
        try:
            r, n, pm, p, _, w = _XL._annuity(rate, nper, pmt, pv, 0.0, when)
        except _Err as e:
            return e.err
        if r == 0:
            return -(p + pm * n)
        g = (1 + r) ** n
        return -(p * g + pm * (1 + r * w) * (g - 1) / r)

    @staticmethod
    def WEEKDAY(s, kind=1.0):
        k = 1 if isinstance(kind, Missing) else int(num(kind))
        return _date1(s, lambda d: float((d.isoweekday() % 7) + 1 if k == 1 else d.isoweekday() if k == 2 else d.weekday()))

    @staticmethod
    def DATEDIF(a, b, unit):
        x, y, u = num(a), num(b), text(unit).upper()
        if isinstance(x, XLError) or isinstance(y, XLError):
            return x if isinstance(x, XLError) else y
        if y < x:
            return NUM
        d1, d2 = to_date(x), to_date(y)
        months = (d2.year - d1.year) * 12 + d2.month - d1.month - (d2.day < d1.day)
        return {"D": float(int(y) - int(x)), "M": float(months), "Y": float(months // 12),
                "YM": float(months % 12)}.get(u, NUM)

    @staticmethod
    def MROUND(x, m):
        a, b = num(x), num(m)
        if isinstance(a, XLError) or isinstance(b, XLError):
            return a if isinstance(a, XLError) else b
        if b == 0:
            return 0.0
        if (a > 0) != (b > 0) and a != 0:
            return NUM
        return math.floor(a / b + 0.5) * b

    @staticmethod
    def QUOTIENT(a, b):
        x, y = num(a), num(b)
        if isinstance(x, XLError) or isinstance(y, XLError):
            return x if isinstance(x, XLError) else y
        return DIV0 if y == 0 else float(int(x / y))

    @staticmethod
    def EXACT(a, b):
        return text(a) == text(b)

    @staticmethod
    def ISEVEN(v):
        x = num(v)
        return x if isinstance(x, XLError) else int(x) % 2 == 0

    @staticmethod
    def ISODD(v):
        x = num(v)
        return x if isinstance(x, XLError) else int(x) % 2 == 1


def _irr(vs, t, guess):
    """Rate with sum(v / (1 + r)^t) = 0: Newton from the guess (as Excel does), then bisection if that fails."""
    def f(r):
        return math.fsum(v / (1 + r) ** ti for v, ti in zip(vs, t))
    r = guess
    try:
        for _ in range(100):
            if r <= -1:
                break
            fr = f(r)
            df = math.fsum(-ti * v / (1 + r) ** (ti + 1) for v, ti in zip(vs, t))
            if df == 0:
                break
            step = fr / df
            r -= step
            if abs(step) < 1e-10 and r > -1:
                return r
    except (OverflowError, ZeroDivisionError):
        pass
    lo, hi = -0.9999, 10.0
    try:
        flo, fhi = f(lo), f(hi)
        if flo * fhi > 0:
            return NUM
        for _ in range(300):
            mid = (lo + hi) / 2
            fm = f(mid)
            if abs(fm) < 1e-12 or hi - lo < 1e-14:
                return mid
            if (fm > 0) == (flo > 0):
                lo, flo = mid, fm
            else:
                hi = mid
        return (lo + hi) / 2
    except (OverflowError, ZeroDivisionError):
        return NUM


def _pick(t, yes, no):
    if isinstance(t, XLError):
        return t
    if t:
        return _unmiss(yes()) if yes is not None else True
    return _unmiss(no()) if no is not None else False


def _unmiss(v):
    return 0.0 if isinstance(v, Missing) else v


def _logical(args, combine):
    vals = []
    for a in args:
        if is_array(a):
            for v in flat(a):
                if isinstance(v, XLError):
                    return v
                if isinstance(v, (bool, float)):
                    vals.append(bool(v))
        else:
            t = truth(a)
            if isinstance(t, XLError):
                return t
            vals.append(t)
    if not vals:
        return VALUE
    return bool(combine(vals))


def _info(v, test):
    if isinstance(v, list) or (isinstance(v, Rng) and v.shape != (1, 1)):
        return elementwise(test, v)  # e.g. SUMPRODUCT(--ISBLANK(range))
    return test(first(v) if isinstance(v, Rng) else v)


def _agg(args, combine, empty):
    xs = []
    for a in args:
        if is_array(a):
            for v in flat(a):
                if isinstance(v, XLError):
                    return v
                if isinstance(v, float) and not isinstance(v, bool):
                    xs.append(v)
        elif isinstance(a, Missing):
            continue
        else:
            x = num(a)
            if isinstance(x, XLError):
                return x
            xs.append(x)
    return combine(xs) if xs else empty


def _crit(c):
    """Criterion (as in SUMIFS) -> test(value)."""
    c = first(c) if is_array(c) else c
    if isinstance(c, bool):
        return lambda v: v is c
    if isinstance(c, float):
        return lambda v: isinstance(v, float) and not isinstance(v, bool) and v == c
    if c is None:
        return lambda v: v is None or v == ""
    if isinstance(c, XLError):
        return lambda v: v == c
    s = str(c)
    m = re.match(r"^(<=|>=|<>|=|<|>)?(.*)$", s, re.S)
    op, rest = m.group(1) or "=", m.group(2)
    target = num(rest) if rest.strip() != "" else None
    if isinstance(target, XLError):
        target = None
    if target is None:
        if rest == "":
            return (lambda v: v is None or v == "") if op == "=" else (lambda v: v is not None and v != "") \
                if op == "<>" else (lambda v: False)
        low = rest.lower()
        if op in ("=", "<>") and re.search(r"[*?~]", rest):
            pat = _wild(rest)
            hit = lambda v: isinstance(v, str) and pat.fullmatch(v) is not None
            return hit if op == "=" else (lambda v: not hit(v))
        tests = {"=": lambda v: isinstance(v, str) and v.lower() == low, "<>": lambda v: not (isinstance(v, str) and v.lower() == low),
                 "<": lambda v: isinstance(v, str) and v.lower() < low, ">": lambda v: isinstance(v, str) and v.lower() > low,
                 "<=": lambda v: isinstance(v, str) and v.lower() <= low, ">=": lambda v: isinstance(v, str) and v.lower() >= low}
        return tests[op]
    t = target
    isnum = lambda v: isinstance(v, float) and not isinstance(v, bool)
    return {"=": lambda v: isnum(v) and v == t, "<>": lambda v: not (isnum(v) and v == t),
            "<": lambda v: isnum(v) and v < t, ">": lambda v: isnum(v) and v > t,
            "<=": lambda v: isnum(v) and v <= t, ">=": lambda v: isnum(v) and v >= t}[op]


def _mask(pairs):
    shape = None
    mask = None
    for i in range(0, len(pairs), 2):
        g = grid(pairs[i])
        if shape is None:
            shape = (len(g), len(g[0]))
            mask = [[True] * shape[1] for _ in range(shape[0])]
        elif (len(g), len(g[0])) != shape:
            return VALUE
        test = _crit(pairs[i + 1])
        for r in range(shape[0]):
            row, mrow = g[r], mask[r]
            for c in range(shape[1]):
                if mrow[c] and not test(row[c]):
                    mrow[c] = False
    return mask


def _ifs(target, pairs, combine, empty):
    g = grid(target)
    mask = _mask(pairs)
    if isinstance(mask, XLError):
        return mask
    if (len(mask), len(mask[0])) != (len(g), len(g[0])):
        return VALUE
    xs = []
    for r in range(len(g)):
        for c in range(len(g[0])):
            if mask[r][c]:
                v = g[r][c]
                if isinstance(v, XLError):
                    return v
                if isinstance(v, float) and not isinstance(v, bool):
                    xs.append(v)
    return combine(xs) if xs else empty


def _wild(p):
    out, i = "", 0
    while i < len(p):
        ch = p[i]
        if ch == "~" and i + 1 < len(p):
            out += re.escape(p[i + 1])
            i += 2
            continue
        out += ".*" if ch == "*" else "." if ch == "?" else re.escape(ch)
        i += 1
    return re.compile(out, re.I | re.S)


def _index(arr, r, c):
    rr = 0 if isinstance(r, Missing) else num(first(r))
    cc = 0 if isinstance(c, Missing) else num(first(c))
    if isinstance(rr, XLError):
        return rr
    if isinstance(cc, XLError):
        return cc
    rr, cc = int(rr), int(cc)
    if isinstance(arr, Rng):
        nr, nc = arr.shape
        if isinstance(c, Missing) and nr == 1 and nc > 1:  # INDEX(row_range, n) picks along the row
            rr, cc = 1, rr
        if rr < 0 or cc < 0 or rr > nr or cc > nc:
            return REF
        r1, r2 = (arr.r1, arr.r2) if rr == 0 else (arr.r1 + rr - 1,) * 2
        c1, c2 = (arr.c1, arr.c2) if cc == 0 else (arr.c1 + cc - 1,) * 2
        return Rng(arr.book, arr.src, arr.sheet, r1, c1, r2, c2)
    g = grid(arr)
    nr, nc = len(g), len(g[0])
    if isinstance(c, Missing) and nr == 1:
        rr, cc = 1, rr
    if isinstance(c, Missing) and nc == 1 and nr > 1:
        cc = 1
    if rr < 0 or cc < 0 or rr > nr or cc > nc:
        return REF
    if rr == 0:
        return [[row[cc - 1]] for row in g]
    if cc == 0:
        return [list(g[rr - 1])]
    return _unmiss(g[rr - 1][cc - 1])


def _approx(vals, value, descending=False):
    """MATCH type 1 / -1 (binary search, as Excel does): position of the last value <= lookup (>= if descending)."""
    lo, hi, found = 0, len(vals) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        v = vals[mid]
        if v is None or isinstance(v, XLError) or _rank(v)[0] != _rank(value)[0]:
            # Excel skips values of another type; step to the nearest comparable one
            j = mid - 1
            while j >= lo and (vals[j] is None or isinstance(vals[j], XLError) or _rank(vals[j])[0] != _rank(value)[0]):
                j -= 1
            if j < lo:
                lo = mid + 1
                continue
            mid, v = j, vals[j]
        c = _cmp(v, value)
        if (c <= 0 and not descending) or (c >= 0 and descending):
            found = mid
            lo = mid + 1
        else:
            hi = mid - 1
    return float(found + 1) if found is not None else NA


def _lookup_pos(value, keys, approx):
    ap = truth(first(approx)) if not isinstance(approx, Missing) else True
    if ap is True:
        i = _approx(keys, value)
        return i if isinstance(i, XLError) else int(i) - 1
    for i, k in enumerate(keys):
        if k is not None and _rank(k)[0] == _rank(value)[0] and _cmp(k, value) == 0:
            return i
    return NA


def _math1(x, fn):
    if is_array(x):
        return elementwise(lambda v: _math1(v, fn), x)
    v = num(x)
    if isinstance(v, XLError):
        return v
    try:
        return fn(v)
    except (ValueError, OverflowError):
        return NUM


def _math2(x, y, fn):
    if is_array(x) or is_array(y):
        return broadcast(lambda a, b: _math2(a, b, fn), x, y)
    a, b = num(x), num(0.0 if isinstance(y, Missing) else y)
    for v in (a, b):
        if isinstance(v, XLError):
            return v
    try:
        return fn(a, b)
    except (ValueError, OverflowError, ZeroDivisionError):
        return NUM


def _round(x, d, mode):
    def one(v, k):
        k = int(k)
        q = Decimal(1).scaleb(-k)
        return float(Decimal(repr(v)).quantize(q, rounding=mode)) if k >= 0 else \
            float((Decimal(repr(v)) / Decimal(10) ** -k).quantize(Decimal(1), rounding=mode) * Decimal(10) ** -k)
    return _math2(x, 0.0 if isinstance(d, Missing) else d, one)


def _date1(s, fn):
    if is_array(s):
        return elementwise(lambda v: _date1(v, fn), s)
    v = num(s)
    if isinstance(v, XLError):
        return v
    if v < 0:
        return NUM
    return fn(to_date(v))


def _date2(s, m, fn):
    if is_array(s) or is_array(m):
        return broadcast(lambda a, b: _date2(a, b, fn), s, m)
    a, b = num(s), num(m)
    for v in (a, b):
        if isinstance(v, XLError):
            return v
    return fn(to_date(a), b)


def _add_months(d: date, months: int, end: bool = False) -> date:
    y, m = divmod(d.month - 1 + months, 12)
    y, m = d.year + y, m + 1
    last = (date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1)).day
    return date(y, m, last if end else min(d.day, last))


def _leap(y):
    return y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)


def _yearfrac(a: date, b: date, basis: int) -> float:
    if a > b:
        a, b = b, a
    if basis == 0 or basis == 4:  # 30/360 US (NASD) / European
        d1, d2 = a.day, b.day
        if basis == 0:
            last_feb = lambda d: d.month == 2 and (d + timedelta(days=1)).month == 3
            if last_feb(a) and last_feb(b):
                d2 = 30
            if last_feb(a):
                d1 = 30
            if d2 == 31 and d1 >= 30:
                d2 = 30
            if d1 == 31:
                d1 = 30
        else:
            d1, d2 = min(d1, 30), min(d2, 30)
        return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1)) / 360.0
    days = (b - a).days
    if basis == 2:
        return days / 360.0
    if basis == 3:
        return days / 365.0
    # basis 1, actual/actual as Excel computes it
    if a.year == b.year or (b.year == a.year + 1 and (a.month, a.day) >= (b.month, b.day)):
        if a.year == b.year:
            den = 366.0 if _leap(a.year) else 365.0
        else:
            leap = (_leap(a.year) and (a.month, a.day) <= (2, 29)) or (_leap(b.year) and (b.month, b.day) >= (2, 29))
            den = 366.0 if leap else 365.0
        return days / den
    years = range(a.year, b.year + 1)
    den = sum(366 if _leap(y) else 365 for y in years) / len(years)
    return days / den


def _text1(s, fn):
    if is_array(s) and not isinstance(s, Rng):
        return elementwise(lambda v: _text1(v, fn), s)
    v = first(s)
    return v if isinstance(v, XLError) else fn(text(v))


def _text2(s, n, fn):
    v, k = first(s), num(first(n)) if not isinstance(n, Missing) else 1.0
    if isinstance(v, XLError):
        return v
    if isinstance(k, XLError) or k < 0:
        return VALUE
    return fn(text(v), k)


def _find(needle, hay, start, ci):
    n, h, s = first(needle), first(hay), num(first(start)) if not isinstance(start, Missing) else 1.0
    for v in (n, h):
        if isinstance(v, XLError):
            return v
    n, h = text(n), text(h)
    if ci:
        pat = _wild(n)
        m = pat.search(h, int(s) - 1)
        return float(m.start() + 1) if m else VALUE
    i = h.find(n, int(s) - 1)
    return float(i + 1) if i >= 0 else VALUE



class _Err(Exception):
    def __init__(self, err):
        self.err = err


xl = _XL()


class UnknownFunction:
    def __init__(self, book, name):
        self.book, self.name = book, name

    def __call__(self, *args):
        self.book.unsupported[self.name] += 1
        if self.book.rec is not None:
            self.book.rec_unknown.add(self.name)
        return NAME


# ---- the book -----------------------------------------------------------------------------------------------

class Book:
    """Where a compiled overlay's row functions get their inputs and keep their results."""

    def __init__(self):
        self.rows = {}           # (sheet, row) -> row function (memoised)
        self.formula_cols = {}   # (sheet, row) -> [(c1, c2)] columns the row function computes
        self.labels = {}         # (sheet, row) -> label, for tracing
        self.overlay = set()     # sheets the overlay covers
        self.inputs = {}         # (sheet, row, col) -> constant on the overlay sheets
        self.overrides = {}      # (sheet, row, col) -> scenario value
        self.memo = {}
        self.range_cache = {}
        self.busy = set()
        self.cycles = []
        self.unsupported = Counter()
        self.volatile = set()
        self.feed = lambda sheet, row, col: None           # this workbook, outside the overlay
        self.ext = lambda idx, sheet, row, col: None       # another workbook, through [idx]
        self.cached = lambda sheet, row, col: None         # the workbook's saved values
        self.feed_log = None                               # set -> records each feed read (for tracing)
        self.rec = None          # a set while one cell's formula runs again for the doctor: every cell it reads
        self.rec_unknown = set()  # ... and the unsupported functions it called

    # registration, used by the compiled module
    def row(self, sheet, row, cols, label=""):
        def deco(fn):
            key = (sheet, row)
            memo, busy, overrides = self.memo, self.busy, self.overrides

            def cell(c):
                k = (sheet, row, c)
                rec = self.rec
                if rec is not None:
                    rec.add(("", sheet, row, c))
                if k in overrides:
                    return overrides[k]
                v = memo.get(k, _MISS)
                if v is not _MISS:
                    return v
                if not any(a <= c <= b for a, b in cols):  # a constant (or blank) in this row, not a formula
                    return self.inputs.get(k)
                if k in busy:  # circular reference: Excel iterates; we take the saved value and say so
                    self.cycles.append(k)
                    return self.cached(sheet, row, c)
                busy.add(k)
                if rec is not None:  # record only the cells the traced formula reads itself, not their inputs
                    self.rec = None
                try:
                    v = scalar(fn(c), row, c)
                finally:
                    busy.discard(k)
                    if rec is not None:
                        self.rec = rec
                memo[k] = v
                return v
            cell.__name__, cell.__doc__ = fn.__name__, fn.__doc__
            cell.raw = fn
            self.rows[key] = cell
            self.formula_cols[key] = cols
            self.labels[key] = label
            return cell
        return deco

    def unknown(self, name):
        return UnknownFunction(self, name)

    # values
    def input(self, sheet, row, col):
        k = (sheet, row, col)
        if self.rec is not None:
            self.rec.add(("", sheet, row, col))
        if k in self.overrides:
            return self.overrides[k]
        return self.inputs.get(k)

    def get(self, src, sheet, row, col):
        if src == "":
            if sheet in self.overlay:
                fn = self.rows.get((sheet, row))
                return fn(col) if fn else self.input(sheet, row, col)
            return self.feed_value(sheet, row, col)
        return self.ext_value(src, sheet, row, col)

    def feed_value(self, sheet, row, col):
        k = (sheet, row, col)
        if self.rec is not None:
            self.rec.add(("", sheet, row, col))
        if k in self.overrides:
            return self.overrides[k]
        v = self.feed(sheet, row, col)
        if self.feed_log is not None:
            self.feed_log[("", sheet, row, col)] = v
        return v

    def ext_value(self, idx, sheet, row, col):
        k = (f"[{idx}]{sheet}", row, col)
        if self.rec is not None:
            self.rec.add((idx, sheet, row, col))
        if k in self.overrides:
            return self.overrides[k]
        v = self.ext(idx, sheet, row, col)
        if self.feed_log is not None:
            self.feed_log[(idx, sheet, row, col)] = v
        return v

    def rng(self, src, sheet, r1, c1, r2, c2):
        return Rng(self, src, sheet, r1, c1, r2, c2)

    def is_formula(self, sheet, row, col):
        return any(a <= col <= b for a, b in self.formula_cols.get((sheet, row), ()))

    def reads(self, sheet, row, col) -> tuple[set, set]:
        """The cells one formula cell reads, found by running its formula again (what it reads is already worked
        out, so this is quick), and the unsupported functions it calls: ({(src, sheet, row, col)}, {name}). src is
        "" for this workbook, else the external link's number."""
        fn = self.rows.get((sheet, row))
        if not fn or not self.is_formula(sheet, row, col):
            return set(), set()
        self.rec, self.rec_unknown = set(), set()
        try:
            fn.raw(col)
        except Exception:
            pass
        finally:
            rec, unknown, self.rec = self.rec, self.rec_unknown, None
        rec.discard(("", sheet, row, col))
        return rec, unknown

    def reset(self):
        self.memo.clear()
        self.range_cache.clear()
        self.cycles.clear()
        self.unsupported.clear()


def scalar(v, row, col):
    """A cell holds one value: a reference or array result is reduced by implicit intersection (the value in the
    cell's own row or column), as Excel does for a formula that isn't array-entered. A formula never returns
    blank: =A1 with A1 empty is 0."""
    if isinstance(v, Rng):
        nr, nc = v.shape
        if nr == 1 and nc == 1:
            v = v.cell(0, 0)
        elif nr == 1 and v.c1 <= col <= v.c2:
            v = v.cell(0, col - v.c1)
        elif nc == 1 and v.r1 <= row <= v.r2:
            v = v.cell(row - v.r1, 0)
        else:
            return VALUE
    elif isinstance(v, list):
        v = v[0][0] if v and v[0] else None
    return 0.0 if v is None or isinstance(v, Missing) else v


def same(a, b, rel=1e-7) -> bool:
    """Python result vs the workbook's saved value."""
    if isinstance(a, bool) or isinstance(b, bool):
        return (a is True or a == 1.0) == (b is True or b == 1.0) if isinstance(a, (bool, float)) and isinstance(b, (bool, float)) else a == b
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= rel * max(1.0, abs(a), abs(b))
    if (a is None or a == "") and (b is None or b == ""):
        return True
    if (a is None and b == 0.0) or (b is None and a == 0.0):
        return True
    return a == b
