"""Excel-faithful helpers for rebuilds (rebuild/model.py imports these).

Behaviour here was checked against Excel 16 for Mac, not guessed:
- approximate match (VLOOKUP TRUE / MATCH 1) on UNSORTED data follows Excel's binary search
  exactly: 700/700 random cases (tests/test_excelfn.py, runs/_research).
- ROUND is half away from zero on the decimal value, not banker's rounding.
- Text in arithmetic coerces only when Excel would ("12,345" yes, Indian grouping "1,23,456" no, on en-IN).
"""
import datetime as dt
import math
import re
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal

EPOCH = dt.date(1899, 12, 30)


class XLError(Exception):
    """An Excel error value. Compare with ==, e.g. result == NA."""

    def __init__(self, code):
        super().__init__(code)
        self.code = code

    def __eq__(self, other):
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)

    def __repr__(self):
        return self.code


NA, VALUE, REF, DIV0, NUM, NAME = (XLError(c) for c in ("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?"))


def is_error(v):
    return isinstance(v, XLError)


def iferror(fn, fallback):
    """IFERROR(expr, fallback). Pass expr as a zero-arg callable so errors raised inside are caught."""
    try:
        v = fn() if callable(fn) else fn
    except XLError as e:
        return fallback
    except ZeroDivisionError:
        return fallback
    return fallback if is_error(v) else v


# ---------------------------------------------------------------- coercion

def num(v, locale="en_IN"):
    """Value as used in arithmetic (=A1*2). Blank -> 0, TRUE -> 1, numeric text coerced, else #VALUE!."""
    if v is None:
        return 0
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, (dt.datetime, dt.date)):
        return to_serial(v)
    if is_error(v):
        raise v
    if isinstance(v, str):
        r = value(v, locale)
        if is_error(r):
            raise r
        return r
    raise VALUE


def n(v):
    """N(): numbers pass, dates become serials, TRUE->1, text and blanks -> 0."""
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, (dt.datetime, dt.date)):
        return to_serial(v)
    if is_error(v):
        raise v
    return 0


_US_GROUP = re.compile(r"^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$")


def value(s, locale="en_IN"):
    """VALUE(text). Accepts plain numbers, 3-digit grouping, %, and dates/times Excel would parse."""
    t = s.strip()
    if t == "":
        return VALUE
    pct = t.endswith("%")
    core = t[:-1].strip() if pct else t
    try:
        x = float(core)
        return x / 100 if pct else x
    except ValueError:
        pass
    if _US_GROUP.match(core):
        x = float(core.replace(",", ""))
        return x / 100 if pct else x
    d = datevalue(t, locale)
    if not is_error(d):
        return d
    return VALUE


def datevalue(s, locale="en_IN"):
    """DATEVALUE(text): day-first for en_IN/en_GB, month-first for en_US. ISO always works."""
    t = s.strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", t)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$", t)
        if not m:
            return VALUE
        a, b, y = map(int, m.groups())
        if y < 100:
            y += 2000 if y < 30 else 1900
        d, mo = (a, b) if locale.lower() in ("en_in", "en_gb", "en_au", "en_nz") else (b, a)
    try:
        return to_serial(dt.date(y, mo, d))
    except ValueError:
        return VALUE


# ---------------------------------------------------------------- dates

def to_serial(d):
    if isinstance(d, dt.datetime):
        delta = d - dt.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    return (d - EPOCH).days


def from_serial(x):
    return EPOCH + dt.timedelta(days=int(x))


def month(v):
    return from_serial(num(v)).month


def year(v):
    return from_serial(num(v)).year


def day(v):
    return from_serial(num(v)).day


def today_serial(today=None):
    return to_serial(today or dt.date.today())


# ---------------------------------------------------------------- rounding

def _approx15(x):
    """The value Excel's ROUND actually sees.

    Excel carries 15 significant decimal digits, so ROUND(2.675, 2) is 2.68 even though
    2.675 is stored as 2.67499999999999982...: to 15 digits that value *is* 2.675, which
    then rounds half away from zero. Going through repr() instead is not equivalent --
    repr() returns the shortest round-tripping string, so whether a value lands exactly
    on a .xx5 midpoint becomes an accident of float formatting and two values that are
    both below their midpoint can round in opposite directions.
    """
    if x == 0 or not math.isfinite(x):
        return Decimal(0)
    d = Decimal(x)  # exact binary value, no formatting in the way
    return d.quantize(Decimal(1).scaleb(d.adjusted() - 14), rounding=ROUND_HALF_EVEN)


def xl_round(x, digits=0):
    q = Decimal(1).scaleb(-digits)
    return float(_approx15(float(x)).quantize(q, rounding=ROUND_HALF_UP))


def roundup(x, digits=0):
    f = 10 ** digits
    return math.copysign(math.ceil(abs(x) * f - 1e-12) / f, x)


def rounddown(x, digits=0):
    f = 10 ** digits
    return math.copysign(math.floor(abs(x) * f + 1e-12) / f, x)


# ---------------------------------------------------------------- comparison

def _rank(v):
    """Excel ordering: numbers < text < logicals; blank behaves like 0 / ''."""
    if isinstance(v, bool):
        return (2, v)
    if isinstance(v, (int, float)):
        return (0, v)
    if isinstance(v, str):
        return (1, v.lower())
    if v is None:
        return (0, 0)
    if isinstance(v, (dt.date, dt.datetime)):
        return (0, to_serial(v))
    raise VALUE


def eq(a, b):
    """Excel '=': case-insensitive for text, no text/number coercion."""
    if a is None and b is None:
        return True
    if a is None:
        a = "" if isinstance(b, str) else 0
    if b is None:
        b = "" if isinstance(a, str) else 0
    return _rank(a) == _rank(b)


def lt(a, b):
    return _rank(a) < _rank(b)


def le(a, b):
    return _rank(a) <= _rank(b)


# ---------------------------------------------------------------- lookup

def _approx_pos(keys, x):
    """Excel's binary search for approximate match (match_type 1 / range_lookup TRUE).
    On unsorted keys this returns what Excel returns, not the 'right' bracket."""
    lo, hi, ans = 0, len(keys) - 1, -1
    rx = _rank(x)
    while lo <= hi:
        mid = (lo + hi) // 2
        rk = _rank(keys[mid])
        if rk == rx:
            while mid + 1 < len(keys) and _rank(keys[mid + 1]) == rx:
                mid += 1
            return mid
        if rk < rx:
            ans, lo = mid, mid + 1
        else:
            hi = mid - 1
    return ans


def _wild(pattern):
    rx = re.escape(pattern).replace(r"\*", ".*").replace(r"\?", ".").replace("~\\*", r"\*").replace("~\\?", r"\?")
    return re.compile("^" + rx + "$", re.I | re.S)


def _exact_pos(keys, x):
    if isinstance(x, str) and any(ch in x for ch in "*?"):
        rx = _wild(x)
        for i, k in enumerate(keys):
            if isinstance(k, str) and rx.match(k):
                return i
        return -1
    for i, k in enumerate(keys):
        if k is not None and eq(k, x):
            return i
    return -1


def match(x, keys, match_type=1):
    if is_error(x):
        raise x
    if match_type == 0:
        i = _exact_pos(keys, x)
    elif match_type == 1:
        i = _approx_pos(keys, x)
    else:                                       # -1: descending data, smallest value >= x (not yet Excel-verified)
        i = max((j for j, k in enumerate(keys) if k is not None and not lt(k, x)), default=-1)
    if i < 0:
        raise NA
    return i + 1


def vlookup(x, table, col, approx=True):
    """table: list of rows (lists). col: 1-based."""
    if is_error(x):
        raise x
    keys = [row[0] for row in table]
    i = _approx_pos(keys, x) if approx else _exact_pos(keys, x)
    if i < 0:
        raise NA
    if col < 1 or col > len(table[i]):
        raise REF
    return table[i][col - 1]


def hlookup(x, table, row, approx=True):
    return vlookup(x, [list(c) for c in zip(*table)], row, approx)


def index(arr, row, col=None):
    if col is None:
        if arr and isinstance(arr[0], (list, tuple)):
            if len(arr[0]) == 1:
                return arr[row - 1][0]
            return arr[0][row - 1] if len(arr) == 1 else arr[row - 1]
        return arr[row - 1]
    return arr[row - 1][col - 1]


# ---------------------------------------------------------------- conditional aggregates

_CRIT = re.compile(r"^(<=|>=|<>|<|>|=)?(.*)$", re.S)


def criteria(c, locale="en_IN"):
    """Build a predicate from a SUMIF/COUNTIF criterion (number, text, '>=5', 'a*', '<>x')."""
    if not isinstance(c, str):
        return lambda v: v is not None and not isinstance(v, str) and not isinstance(v, bool) and eq(v, c)
    op, rest = _CRIT.match(c).groups()
    op = op or "="
    target = rest
    numeric = value(rest, locale) if rest.strip() else None
    if numeric is not None and not is_error(numeric):
        t = numeric
        ops = {"=": lambda v: _numeric(v) and v == t,
               "<>": lambda v: not (_numeric(v) and v == t),
               "<": lambda v: _numeric(v) and v < t,
               ">": lambda v: _numeric(v) and v > t,
               "<=": lambda v: _numeric(v) and v <= t,
               ">=": lambda v: _numeric(v) and v >= t}
        return ops[op]
    if op in ("=", "<>"):
        if target == "":
            pred = lambda v: v is None or v == ""
        elif any(ch in target for ch in "*?"):
            rx = _wild(target)
            pred = lambda v: isinstance(v, str) and bool(rx.match(v))
        else:
            pred = lambda v: isinstance(v, str) and v.lower() == target.lower()
        return pred if op == "=" else (lambda v: not pred(v))
    t = target.lower()
    cmp = {"<": lambda a: a < t, ">": lambda a: a > t, "<=": lambda a: a <= t, ">=": lambda a: a >= t}[op]
    return lambda v: isinstance(v, str) and cmp(v.lower())


def _numeric(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def sumif(rng, crit, sum_rng=None):
    pred = criteria(crit)
    sum_rng = sum_rng if sum_rng is not None else rng
    total = 0
    for k, s in zip(rng, sum_rng):
        if pred(k):
            if is_error(s):
                raise s
            if _numeric(s):
                total += s
    return total


def sumifs(sum_rng, *pairs):
    preds = [(pairs[i], criteria(pairs[i + 1])) for i in range(0, len(pairs), 2)]
    total = 0
    for j, s in enumerate(sum_rng):
        if all(p(r[j]) for r, p in preds):
            if is_error(s):
                raise s
            if _numeric(s):
                total += s
    return total


def countif(rng, crit):
    pred = criteria(crit)
    return sum(1 for v in rng if pred(v))


def countifs(*pairs):
    preds = [(pairs[i], criteria(pairs[i + 1])) for i in range(0, len(pairs), 2)]
    n_ = len(pairs[0])
    return sum(1 for j in range(n_) if all(p(r[j]) for r, p in preds))


def xl_sum(*args):
    """SUM over ranges: text and blanks in ranges are ignored (unlike + which coerces)."""
    total = 0
    for a in args:
        items = a if isinstance(a, (list, tuple)) else [a]
        for v in items:
            if isinstance(v, (list, tuple)):
                total += xl_sum(v)
            elif is_error(v):
                raise v
            elif _numeric(v):
                total += v
    return total


def average(*args):
    vals = [v for a in args for v in (a if isinstance(a, (list, tuple)) else [a]) if _numeric(v)]
    if not vals:
        raise DIV0
    return sum(vals) / len(vals)


# ---------------------------------------------------------------- iteration

def iterate(step, state, max_iter=100, max_change=0.001):
    """Mimic Excel iterative calculation: repeat step(state)->state until every numeric value
    changes by less than max_change, or max_iter passes. Returns the final state (a dict)."""
    for _ in range(max_iter):
        new = step(dict(state))
        delta = max((abs(new[k] - state.get(k, 0)) for k in new if _numeric(new[k])), default=0)
        state = new
        if delta < max_change:
            break
    return state


# ---------------------------------------------------------------- input access for model.py

class Cells:
    """Read the inputs dict by Excel address. Missing cells read as None (blank).

        c = Cells(inputs)
        c["Rates!B2"]                 -> one value
        c.col("Inputs!E2:E61")        -> list of values down a column
        c.table("Rates!A2:B10")       -> list of rows
    """

    def __init__(self, inputs):
        self.d = inputs

    def __getitem__(self, key):
        return self.d.get(key)

    def _rect(self, rng):
        from openpyxl.utils import get_column_letter
        from .refs import parse_ref
        r = parse_ref(rng, rng.rsplit("!", 1)[0])
        return [[self.d.get(f"{r.sheet}!{get_column_letter(c)}{row}") for c in range(r.c1, r.c2 + 1)]
                for row in range(r.r1, r.r2 + 1)]

    def col(self, rng):
        return [row[0] for row in self._rect(rng)]

    def row(self, rng):
        return self._rect(rng)[0]

    def table(self, rng):
        return self._rect(rng)
