import json
from pathlib import Path

import pytest

from sheetre import excelfn as X
from decimal import Decimal

FIX = json.loads((Path(__file__).parent / "fixtures" / "approx_match_excel.json").read_text())


def test_approx_match_reproduces_excel_on_unsorted_data():
    misses = []
    for c in FIX["cases"]:
        try:
            got = X.match(c["x"], c["keys"], 1)
        except X.XLError:
            got = 0
        if got != c["excel"]:
            misses.append(c)
    assert not misses, f"{len(misses)}/{len(FIX['cases'])} differ from Excel, e.g. {misses[:3]}"


def test_vlookup_approx_uses_same_search():
    table = [[0, "a"], [25, "b"], [50, "c"], [150, "d"], [100, "e"], [250, "f"]]
    assert X.vlookup(160, table, 2, True) == X.index([r[1] for r in table], X.match(160, [r[0] for r in table], 1))
    with pytest.raises(X.XLError):
        X.vlookup(-1, table, 2, True)


def test_exact_lookup_first_match_and_case_insensitive():
    table = [["ab", 1], ["AB", 2], ["c", 3]]
    assert X.vlookup("Ab", table, 2, False) == 1
    with pytest.raises(X.XLError):
        X.vlookup("ab ", table, 2, False)


def test_round_half_away_from_zero():
    assert X.xl_round(2.675, 2) == 2.68
    assert X.xl_round(-0.5) == -1.0
    assert X.xl_round(0.125, 2) == 0.13
    assert X.roundup(1.231, 2) == 1.24 and X.rounddown(-1.239, 2) == -1.23


def test_round_sees_15_significant_digits():
    """ROUND must round the 15-significant-digit value, not repr().

    All three of these sit just *below* a .xx5 midpoint in binary but land exactly *on*
    it at 15 significant digits, which is why Excel rounds them away from zero. Going
    through repr() got the first one right by luck and the other two inconsistently,
    because repr() returns the shortest round-tripping string: '2.675' and '282705.665'
    look like exact midpoints while '112187.94499999999' does not.
    Found by the Team_Budget_Control rebuild cross-check.
    """
    assert X._approx15(2.675) == Decimal("2.67500000000000")
    assert X.xl_round(2.675, 2) == 2.68

    q1 = 2019383.0099999998 / 18          # 112187.944999999992433...
    assert X._approx15(q1) == Decimal("112187.945000000")
    assert X.xl_round(q1, 2) == 112187.95

    q2 = 9611992.61 / 34                  # 282705.664999999979045...
    assert X._approx15(q2) == Decimal("282705.665000000")
    assert X.xl_round(q2, 2) == 282705.67


def test_round_below_midpoint_at_15_digits_rounds_down():
    """A value that is still below the midpoint after the 15-digit step rounds down."""
    assert X.xl_round(1.0049999, 2) == 1.0
    assert X.xl_round(-1.0049999, 2) == -1.0
    assert X.xl_round(0.0, 2) == 0.0
    assert X.xl_round(1234.0, -2) == 1200.0


def test_text_coercion():
    assert X.value("12,345") == 12345
    assert X.value("1,23,456") == X.VALUE
    assert X.value("15%") == 0.15
    assert X.num(None) == 0 and X.num(True) == 1
    assert X.n("123") == 0 and X.n(5) == 5
    with pytest.raises(X.XLError):
        X.num("abc")


def test_datevalue_locale():
    assert X.from_serial(X.datevalue("05/06/2028", "en_IN")).month == 6
    assert X.from_serial(X.datevalue("05/06/2028", "en_US")).month == 5
    assert X.from_serial(X.datevalue("2028-06-05")).day == 5
    assert X.datevalue("31/02/2028") == X.VALUE


def test_sumif_criteria():
    keys = ["North", "north ", "South", None, 5, 12]
    vals = [1, 2, 4, 8, 16, 32]
    assert X.sumif(keys, "north", vals) == 1
    assert X.sumif(keys, "No*", vals) == 3
    assert X.sumif(keys, ">10", vals) == 32
    assert X.sumif(keys, "<>South", vals) == 1 + 2 + 8 + 16 + 32
    assert X.countif(keys, "") == 1
    assert X.sumifs(vals, keys, "No*", vals, ">1") == 2


def test_sum_ignores_text_but_plus_coerces():
    assert X.xl_sum([1, "2", None, 3]) == 4
    assert X.num("2") + 1 == 3


def test_iferror_catches_raised_and_returned_errors():
    assert X.iferror(lambda: X.vlookup("zz", [["a", 1]], 2, False), 0) == 0
    assert X.iferror(lambda: 1 / 0, -1) == -1
    assert X.iferror(X.NA, "x") == "x"


def test_iterate_converges_like_interest_on_average_balance():
    def step(s):
        interest = (s["open"] + s["close"]) / 2 * 0.06 / 12
        return dict(s, interest=interest, close=s["open"] + 100 + interest)
    s = X.iterate(step, dict(open=1000.0, close=0.0, interest=0.0), max_iter=100, max_change=1e-9)
    assert abs(s["close"] - (1000 + 100 + (1000 + s["close"]) / 2 * 0.005)) < 1e-6


def test_cells_helper():
    c = X.Cells({"S!A2": 1, "S!A3": 2, "S!B2": "x"})
    assert c.col("S!A2:A4") == [1, 2, None]
    assert c.table("S!A2:B3") == [[1, "x"], [2, None]]
    assert c["S!B2"] == "x"
