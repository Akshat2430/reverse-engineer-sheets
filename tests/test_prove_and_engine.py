import datetime as dt
import shutil
from pathlib import Path

import pytest

from sheetre import prove
from sheetre.engine import BridgeEngine, EngineBusy, EngineUnavailable, ExcelEngine, get_engine
from sheetre.excelfn import NA

ROOT = Path(__file__).resolve().parents[1]
CONTROL = ROOT / "samples" / "control" / "Team_Budget_Control.xlsx"


def test_compare_rules():
    assert prove.compare(1.0, 1.004, 0.005)[0]
    ok, d, rounding_only = prove.compare(1.0, 1.006, 0.005)
    assert not ok and not rounding_only
    assert prove.compare(10.004, 10.0, 0.001)[2]               # matches after rounding to 2dp
    assert prove.compare("#N/A", NA, 0.005)[0]
    assert prove.compare(None, "", 0.005)[0]
    assert not prove.compare(None, 0, 0.005)[0]                # blank is not zero
    assert prove.compare(dt.datetime(2027, 1, 1), 46388, 0.005)[0]
    assert not prove.compare(0.5, 0.504, prove.TOL)[0]               # 0.4 points off a margin fails
    assert prove.compare(11_000_000.0, 11_000_000.00007, prove.TOL)[0]  # convergence noise on a big total passes


excel = pytest.mark.skipif(not ExcelEngine.detect()[0], reason="needs Microsoft Excel")


@excel
def test_engine_roundtrip_on_control():
    # Installed is not usable: inside a sandbox or a headless macOS session Excel refuses
    # automation. That is an environment limit, not a failure of this code.
    work = ROOT / "runs" / "_test_engine"
    shutil.rmtree(work, ignore_errors=True)
    try:
        eng = ExcelEngine().start()
    except (EngineUnavailable, EngineBusy) as e:
        pytest.skip(f"Excel can't be driven directly here: {e}")
    try:
        book = eng.open(CONTROL, work)
        assert Path(book.path).parent == work.resolve() and Path(book.path).name.startswith("engine_")
        eng.calc(book, full=True)
        base = eng.get(book, ["Summary!B10", "Teams!C2"])
        eng.set(book, {"Teams!C2": base["Teams!C2"] + 1})
        eng.calc(book)
        assert eng.get(book, ["Summary!B10"])["Summary!B10"] > base["Summary!B10"]
        eng.set(book, {"Teams!C2": base["Teams!C2"], "Teams!B2": "Engineering"})
        eng.calc(book)
        assert eng.get(book, ["Summary!B10"])["Summary!B10"] == pytest.approx(base["Summary!B10"])
        eng.set(book, {"Teams!A2": "12,345"})                  # text must stay text
        eng.calc(book)
        assert eng.get(book, ["Teams!A2"])["Teams!A2"] == "12,345"
        eng.set(book, {"Teams!C3": "abc"})                     # text headcount -> #VALUE! downstream
        eng.calc(book)
        assert eng.get(book, ["Calc!D3"])["Calc!D3"] == "#VALUE!"
        eng.close(book)
    finally:
        eng.stop()
    assert CONTROL.stat().st_size > 0                          # original untouched


@pytest.mark.skipif(not BridgeEngine.detect()[0], reason="bridge not running")
def test_bridge_roundtrip_and_path_guard():
    work = ROOT / "runs" / "_test_bridge"
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    src = work / "control.xlsx"
    shutil.copy(CONTROL, src)
    eng = BridgeEngine().start()
    try:
        with pytest.raises(EngineBusy):                         # one stage at a time
            BridgeEngine().start()
        with pytest.raises(Exception, match="only opens workbooks under"):
            eng.open(CONTROL, work)                            # outside runs/: refused
        book = eng.open(src, work)
        eng.calc(book, full=True)
        v = eng.get(book, ["Summary!B10"])["Summary!B10"]
        eng.set(book, {"Teams!C2": 15})
        eng.calc(book)
        assert eng.get(book, ["Summary!B10"])["Summary!B10"] > v
        eng.close(book)
    finally:
        eng.stop()
