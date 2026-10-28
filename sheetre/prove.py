"""Stage 6: Prove. Run N input sets through the original (recalculated in Excel) and through
rebuild/model.py, and compare every output.

model.py contract:
    def run(inputs: dict, today: datetime.date) -> dict
        inputs: {"Sheet!A1": value} for every true input (dates as datetime.date/datetime)
        returns {"Sheet!A1": value} for every output cell (errors as excelfn.XLError or "#N/A" text)

Cases: case 0 = the workbook's own saved inputs and N random input sets (default 1,000), all
recalculated fresh in Excel; plus EVERY probe run, replayed against Excel's recorded answer from
probe_moves.jsonl (no recalculation needed). Default tolerance 0.005 on stored values.
"""
import csv
import datetime as dt
import importlib.util
import json
import random
import traceback
from collections import Counter, defaultdict
from pathlib import Path

from . import refs as R
from .engine import EngineError
from .probe import has_error
from .excelfn import XLError

# Excel keeps 15 significant digits: differences below ~1e-14 of the value are float noise, not logic
PRECISION_FLOOR = 1e-13



def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _from_json(v):
    if isinstance(v, dict) and "date" in v:
        return dt.datetime.fromisoformat(v["date"])
    return v


def _to_json(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return {"date": v.isoformat() if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day).isoformat()}
    return v


def norm(v):
    """Common form for comparing Excel and model values."""
    if isinstance(v, XLError):
        return v.code
    if v == "" or v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, dt.datetime):
        from .excelfn import to_serial
        return to_serial(v)
    if isinstance(v, dt.date):
        from .excelfn import to_serial
        return to_serial(v)
    return v


def compare(excel, model, tol):
    a, b = norm(excel), norm(model)
    if _num(a) and _num(b):
        d = abs(a - b)
        if d <= tol:
            return True, d, False
        if d <= PRECISION_FLOOR * max(abs(a), abs(b)):    # beyond Excel's 15 significant digits
            return True, d, False
        return False, d, round(a, 2) == round(b, 2)
    if a is None and b == 0 or b is None and a == 0:
        return False, None, False
    return a == b, None, False


def load_model(path):
    spec = importlib.util.spec_from_file_location("rebuild_model", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    if not hasattr(mod, "run"):
        raise RuntimeError(f"{path} has no run(inputs, today) function")
    return mod


def _decimals(x):
    s = repr(float(x))
    return 0 if float(x).is_integer() else min(len(s.split(".")[1]), 6)


def _per_cell(blk):
    """Small or mixed-scale blocks (assumption lists) hold different parameters: vary each around its own value."""
    if blk.get("cells", 0) <= 8:
        return True
    lo, hi = abs(blk.get("min", 0) or 0), abs(blk.get("max", 0) or 0)
    return lo > 0 and hi / lo > 1000


def _random_value(rng, blk, base):
    if isinstance(base, bool):
        return rng.choice([True, False])
    if _num(base) and "min" in blk and _per_cell(blk) and base != 0:
        v = base * rng.uniform(0.5, 1.5)
        return int(round(v)) if blk.get("integers") and float(base).is_integer() else round(v, max(_decimals(base), 4))
    if _num(base) and "min" in blk:
        lo, hi = blk["min"], blk["max"]
        if lo == hi:
            lo, hi = (lo * 0.5, lo * 1.5) if lo else (-1, 1)
            lo, hi = min(lo, hi), max(lo, hi)
        v = rng.uniform(lo, hi)
        if blk.get("integers"):
            return int(round(v))
        return round(v, max(_decimals(base), 2))
    if isinstance(base, (dt.date, dt.datetime)) and "min_date" in blk:
        lo, hi = dt.datetime.fromisoformat(blk["min_date"]), dt.datetime.fromisoformat(blk["max_date"])
        if lo == hi:
            lo, hi = lo - dt.timedelta(days=60), hi + dt.timedelta(days=60)
        return lo + dt.timedelta(days=rng.randint(0, (hi - lo).days))
    if isinstance(base, str) and blk.get("distinct"):
        return rng.choice(blk["distinct"])
    return base


def make_cases(graph, run_dir, n, seed):
    """Excel cases: the saved inputs plus n random input sets within each block's observed range."""
    rng = random.Random(seed)
    nodes = graph["nodes"]
    base = {k: _from_json(nodes[k]["value"]) for k in graph["inputs"]}
    cases = [dict(id=0, kind="saved inputs", source="excel", changes={})]
    by_block = []
    for blk in graph["input_blocks"]:
        col = R.parse_cell(f"{blk['col']}1")[0]
        cells = [R.cell_key(blk["sheet"], col, r) for r in range(blk["first_row"], blk["last_row"] + 1)]
        by_block.append((blk, [c for c in cells if c in base]))
    while len(cases) < n + 1:
        changes = {}
        for blk, cells in by_block:
            if rng.random() < 0.5:
                continue
            for c in cells:
                b = base[c]
                if rng.random() < 0.6 and not (isinstance(b, str) and blk.get("types", {}).get("number")):
                    changes[c] = _random_value(rng, blk, b)
        cases.append(dict(id=len(cases), kind="random", source="excel", changes=changes))
    return base, cases


def replay_cases(run_dir, start_id, outputs):
    """Every probe run, with Excel's recorded answer: checked against the rebuild without recalculating."""
    run_dir = Path(run_dir)
    summ, moves = run_dir / "probe_summary.json", run_dir / "probe_moves.jsonl"
    if not (summ.exists() and moves.exists()):
        return [], None
    ps = json.loads(summ.read_text())
    base_out = {o: _from_json(v) for o, v in ps.get("base_outputs", {}).items()}
    if not base_out or any(o not in base_out for o in outputs):
        return [], None
    cases = []
    for line in open(moves):
        r = json.loads(line)
        xl = dict(base_out)
        xl.update({o: _from_json(v) for o, v in r["moved"].items()})
        cases.append(dict(id=start_id + len(cases), kind=f"replay: {r['variant']}", source="replay",
                          changes={r["input"]: _from_json(r["value"])}, excel=xl, probe_run=r["run_id"]))
    return cases, dt.date.fromisoformat(ps["probe_date"])


def _parse_probe_value(text, base):
    if text == "":
        return None
    if isinstance(base, (dt.date, dt.datetime)):
        try:
            return dt.datetime.fromisoformat(text)
        except ValueError:
            return base
    if isinstance(base, str):
        return text
    try:
        f = float(text)
        return int(f) if f.is_integer() and isinstance(base, int) else f
    except ValueError:
        return text


def run(run_dir, engine, work_file, n=200, seed=1, tol=0.005, log=print):
    run_dir = Path(run_dir)
    graph = json.loads((run_dir / "graph.json").read_text())
    model_path = run_dir / "rebuild" / "model.py"
    if not model_path.exists():
        raise SystemExit(f"missing {model_path}: write the rebuild before proving")
    model = load_model(model_path)
    outputs = list(getattr(model, "OUTPUTS", None) or graph["outputs"])
    base, cases = make_cases(graph, run_dir, n, seed)
    replays, probe_day = replay_cases(run_dir, len(cases), outputs)
    cases += replays
    today = dt.date.today()

    book = engine.open(work_file, run_dir / "work")
    per_output = defaultdict(lambda: Counter())
    fail_rows, case_rows, failing_cases, rounding_only = [], [], [], set()
    block_fail, block_seen = Counter(), Counter()
    try:
        engine.calc(book, full=True)
        prev = {}
        iterative = bool(engine.info.get("iteration", {}).get("on"))
        resets = 0
        for case in cases:
            ch = case["changes"]
            if case["source"] == "excel":
                restore = {c: _to_json(base[c]) for c in prev if c not in ch}
                engine.set(book, {**restore, **{c: _to_json(v) for c, v in ch.items()}})
                engine.calc(book)
                xl = engine.get(book, outputs)
                prev = ch
                day = today
                if iterative and has_error(xl.values()):
                    # errors stick inside circular chains: start the next case from a fresh copy
                    engine.close(book)
                    book = engine.open(work_file, run_dir / "work")
                    engine.calc(book, full=True)
                    prev, resets = {}, resets + 1
            else:
                xl, day = case["excel"], probe_day
            inputs = {**base, **ch}
            try:
                got = model.run(dict(inputs), day)
                err = None
            except Exception as e:
                got, err = {}, f"{type(e).__name__}: {e}"
            fails = []
            for o in outputs:
                if err:
                    ok, d, ro = False, None, False
                else:
                    ok, d, ro = compare(xl.get(o), got.get(o, "<missing>"), tol)
                per_output[o]["pass" if ok else "fail"] += 1
                if not ok:
                    if ro:
                        rounding_only.add(o)
                    fails.append(o)
                    if len(fail_rows) < 5000:
                        fail_rows.append(dict(case_id=case["id"], kind=case["kind"], output=o, excel=norm(xl.get(o)),
                                              model=err or norm(got.get(o, "<missing>")), diff=d,
                                              rounding_only=ro))
            touched = {_block_of(graph, c) for c in ch}
            for b in touched:
                block_seen[b] += 1
                if fails:
                    block_fail[b] += 1
            case_rows.append(dict(case_id=case["id"], kind=case["kind"], changed_inputs=len(ch),
                                  outputs_checked=len(outputs), failures=len(fails), model_error=err or ""))
            if fails and len(failing_cases) < 10:
                failing_cases.append(dict(case=case, failures=[r for r in fail_rows if r["case_id"] == case["id"]][:15],
                                          model_error=err))
            if case["id"] and case["id"] % 50 == 0:
                log(f"  {case['id']}/{len(cases)}")
        engine.set(book, {c: _to_json(base[c]) for c in prev})
    finally:
        engine.close(book)

    passed_cases = sum(1 for r in case_rows if r["failures"] == 0)
    by_src = {src: [r for r, c in zip(case_rows, cases) if c["source"] == src] for src in ("excel", "replay")}
    rate = lambda rows: (sum(1 for r in rows if r["failures"] == 0) / len(rows)) if rows else None
    summary = dict(
        engine=engine.info, today=today.isoformat(), probe_date=probe_day.isoformat() if probe_day else None,
        tolerance=tol, cases=len(cases), outputs=len(outputs),
        excel_cases=len(by_src["excel"]), replay_cases=len(by_src["replay"]),
        excel_pass_rate=rate(by_src["excel"]), replay_pass_rate=rate(by_src["replay"]),
        pass_rate=passed_cases / len(cases) if cases else 0,
        outputs_all_pass=sum(1 for o in outputs if per_output[o]["fail"] == 0),
        outputs_failing={o: dict(per_output[o]) for o in outputs if per_output[o]["fail"]},
        rounding_only=sorted(rounding_only),
        failing_input_blocks={b: dict(fail=block_fail[b], changed=block_seen[b]) for b in block_fail},
        case_kinds=dict(Counter(c["kind"].split(":")[0] for c in cases)),
        resets_after_circular_errors=resets,
    )
    _write(run_dir / "proof_runs.csv", case_rows, ["case_id", "kind", "changed_inputs", "outputs_checked", "failures", "model_error"])
    _write(run_dir / "proof_failures.csv", fail_rows, ["case_id", "kind", "output", "excel", "model", "diff", "rounding_only"])
    (run_dir / "proof_summary.json").write_text(json.dumps(summary, indent=1, default=str))
    (run_dir / "proof_report.md").write_text(_report(summary, failing_cases, outputs, per_output))
    return summary


def _block_of(graph, cell):
    s, c, r = R.split_key(cell)
    for b in graph["input_blocks"]:
        if b["sheet"] == s and R.parse_cell(f"{b['col']}1")[0] == c and b["first_row"] <= r <= b["last_row"]:
            return b["range"]
    return cell


def _write(path, rows, fields):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def _report(s, failing_cases, outputs, per_output):
    pct = f"{s['pass_rate']:.1%}"
    rr = f"{s['replay_pass_rate']:.1%}" if s.get("replay_pass_rate") is not None else "n/a"
    xr = f"{s['excel_pass_rate']:.1%}" if s.get("excel_pass_rate") is not None else "n/a"
    lines = [f"# Proof report", "",
             f"**{pct} of {s['cases']} input sets matched on every output**: {s['excel_cases']} recalculated fresh in "
             f"Excel ({xr}) and {s['replay_cases']} probe runs replayed against Excel's recorded answers ({rr}, date "
             f"pinned to the probe day {s.get('probe_date')}). "
             f"({s['outputs_all_pass']}/{s['outputs']} outputs never mismatched). "
             f"Tolerance {s['tolerance']} on stored values. Engine: {s['engine'].get('engine')} "
             f"{s['engine'].get('version', '')}, locale {s['engine'].get('locale', '?')}, today pinned to {s['today']}.", "",
             f"Case mix: {s['case_kinds']}", ""]
    if s["outputs_failing"]:
        lines += ["## Outputs with mismatches", "", "| Output | Pass | Fail |", "|---|---|---|"]
        for o, c in sorted(s["outputs_failing"].items(), key=lambda kv: -kv[1].get("fail", 0)):
            lines.append(f"| {o} | {c.get('pass', 0)} | {c.get('fail', 0)} |")
        if s["rounding_only"]:
            lines += ["", f"Match only after rounding to 2 dp: {', '.join(s['rounding_only'])}"]
        lines += ["", "## Mismatches by input block changed", "", "| Input block | Failing cases | Cases where it changed |", "|---|---|---|"]
        for b, c in sorted(s["failing_input_blocks"].items(), key=lambda kv: -kv[1]["fail"]):
            lines.append(f"| {b} | {c['fail']} | {c['changed']} |")
        lines += ["", "## First failing cases", ""]
        for fc in failing_cases:
            case = fc["case"]
            ch = ", ".join(f"{k}={v!r}" for k, v in list(case["changes"].items())[:12])
            more = f" (+{len(case['changes']) - 12} more)" if len(case["changes"]) > 12 else ""
            lines.append(f"### Case {case['id']} ({case['kind']})")
            lines.append(f"Changed: {ch or 'nothing (saved inputs)'}{more}")
            if fc["model_error"]:
                lines.append(f"Model raised: `{fc['model_error']}`")
            for f in fc["failures"]:
                lines.append(f"- {f['output']}: Excel {f['excel']!r} vs model {f['model']!r}" + (f" (diff {f['diff']:.6g})" if f["diff"] else ""))
            lines.append("")
    else:
        lines += ["Every output matched on every case."]
    return "\n".join(lines) + "\n"
