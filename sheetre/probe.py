"""Stage 3: Probe. Recalculate the original with one input changed at a time.

Variants per sampled input cell:
  numbers: zero, blank, negative, x10, each downstream threshold T-1 / T / T+1, and a sweep
  text:    blank, other values seen in the same block, value + trailing space
  dates:   blank, +-31 days, +92 days
  dropdown cells (data validation lists): every allowed value
The base case runs twice; any output that differs between the two is non-deterministic.
Writes probes.csv (moved outputs only), probe_runs.csv (every run) and probe_summary.json.
"""
import csv
import datetime as dt
import json
import random
import re
from collections import defaultdict
from pathlib import Path

from . import refs as R
from .engine import EngineError

SEED = 11


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def same(a, b, tol=1e-9):
    if _num(a) and _num(b):
        return abs(a - b) <= tol * max(1.0, abs(a), abs(b))
    return a == b


def _date(v):
    if isinstance(v, dict) and "date" in v:
        return dt.datetime.fromisoformat(v["date"])
    return v if isinstance(v, (dt.date, dt.datetime)) else None


def _dv_lists(inventory, graph):
    """{cell: [allowed values]} for list validations that cover true inputs."""
    out = {}
    inputs = set(graph["inputs"])
    for dv in inventory.get("flags", {}).get("data_validations", {}).get("items", []):
        if dv.get("type") != "list" or not dv.get("formula1"):
            continue
        f = dv["formula1"]
        if f.startswith('"'):
            values = [x.strip() for x in f.strip('"').split(",")]
        else:
            ref = R.parse_ref(f.lstrip("="), dv["sheet"])
            values = [graph["nodes"].get(k, {}).get("value") for k in ref.cells(max_row=5000, max_col=200)] \
                if ref.kind in ("cell", "range") else []
            values = [v for v in values if v not in (None, "")]
        for part in str(dv["sqref"]).split():
            r = R.parse_ref(part, dv["sheet"])
            if r.kind in ("cell", "range"):
                for k in r.cells(max_row=r.r2 or 1, max_col=r.c2 or 1, limit=20000):
                    if k in inputs:
                        out[k] = values
    return out


def plausible_thresholds(blk):
    """Comparison literals downstream of a block that are on the block's own scale."""
    lo, hi = blk.get("min"), blk.get("max")
    if lo is None:
        return []
    out = []
    for t in blk.get("thresholds", []):
        if lo >= 0 and hi > 0:
            if hi / 50 <= abs(t) <= hi * 5 or (lo <= t <= hi and t != 0):
                out.append(t)
        elif min(lo, 0) - abs(hi) <= t <= hi + abs(hi):
            out.append(t)
    return out


def plan(graph, inventory, cap=2000, sweep_points=12):
    """List of (input_cell, variant, value) single-change runs."""
    rng = random.Random(SEED)
    nodes = graph["nodes"]
    dv = _dv_lists(inventory, graph)
    runs, seen = [], set()

    def add(cell, variant, value):
        key = (cell, variant, json.dumps(value, default=str))
        if key not in seen:
            seen.add(key)
            runs.append((cell, variant, value))

    for cell, values in dv.items():
        for v in values:
            if v != nodes[cell]["value"]:
                add(cell, f"dropdown={v}", v)

    for blk in graph["input_blocks"]:
        cells = [R.cell_key(blk["sheet"], R.parse_cell(f"{blk['col']}1")[0], r)
                 for r in range(blk["first_row"], blk["last_row"] + 1)]
        cells = [c for c in cells if c in nodes]
        if not cells:
            continue
        # small blocks (assumption lists, rate tables) are often distinct parameters: probe every cell
        sample = cells if len(cells) <= 8 else [cells[0], cells[len(cells) // 2], cells[-1], rng.choice(cells[1:-1])]
        for i, c in enumerate(dict.fromkeys(sample)):
            base = nodes[c]["value"]
            d = _date(base)
            if _num(base):
                for variant, v in (("zero", 0), ("blank", None), ("negative", -base), ("x10", base * 10)):
                    if v != base:
                        add(c, variant, v)
                lo, hi = blk.get("min", base), blk.get("max", base)
                if i < 3:
                    for t in plausible_thresholds(blk):
                        step = 1 if blk.get("integers", True) and abs(t) >= 2 else max(abs(t) * 0.01, 1e-6)
                        for variant, v in (("threshold-1", t - step), ("threshold", t), ("threshold+1", t + step)):
                            add(c, f"{variant}({t:g})", v)
                if i == 0:
                    if hi > lo:
                        pts = sorted({lo + (hi - lo) * k / (sweep_points - 1) for k in range(sweep_points)} |
                                     {t for t in blk.get("thresholds", []) if lo <= t <= hi})
                        for v in pts:
                            add(c, "sweep", round(v, 6) if not blk.get("integers") else round(v))
            elif d is not None:
                add(c, "blank", None)
                for days in (-31, 31, 92):
                    add(c, f"date{days:+d}d", {"date": (d + dt.timedelta(days=days)).isoformat()})
            elif isinstance(base, bool):
                add(c, "flip", not base)
            elif isinstance(base, str):
                add(c, "blank", None)
                add(c, "trailing-space", base + " ")
                others = [x for x in blk.get("distinct", []) if x != base][:3]
                for o in others:
                    add(c, "swap", o)
    truncated = len(runs) > cap
    return runs[:cap], truncated, len(runs)


def run(run_dir, engine, work_file, cap=2000, log=print):
    run_dir = Path(run_dir)
    graph = json.loads((run_dir / "graph.json").read_text())
    inventory = json.loads((run_dir / "inventory.json").read_text())
    outputs = graph["outputs"]
    runs, truncated, planned = plan(graph, inventory, cap=cap)
    book = engine.open(work_file, run_dir / "work")
    nodes = graph["nodes"]
    try:
        engine.calc(book, full=True)
        base1 = engine.get(book, outputs)
        engine.calc(book, full=True)
        base2 = engine.get(book, outputs)
        nondet = [o for o in outputs if not same(base1[o], base2[o])]
        cached_mismatch = [o for o in outputs if not same(base1[o], _cached(nodes[o]["value"]), tol=1e-6)]
        moved_rows, run_rows = [], []
        moved_by_input = defaultdict(lambda: defaultdict(float))
        sweeps = defaultdict(list)
        thresh_pts = defaultdict(dict)
        restore = {}
        log(f"probe: {len(runs)} runs planned ({planned} before cap), {len(outputs)} outputs")
        for i, (cell, variant, value) in enumerate(runs, start=1):
            change = dict(restore)
            change[cell] = value
            try:
                engine.set(book, change)
                engine.calc(book)
                after = engine.get(book, outputs)
            except EngineError as e:
                run_rows.append(dict(run_id=i, input=cell, variant=variant, value=value, moved=0, error=str(e)[:200]))
                raise
            restore = {cell: nodes[cell]["value"]}
            missing = [o for o in outputs if o not in after]
            if missing:
                raise EngineError(f"run {i} ({cell} {variant}={value!r}): engine returned no value for {missing[:5]}")
            moved = [o for o in outputs if not same(after[o], base1[o])]
            run_rows.append(dict(run_id=i, input=cell, variant=variant, value=_show(value), moved=len(moved), error=""))
            for o in moved:
                b, a = base1[o], after[o]
                delta = (a - b) if _num(a) and _num(b) else ""
                moved_rows.append(dict(run_id=i, input=cell, variant=variant, value=_show(value), output=o,
                                       before=b, after=a, delta=delta))
                moved_by_input[cell][o] = max(moved_by_input[cell][o], abs(delta) if delta != "" else float("inf"))
            if variant == "sweep":
                sweeps[cell].append((value, {o: after[o] for o in outputs}))
            m = re.match(r"threshold([+-]1)?\((.+)\)$", variant)
            if m:
                thresh_pts[(cell, float(m.group(2)))][m.group(1) or "0"] = {o: after[o] for o in outputs}
            if i % 100 == 0:
                log(f"  {i}/{len(runs)}")
        engine.set(book, restore)
    finally:
        engine.close(book)

    _write_csv(run_dir / "probes.csv", moved_rows,
               ["run_id", "input", "variant", "value", "output", "before", "after", "delta"])
    _write_csv(run_dir / "probe_runs.csv", run_rows, ["run_id", "input", "variant", "value", "moved", "error"])
    volatile = [g for g in inventory.get("flags", {}).get("volatile", {}).get("items", [])
                if set(g.get("functions", [])) & {"TODAY", "NOW", "RAND", "RANDBETWEEN", "RANDARRAY"}]
    summary = dict(
        engine=engine.info, runs=len(runs), planned=planned, cap=cap, cap_hit=truncated,
        nondeterministic_outputs=nondet,
        date_or_random_functions=[dict(range=g.get("range"), functions=g.get("functions")) for g in volatile],
        note_on_dates=("TODAY()/NOW() return the same value in both base runs on one day, so the twice-run check cannot "
                       "see date dependence. Treat these cells as date-dependent; Prove pins today's date.") if volatile else None,
        recalculated_differs_from_cached=cached_mismatch[:50],
        inputs_with_no_effect=sorted({c for c, _, _ in runs} - set(moved_by_input)),
        inputs_not_probed=sorted(set(graph["inputs"]) - {c for c, _, _ in runs}),
        sensitivity={c: dict(sorted(v.items(), key=lambda kv: -kv[1])[:10]) for c, v in moved_by_input.items()},
        sweeps=_sweep_findings(sweeps, outputs),
        thresholds=_threshold_findings(thresh_pts, outputs, nodes),
    )
    (run_dir / "probe_summary.json").write_text(json.dumps(summary, indent=1, default=str))
    return summary


def _cached(v):
    return _date(v) if isinstance(v, dict) else v


def _show(v):
    if isinstance(v, dict) and "date" in v:
        return v["date"][:10]
    return v


def _write_csv(path, rows, fields):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def _sweep_findings(sweeps, outputs):
    """For each swept input, outputs that jump or move against the input's direction."""
    out = {}
    for cell, pts in sweeps.items():
        pts.sort(key=lambda p: p[0])
        notes = []
        for o in outputs:
            ys = [p[1][o] for p in pts]
            if not all(_num(y) for y in ys) or len(set(ys)) <= 1:
                continue
            diffs = [b - a for a, b in zip(ys, ys[1:])]
            ups, downs = sum(d > 1e-9 for d in diffs), sum(d < -1e-9 for d in diffs)
            if ups and downs:
                k = next(j for j, d in enumerate(diffs) if (d > 0) != (diffs[0] > 0) and abs(d) > 1e-9)
                notes.append(dict(output=o, pattern="non-monotonic", first_reversal_between=[pts[k][0], pts[k + 1][0]]))
        if notes:
            out[cell] = dict(points=[p[0] for p in pts], findings=notes[:10])
    return out


def _threshold_findings(points, outputs, nodes):
    """At each threshold T: compare outputs at T-1, T, T+1. Flags a step exactly at T (a >= boundary),
    a step just after T (a > boundary), and reversals (input goes up by one step, output goes down)."""
    out = []
    for (cell, t), pts in points.items():
        if not all(k in pts for k in ("-1", "0", "+1")):
            continue
        for o in outputs:
            a, b, c = pts["-1"][o], pts["0"][o], pts["+1"][o]
            if not all(_num(x) for x in (a, b, c)):
                continue
            j1, j2 = b - a, c - b
            small = max(abs(j1), abs(j2)) * 0.05
            kind = None
            if (j1 > 1e-9 and j2 < -1e-9) or (j1 < -1e-9 and j2 > 1e-9):
                kind = "reversal"
            elif abs(j1) > 20 * max(abs(j2), 1e-9) and abs(j1) > 1e-6:
                kind = "step at T (behaves like >=)"
            elif abs(j2) > 20 * max(abs(j1), 1e-9) and abs(j2) > 1e-6:
                kind = "step just above T (behaves like >)"
            if kind:
                out.append(dict(input=cell, threshold=t, output=o, kind=kind, at_t_minus_1=a, at_t=b, at_t_plus_1=c))
    # keep it readable: one line per (input, threshold, kind) with the outputs listed
    grouped = {}
    for f in out:
        k = (f["input"], f["threshold"], f["kind"])
        g = grouped.setdefault(k, dict(input=f["input"], threshold=f["threshold"], kind=f["kind"], outputs=[], example=None))
        g["outputs"].append(f["output"])
        if g["example"] is None:
            g["example"] = {key: f[key] for key in ("output", "at_t_minus_1", "at_t", "at_t_plus_1")}
    res = list(grouped.values())
    for g in res:
        g["outputs"] = g["outputs"][:8]
    return res
