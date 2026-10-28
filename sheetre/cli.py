"""Command line for the /reverse-engineer skill.

    python -m sheetre.cli init <workbook>          -> prints the run folder (copies the file into it)
    python -m sheetre.cli inventory <run>
    python -m sheetre.cli map <run>
    python -m sheetre.cli probe <run> [--coverage standard|exhaustive] [--cap N]
    python -m sheetre.cli try <run> --set "Sheet!A1=5" [--set ...] [--get "Sheet!B2"]
    python -m sheetre.cli ledger-check <run>
    python -m sheetre.cli prove <run> [--n 300] [--seed 1]
    python -m sheetre.cli report <run>
    python -m sheetre.cli status <run>
    python -m sheetre.cli engine-check

Exit codes: 0 ok, 1 error, 2 proof below 100%, 3 no recalculation engine (Probe/Prove must be marked UNKNOWN),
4 Excel busy with another stage (wait for it, then rerun).
"""
import argparse
import datetime as dt
import hashlib
import json
import shutil
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

ROOT = Path(__file__).resolve().parents[1]


def _run_json(run):
    return json.loads((Path(run) / "run.json").read_text())


def cmd_init(a):
    src = Path(a.workbook).expanduser().resolve()
    if not src.exists():
        sys.exit(f"not found: {src}")
    if src.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        sys.exit(f"unsupported file type {src.suffix} (need .xlsx/.xlsm)")
    from .engine import safe_name
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    run = Path(a.runs).resolve() / f"{stamp}_{safe_name(src.stem)}"
    (run / "work").mkdir(parents=True)
    (run / "rebuild").mkdir()
    work = run / "work" / safe_name(src.name)
    shutil.copy2(src, work)
    meta = dict(original=str(src), work_file=str(work), created=stamp,
                sha256=hashlib.sha256(src.read_bytes()).hexdigest(), size=src.stat().st_size)
    (run / "run.json").write_text(json.dumps(meta, indent=1))
    print(run)


def cmd_inventory(a):
    from . import inventory
    meta = _run_json(a.run)
    inv = inventory.run(meta["work_file"], a.run)
    print(json.dumps(dict(sheets=[(s["name"], s["state"]) for s in inv["sheets"]], vba=bool(inv["vba"]["modules"]),
                          calc=inv["calc"], flags=inv["summary"]), indent=1))


def cmd_map(a):
    from . import mapgraph
    meta = _run_json(a.run)
    g = mapgraph.build(meta["work_file"], a.run)
    print(json.dumps(dict(stats=g["stats"], output_rule=g["output_rule"], input_blocks=len(g["input_blocks"]),
                          dynamic=len(g["dynamic"]), cycles=len(g["cycles"]),
                          lookup_problems=sum(1 for l in g["lookups"] if l["problems"]),
                          aggregate_gaps=len(g["aggregate_gaps"])), indent=1))


def _engine(a):
    from .engine import EngineBusy, EngineUnavailable, get_engine
    try:
        return get_engine(prefer=getattr(a, "engine", None))
    except EngineBusy as e:
        print(f"EXCEL BUSY: {e}")
        sys.exit(4)
    except EngineUnavailable as e:
        print(f"NO RECALCULATION ENGINE: {e}\nMark Probe and Prove as UNKNOWN in the ledger and report.")
        sys.exit(3)


def cmd_probe(a):
    from . import probe
    meta = _run_json(a.run)
    eng = _engine(a)
    try:
        s = probe.run(a.run, eng, meta["work_file"], cap=a.cap, coverage=a.coverage)
    finally:
        eng.stop()
    print(json.dumps({k: s[k] for k in ("runs", "planned", "cap_hit", "nondeterministic_outputs",
                                        "date_or_random_functions", "inputs_with_no_effect")}, indent=1, default=str)[:3000])
    print(f"sweep findings: {len(s['sweeps'])} input(s); recalculated != cached: {len(s['recalculated_differs_from_cached'])}")


def cmd_ledger_check(a):
    from . import report
    problems = report.check_ledger(Path(a.run))
    if problems:
        print("LEDGER PROBLEMS:\n- " + "\n- ".join(problems))
        sys.exit(1)
    print("ledger ok")


def cmd_prove(a):
    from . import prove
    meta = _run_json(a.run)
    eng = _engine(a)
    try:
        s = prove.run(a.run, eng, meta["work_file"], n=a.n, seed=a.seed,
                       tol=prove.TOL if a.tol is None else a.tol)
    finally:
        eng.stop()
    print(json.dumps({k: s[k] for k in ("cases", "excel_cases", "replay_cases", "outputs", "pass_rate", "excel_pass_rate",
                                        "replay_pass_rate", "outputs_all_pass", "outputs_failing", "rounding_only")},
                     indent=1, default=str)[:3000])
    if s["pass_rate"] < 1:
        sys.exit(2)


def _parse_value(text):
    t = text.strip()
    if t == "":
        return None
    if t.startswith('"') and t.endswith('"'):
        return t[1:-1]
    if t.upper() in ("TRUE", "FALSE"):
        return t.upper() == "TRUE"
    if len(t) == 10 and t[4] == "-" and t[7] == "-":
        return {"date": t + "T00:00:00"}
    try:
        f = float(t)
        return int(f) if f.is_integer() and "." not in t else f
    except ValueError:
        return t


def cmd_try(a):
    """Ad-hoc probe: set cells, recalc, show what moved. Logged to try_log.jsonl in the run folder."""
    from .probe import same
    meta = _run_json(a.run)
    graph = json.loads((Path(a.run) / "graph.json").read_text())
    changes = {}
    for item in a.set or []:
        cell, _, val = item.partition("=")
        changes[cell.strip()] = _parse_value(val)
    ranges = [g for g in a.get or [] if ":" in g]
    if ranges:   # a range sent to Excel hangs the bridge for the rest of the session
        raise SystemExit(f"try --get takes single cells, not ranges: {ranges}. Repeat --get per cell.")
    watch = list(dict.fromkeys((a.get or []) + graph["outputs"]))
    eng = _engine(a)
    try:
        book = eng.open(meta["work_file"], Path(a.run) / "work")
        eng.calc(book, full=True)
        before = eng.get(book, watch)
        if changes:
            eng.set(book, changes)
            eng.calc(book)
        after = eng.get(book, watch)
        eng.close(book)
    finally:
        eng.stop()
    moved = {k: [before[k], after[k]] for k in watch if not same(before[k], after[k])}
    shown = {k: after[k] for k in (a.get or [])}
    rec = dict(set=changes, get=shown, moved=moved)
    with open(Path(a.run) / "try_log.jsonl", "a") as f:
        f.write(json.dumps(rec, default=str) + "\n")
    n = sum(1 for _ in open(Path(a.run) / "try_log.jsonl"))
    print(f"try #{n} (cite as 'try {n}' in evidence)")
    print(json.dumps(rec, indent=1, default=str)[:6000])


def cmd_report(a):
    from . import report
    paths = report.build(Path(a.run))
    print("\n".join(str(p) for p in paths))


def cmd_status(a):
    run = Path(a.run)
    for name in ("inventory.json", "graph.json", "map.md", "probes.csv", "probe_summary.json", "ledger.csv",
                 "rebuild/model.py", "proof_runs.csv", "proof_report.md", "narrative.md", "REPORT.md", "REPORT.html"):
        print(f"{'x' if (run / name).exists() else ' '} {name}")


def cmd_engine_check(a):
    """Report whether each engine is installed AND can actually be driven.

    Installed is not the same as usable: on a headless macOS session Excel answers basic
    queries but refuses to open a workbook, so detection alone used to report "found" and
    Probe then died mid-run. Pass --no-start to skip the handshake.
    """
    from .engine import BridgeEngine, EngineUnavailable, ExcelEngine, LibreOfficeEngine
    usable = 0
    for cls in (BridgeEngine, ExcelEngine, LibreOfficeEngine):
        ok, why = cls.detect()
        if not ok:
            print(f"{cls.name}: missing - {why}")
            continue
        if getattr(a, "no_start", False):
            print(f"{cls.name}: installed (not started)")
            continue
        eng = cls(timeout=60) if cls is ExcelEngine else cls()
        try:
            eng.start()
        except EngineUnavailable as e:
            print(f"{cls.name}: installed but UNUSABLE - {e}")
            continue
        except Exception as e:  # noqa: BLE001 - any failure here means unusable
            print(f"{cls.name}: installed but UNUSABLE - {e}")
            continue
        try:
            print(f"{cls.name}: usable")
            usable += 1
        finally:
            eng.stop()
    if not usable and not getattr(a, "no_start", False):
        print("No usable recalculation engine: Probe and Prove will exit 3. "
              "Inventory, Map and Ledger still run; mark anything needing recalculation UNKNOWN.")


def main(argv=None):
    p = argparse.ArgumentParser(prog="sheetre")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("init"); s.add_argument("workbook"); s.add_argument("--runs", default=str(ROOT / "runs"))
    s.set_defaults(fn=cmd_init)
    for name, fn in (("inventory", cmd_inventory), ("map", cmd_map), ("ledger-check", cmd_ledger_check),
                     ("report", cmd_report), ("status", cmd_status)):
        s = sub.add_parser(name); s.add_argument("run"); s.set_defaults(fn=fn)
    s = sub.add_parser("probe"); s.add_argument("run"); s.add_argument("--cap", type=int, default=None)
    s.add_argument("--coverage", choices=["standard", "exhaustive"], default="standard")
    s.add_argument("--engine", choices=["excel", "libreoffice"]); s.set_defaults(fn=cmd_probe)
    s = sub.add_parser("prove"); s.add_argument("run"); s.add_argument("--n", type=int, default=300)
    s.add_argument("--seed", type=int, default=1); s.add_argument("--tol", type=float, default=None)
    s.add_argument("--engine", choices=["excel", "libreoffice"]); s.set_defaults(fn=cmd_prove)
    s = sub.add_parser("try"); s.add_argument("run"); s.add_argument("--set", action="append", metavar="Sheet!A1=value")
    s.add_argument("--get", action="append", metavar="Sheet!A1"); s.add_argument("--engine", choices=["excel", "libreoffice"])
    s.set_defaults(fn=cmd_try)
    s = sub.add_parser("engine-check"); s.add_argument("--no-start", action="store_true",
        help="only check whether each engine is installed, do not try to drive it")
    s.set_defaults(fn=cmd_engine_check)
    a = p.parse_args(argv)
    a.fn(a)


if __name__ == "__main__":
    main()
