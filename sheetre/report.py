"""Ledger checks and the final REPORT.md / REPORT.html.

Numbers in the report (proof pass rate, counts) are read from the run's JSON/CSV files,
never typed by hand. Claude writes narrative.md (outputs + logic in plain English) and ledger.csv.
"""
import csv
import datetime as dt
import html
import json
import re
from pathlib import Path

LEDGER_FIELDS = ["id", "location", "finding", "category", "tag", "problem", "evidence", "settle_by", "ask"]
TAGS = {"SEEN", "INFERRED", "UNKNOWN"}
CELL_RE = re.compile(r"[A-Za-z_][\w .']*!\$?[A-Z]{1,3}\$?\d+|\brun[ _#]?\d+|\btry[ _#]?\d+|\bprobe|\bVBA\b|inventory|graph\.json|map\.md|"
                     r"proof|name:|defined name|externalLink|calcPr|conditional format|data validation", re.I)


def read_ledger(run):
    path = Path(run) / "ledger.csv"
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def check_ledger(run):
    run = Path(run)
    path = run / "ledger.csv"
    if not path.exists():
        return [f"{path} does not exist"]
    rows = read_ledger(run)
    problems = []
    with open(path, newline="") as f:
        header = next(csv.reader(f), [])
    missing = [c for c in LEDGER_FIELDS if c not in header]
    if missing:
        return [f"ledger.csv is missing columns {missing}; expected {LEDGER_FIELDS}"]
    ids = set()
    for i, r in enumerate(rows, start=2):
        rid = (r["id"] or "").strip()
        where = f"row {i} ({rid or 'no id'})"
        if not rid:
            problems.append(f"{where}: empty id")
        if rid in ids:
            problems.append(f"{where}: duplicate id")
        ids.add(rid)
        tag = (r["tag"] or "").strip().upper()
        if tag not in TAGS:
            problems.append(f"{where}: tag must be SEEN, INFERRED or UNKNOWN (got {r['tag']!r})")
        if (r["problem"] or "").strip().upper() not in ("Y", "N"):
            problems.append(f"{where}: problem must be Y or N")
        ev = (r["evidence"] or "").strip()
        if not ev:
            problems.append(f"{where}: evidence is empty")
        elif tag == "SEEN" and not CELL_RE.search(ev):
            problems.append(f"{where}: SEEN needs evidence that quotes a cell (Sheet!A1), a probe run, VBA or another run file")
        if tag == "INFERRED" and not (r["settle_by"] or "").strip():
            problems.append(f"{where}: INFERRED needs settle_by (what would confirm it)")
        if tag == "UNKNOWN" and not ((r["settle_by"] or "").strip() or (r["ask"] or "").strip()):
            problems.append(f"{where}: UNKNOWN needs settle_by or ask (who or what could settle it)")
        if not (r["finding"] or "").strip():
            problems.append(f"{where}: finding is empty")
    if not rows:
        problems.append("ledger.csv has no rows")
    if not problems:
        (run / "ledger.md").write_text(ledger_md(rows))
    return problems


def _cell(s):
    return (s or "").replace("|", "\\|").replace("\n", " ").strip()


def ledger_md(rows):
    order = {"Y": 0, "N": 1}
    rows = sorted(rows, key=lambda r: (order.get((r["problem"] or "").upper(), 2), r["id"]))
    lines = ["| ID | Tag | Problem | Finding | Location | Evidence | What would settle it | Ask |",
             "|---|---|---|---|---|---|---|---|"]
    for r in rows:
        lines.append("| " + " | ".join(_cell(r[k]) for k in ("id", "tag", "problem", "finding", "location", "evidence",
                                                             "settle_by", "ask")) + " |")
    return "\n".join(lines) + "\n"


def _load(run, name):
    p = Path(run) / name
    return json.loads(p.read_text()) if p.exists() else None


def build(run):
    run = Path(run)
    meta = _load(run, "run.json") or {}
    inv = _load(run, "inventory.json") or {}
    graph = _load(run, "graph.json") or {}
    probe = _load(run, "probe_summary.json")
    proof = _load(run, "proof_summary.json")
    narrative = (run / "narrative.md").read_text() if (run / "narrative.md").exists() else \
        "_narrative.md not written: outputs and logic are missing from this report._\n"
    rows = read_ledger(run) if (run / "ledger.csv").exists() else []

    name = Path(meta.get("original", inv.get("file", "workbook"))).name
    L = [f"# What this spreadsheet actually does", "",
         f"**{name}**  ", f"Analysed {dt.date.today().isoformat()} | sha256 `{meta.get('sha256', '')[:12]}` | "
         f"{len(inv.get('sheets', []))} sheets, {graph.get('stats', {}).get('formulas', '?')} formulas, "
         f"{graph.get('stats', {}).get('outputs', '?')} outputs, {graph.get('stats', {}).get('inputs', '?')} true inputs", ""]
    L += [_proof_line(proof, probe), ""]
    L += [narrative.strip(), ""]

    counts = {t: sum(1 for r in rows if r["tag"].upper() == t) for t in ("SEEN", "INFERRED", "UNKNOWN")}
    probs = sum(1 for r in rows if r["problem"].upper() == "Y")
    L += ["## The ledger", "",
          f"{len(rows)} findings: {counts['SEEN']} SEEN, {counts['INFERRED']} INFERRED, {counts['UNKNOWN']} UNKNOWN. "
          f"{probs} flagged as problems. SEEN = quoted from a cell or a probe run. INFERRED = evidence points there, "
          f"not confirmed. UNKNOWN = the file cannot answer it.", "", ledger_md(rows) if rows else "_No ledger._", ""]

    agenda = [r for r in rows if r["tag"].upper() == "UNKNOWN" or (r["tag"].upper() == "INFERRED" and r["ask"].strip())]
    L += ["## Unknowns: agenda for the owner", ""]
    if agenda:
        for i, r in enumerate(agenda, start=1):
            who = f" Ask: {r['ask'].strip()}." if r["ask"].strip() else ""
            settle = f" Settles it: {r['settle_by'].strip()}." if r["settle_by"].strip() else ""
            L.append(f"{i}. **{r['finding'].strip()}** ({r['id']}, {r['tag'].upper()}).{who}{settle}")
    else:
        L.append("_No open questions._")
    L.append("")

    L += ["## Proof", ""]
    if proof:
        def pct(x):
            return "n/a" if x is None else f"{x:.1%}"
        L.append(f"The rebuild (`rebuild/model.py`) was checked against the original on {proof['cases']:,} input sets: "
                 f"{proof.get('excel_cases', 0):,} random sets recalculated fresh in {proof['engine'].get('engine')} "
                 f"({pct(proof.get('excel_pass_rate'))} matched) and every one of the {proof.get('replay_cases', 0):,} probe runs "
                 f"replayed against Excel's recorded answers ({pct(proof.get('replay_pass_rate'))} matched). Overall "
                 f"{proof['pass_rate']:.1%}; {proof['outputs_all_pass']}/{proof['outputs']} outputs never mismatched "
                 f"(tolerance {proof['tolerance']}).")
        if proof["outputs_failing"]:
            L.append("")
            L.append("Outputs that still mismatch: " + ", ".join(f"{o} ({c.get('fail', 0)} cases)"
                                                                for o, c in list(proof["outputs_failing"].items())[:12]))
            L.append("See proof_report.md for the failing cases in full.")
    else:
        L.append("**UNKNOWN**: Prove did not run (no recalculation engine, or no rebuild), so the rebuild "
                 "was never checked against the original recalculated in Excel.")
        subs = [(rel, label) for rel, label in SUBSTITUTE_CHECKS if (run / rel).exists()]
        if subs:
            L += ["", "Checks that did run. These narrow the gap but do not close it:", ""]
            for rel, label in subs:
                head = _headline(run / rel)
                L.append(f"- **{label}** (`{rel}`){': ' + head if head else ''}")
    L.append("")

    candidates = ["inventory.json", "graph.json", "map.md", "probes.csv", "probe_summary.json", "ledger.csv",
                  "ledger.md", "narrative.md", "rebuild/model.py", "rebuild/model_clean.xlsx", "rebuild/ROUNDS.md",
                  "rebuild/basecase_report.md", "rebuild/thirdopinion_report.md", "proof_runs.csv", "proof_report.md"]
    present = [c for c in candidates if (run / c).exists()]
    L += ["## How this was produced", "",
          "Inventory -> map -> probe -> ledger -> rebuild -> prove, run by the `/reverse-engineer` skill. "
          "Files in this run folder: " + ", ".join(present) + ".", ""]
    if probe:
        L.append(f"Probe: {probe['runs']} recalculations ({'cap hit' if probe['cap_hit'] else 'all planned runs'}), "
                 f"engine {probe['engine'].get('engine')} {probe['engine'].get('version', '')}, locale {probe['engine'].get('locale', '?')}.")
    L.append("")
    md = "\n".join(L)
    (run / "REPORT.md").write_text(md)
    (run / "REPORT.html").write_text(to_html(md, f"What this spreadsheet actually does: {name}"))
    return [run / "REPORT.md", run / "REPORT.html"]


# Checks that can stand in for the proof when no engine is available. Not equivalents: the
# proof recalculates the ORIGINAL in Excel, these do not. Written by the rebuild stage.
SUBSTITUTE_CHECKS = (
    ("rebuild/basecase_report.md", "Base case against the values Excel itself saved in the file"),
    ("rebuild/thirdopinion_report.md", "Many input sets against a second, non-oracle engine"),
)


def _headline(path):
    """The summary numbers out of a substitute check's report: its bold lines, compacted."""
    out = []
    for ln in path.read_text().splitlines():
        ln = ln.strip()
        if "**" not in ln:
            continue
        if ln.startswith("- ") or (ln.startswith("**") and ln.rstrip(".").endswith("**")):
            out.append(re.sub(r"\*\*", "", ln.lstrip("- ")).rstrip("."))
        if len(out) == 4:
            break
    return "; ".join(out)


def _proof_line(proof, probe):
    if proof:
        return (f"> **Proof: {proof['pass_rate']:.1%} of {proof['cases']:,} input sets reproduced exactly** by the rebuild "
                f"({proof.get('excel_cases', 0):,} recalculated fresh in Excel, {proof.get('replay_cases', 0):,} probe runs replayed; "
                f"{proof['outputs_all_pass']}/{proof['outputs']} outputs always matched).")
    if probe:
        return "> **Proof: not run.** The logic below was probed but the rebuild was not verified."
    return "> **Proof: UNKNOWN.** No recalculation engine was available; nothing below was verified by recalculation."


# ---------------------------------------------------------------- tiny markdown -> html

def _inline(s):
    s = html.escape(s)
    s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\w)_([^_]+)_(?!\w)", r"<em>\1</em>", s)
    return s


def to_html(md, title):
    out, lines, i = [], md.split("\n"), 0
    in_list = None
    while i < len(lines):
        line = lines[i]
        if line.startswith("```"):
            j = i + 1
            while j < len(lines) and not lines[j].startswith("```"):
                j += 1
            out.append("<pre>" + html.escape("\n".join(lines[i + 1:j])) + "</pre>")
            i = j + 1
            continue
        if line.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s\-:|]+\|$", lines[i + 1]):
            head = [c.strip() for c in line.strip("|").split(" | ")]
            out.append("<div class='t'><table><thead><tr>" + "".join(f"<th>{_inline(h)}</th>" for h in head) + "</tr></thead><tbody>")
            i += 2
            while i < len(lines) and lines[i].startswith("|"):
                cells = re.split(r"(?<!\\)\|", lines[i].strip().strip("|"))
                out.append("<tr>" + "".join(f"<td>{_inline(c.strip().replace(chr(92) + '|', '|'))}</td>" for c in cells) + "</tr>")
                i += 1
            out.append("</tbody></table></div>")
            continue
        m = re.match(r"^(#{1,4}) (.*)", line)
        lm = re.match(r"^(\d+\.|-) (.*)", line)
        if in_list and not lm:
            out.append(f"</{in_list}>")
            in_list = None
        if m:
            n = len(m.group(1))
            out.append(f"<h{n}>{_inline(m.group(2))}</h{n}>")
        elif lm:
            kind = "ol" if lm.group(1)[0].isdigit() else "ul"
            if in_list != kind:
                if in_list:
                    out.append(f"</{in_list}>")
                out.append(f"<{kind}>")
                in_list = kind
            out.append(f"<li>{_inline(lm.group(2))}</li>")
        elif line.startswith("> "):
            out.append(f"<blockquote>{_inline(line[2:])}</blockquote>")
        elif line.strip():
            out.append(f"<p>{_inline(line)}</p>")
        i += 1
    if in_list:
        out.append(f"</{in_list}>")
    css = """
:root{--bg:#fbfaf7;--fg:#1d1d1b;--muted:#6b6a65;--line:#e4e1d8;--accent:#1f4e78;--code:#f1efe8}
@media (prefers-color-scheme: dark){:root{--bg:#161615;--fg:#ecebe6;--muted:#a3a29b;--line:#34332f;--accent:#8ab4e0;--code:#22211f}}
body{background:var(--bg);color:var(--fg);font:16px/1.55 -apple-system,Segoe UI,Inter,sans-serif;max-width:1100px;margin:0 auto;padding:24px 16px}
h1{font-size:28px;margin:0 0 8px}h2{margin-top:36px;border-bottom:1px solid var(--line);padding-bottom:6px}
blockquote{margin:16px 0;padding:12px 16px;border-left:4px solid var(--accent);background:var(--code)}
code,pre{background:var(--code);border-radius:4px;font:13px ui-monospace,Menlo,monospace}code{padding:1px 4px}
pre{padding:12px;overflow:auto}.t{overflow-x:auto}table{border-collapse:collapse;font-size:13px;width:100%}
th,td{border:1px solid var(--line);padding:6px 8px;vertical-align:top;text-align:left}th{background:var(--code)}
p,li{max-width:80ch}
"""
    return (f"<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
            f"<title>{html.escape(title)}</title><style>{css}</style></head><body>" + "\n".join(out) + "</body></html>\n")
