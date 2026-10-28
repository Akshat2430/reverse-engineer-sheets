---
name: reverse-engineer
description: Reverse engineer an inherited Excel workbook (.xlsx/.xlsm). Inventories everything hidden, maps inputs to outputs, probes the original by recalculating it, writes an evidence ledger tagged SEEN / INFERRED / UNKNOWN, rebuilds the logic in Python and proves the rebuild matches the original on hundreds of input sets. Use when someone asks what a spreadsheet really does, wants an inherited model audited, explained or rebuilt, or runs /reverse-engineer <file>.
argument-hint: <path to .xlsx or .xlsm>
---

# /reverse-engineer

Map every input-output pair, separate what you saw from what you guessed, rebuild the model, prove the numbers match.

## Ground rules (read first, follow throughout)

1. **Say this once, before starting:** "Workbook contents will be sent to the model provider. If this file is confidential, run this in an approved environment."
2. Work only on the file the user gave. Never write to it: `init` copies it into a run folder and every stage works on the copy.
3. No network calls. No installing packages.
4. **Ledger rule.** Every finding is exactly one of:
   - **SEEN**: you can quote it. Cite the cell and formula text, a probe run (`run 123`), a `try N`, a VBA line, or a run file.
   - **INFERRED**: say what evidence points there and what would confirm it (`settle_by`).
   - **UNKNOWN**: say who or what could settle it (`ask` / `settle_by`).
5. **Never upgrade an inference to a fact.** A tool that returns nothing is UNKNOWN, not "absent". Intent (why someone built it this way) is almost never SEEN.
6. Inspect hidden and very hidden sheets, macros, external links and IFERROR wrappers **before** any rebuild.
7. Rebuild the workbook **as it is, bugs included**. The proof checks fidelity to the original, not correctness. Fixes go in the ledger and the narrative.
8. **The toolkit is read-only for you**: `sheetre/`, `tests/` and this skill. If a tool is wrong or missing something, work around it inside the run folder and record it in `RUN/toolkit_issues.md` (what happened, evidence, suggested fix). Problems with the environment (no engine, a dialog, a slow machine) go there too, not in the workbook ledger.

## Tools

Run from the repo root. `PY` is `.venv/bin/python` (macOS/Linux) or `.venv\Scripts\python` (Windows).

| Command | What it does |
|---|---|
| `$PY -m sheetre.cli init "<file>"` | Copies the file into `runs/<timestamp>_<name>/` and prints that folder (RUN) |
| `$PY -m sheetre.cli inventory RUN` | `inventory.json`: sheets and states, names, links, VBA (`RUN/vba/`), calc settings, flags |
| `$PY -m sheetre.cli map RUN` | `graph.json` and `map.md`: dependency graph, outputs, true inputs, lookups, gaps, cycles |
| `$PY -m sheetre.cli probe RUN` | Recalculates the copy in Excel with one input changed at a time (thresholds on every input cell, other variants on sampled cells): `probes.csv`, `probe_summary.json`, `probe_moves.jsonl` |
| `$PY -m sheetre.cli try RUN --set "Sheet!A1=5" --get "Sheet!B2"` | One targeted experiment; cite it as `try N` |
| `$PY -m sheetre.cli ledger-check RUN` | Validates `ledger.csv` and writes `ledger.md` |
| `$PY -m sheetre.cli prove RUN [--n 300]` | Runs N random input sets through Excel and `rebuild/model.py`, replays every probe run against the rebuild, compares every output |
| `$PY -m sheetre.cli report RUN` | Builds `REPORT.md` and `REPORT.html` from the run files |

Exit code 3 from probe, try or prove means no recalculation engine. Run `engine-check` once to see why. If Excel is installed but this session can't drive it (common inside a sandbox), ask the user to run `.venv/bin/python -m sheetre.bridge` in a normal terminal; the engine uses the bridge automatically. If there's still no engine, continue with inventory, map and ledger, and mark anything that needed recalculation as UNKNOWN.

## Steps

1. **Init.** `init` the file and keep the RUN path.
2. **Inventory and map, then read before theorising.** Run both. Read `inventory.json` (start with `summary`, then every flag list) and all of `map.md`. Read every file in `RUN/vba/`. Do not form a theory until you have read them. Use `references/excel-gotchas.md` as the checklist of what each flag can mean, and look at the actual cells (small openpyxl snippets are fine) for anything unclear.
3. **Probe.** Run `probe` (it takes minutes). Read `probe_summary.json`: `thresholds` (steps and reversals at boundaries), `sweeps`, `nondeterministic_outputs`, `date_or_random_functions`, `inputs_with_no_effect`, `recalculated_differs_from_cached`. Use `try` to confirm or kill specific hypotheses. See `references/probe-recipes.md`.
4. **Ledger.** Write `RUN/ledger.csv` with columns `id,location,finding,category,tag,problem,evidence,settle_by,ask`, one row per finding, following `references/ledger-rules.md`. Run `ledger-check` until it passes. Cover the logic (what each output is), not only the problems.
5. **Rebuild** (only after the ledger exists).
   - `RUN/rebuild/model.py` defines `run(inputs, today) -> dict`. `inputs` maps `"Sheet!A1"` to its value for every true input in `graph.json`. Return every output cell in `graph.json["outputs"]`. Use `sheetre.excelfn` for Excel-exact behaviour: `Cells`, `vlookup`/`match` (exact Excel search, including unsorted approximate match), `sumif`, `countif`, `xl_round`, `num`, `n`, `value`, `datevalue` (locale), `iterate`, `XLError`. Make the code readable: one function per stage of the logic, with comments that cite cells.
   - `RUN/rebuild/build_clean.py` writes `RUN/rebuild/model_clean.xlsx`: the same logic laid out readably, with one assumptions sheet, labelled steps and a notes sheet listing the ledger's problems. Run it.
6. **Prove.** Run `prove` (default N=300). Read `proof_report.md`: it reports fresh Excel cases and replayed probe runs separately. For every mismatch, find out why, add or change a ledger row (a mismatch means you misunderstood something), fix `model.py` and prove again. Stop after 3 rebuild rounds. Log each round in `RUN/rebuild/ROUNDS.md` (what failed, what you learned, what changed). Unexplained mismatches are listed as UNKNOWN, never hidden.
7. **Narrative and report.** Write `RUN/narrative.md` with three sections: `## Outputs` (each output and what it means), `## How it works` (the logic in plain English, step by step, citing cells) and `## What would worry me` (the 3 to 5 findings that matter most, by ledger id). Then run `report`. The report leads with outputs and logic, then the ledger, then the unknowns as a meeting agenda, then the proof.
8. **Answer the user** briefly: the proof result, the top problems (by ledger id), how many open questions are on the agenda, and the path to `REPORT.html`.

## Time and scale

Probe runs a few hundred to 3,000 recalculations at about 0.2-0.6s each, so 5 to 30 minutes. Run it in the background and read the inventory and map while it works. Prove recalculates 300 random input sets (5 to 15 minutes) and replays every probe run against the rebuild in seconds. Use the defaults; don't lower them to save time. For a final, published proof the operator reruns `probe --coverage exhaustive` and `prove --n 1000`. **Stages share one Excel: never run `try` or `prove` while a background `probe` is still running.** A second stage gets exit code 4 ("Excel busy"); wait for the first to finish, then rerun it. Don't lower `--n` to save time: the proof is the point.

**Never give your final answer while a stage is still running.** This session may be non-interactive: ending your turn ends the session and kills any background stage, so the run is left without a proof. When you have nothing else to do, wait in the foreground for the stage's summary file, which is written only when it finishes: `while [ ! -f RUN/probe_summary.json ]; do sleep 20; done` (`proof_summary.json` for prove; Bash timeout 600000; repeat until it returns), then carry on through prove, narrative and report in the same turn. Don't wait with `pgrep`: it can't see processes inside a sandbox and reports the stage as finished. When rerunning prove, `proof_summary.json` already exists, so wait for its modification time to change instead. If Excel shows a dialog, the engine stops with `EngineBlocked`. Ask the user to close the dialog, then rerun that stage.
