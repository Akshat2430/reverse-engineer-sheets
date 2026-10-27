# Pending skill updates from the Team_Budget_Control run (2026-10-27)

The sandbox blocks writes to `.claude/skills/`, so the two self-annealing edits below were
not applied. Everything else from this run is already committed in `sheetre/` and `tests/`.
Apply these by hand, or rerun with write access to `.claude/skills/`.

Code changes already made (no action needed):

- `sheetre/excelfn.py` - `xl_round` now rounds the 15-significant-digit value instead of
  going through `repr()`; added `_approx15`.
- `tests/test_excelfn.py` - two regression tests for the above.
- `sheetre/engine.py` - a failed Excel handshake now raises `EngineUnavailable` (condensed
  message) so callers fall through to the no-engine path instead of a raw traceback.
- `sheetre/cli.py` - `engine-check` starts each engine to test it is actually usable;
  added `--no-start`.
- `sheetre/report.py` - the Proof section lists substitute checks when `prove` did not run;
  the footer lists only files that exist.

---

## 1. `SKILL.md` - replace the "Exit code 3" paragraph under **Tools**

Replace this line:

> Exit code 3 from probe, try or prove means no recalculation engine. Continue with inventory, map and ledger, and mark anything that needed recalculation as UNKNOWN.

with:

```markdown
Exit code 3 from probe, try or prove means no recalculation engine. Continue with inventory, map and ledger, and mark anything that needed recalculation as UNKNOWN.

**Check the engine before the long stages:** `$PY -m sheetre.cli engine-check` starts each
engine to see whether it can actually be driven, not just whether it is installed. Excel
being installed is not enough -- in a headless or non-interactive session (no window
server) it answers `get version` but refuses `display_alerts` and `open workbook` with
`OSERROR -10004, A privilege violation occurred`. `engine-check` reports that as
`installed but UNUSABLE`; probe and prove then exit 3 rather than dying mid-run. The fix
is to rerun those stages from an interactive macOS terminal that has Automation consent
for Excel. (Use `--no-start` for a presence-only check.)

**When there is no engine, do these two instead of a proof,** and never call either one a
proof:

1. **Base case against Excel's own saved values.** The cached values in the workbook were
   written by Excel, so they are the real oracle for one input set -- the file's own.
   Compare every output in `graph.json["outputs"]` against them. Write
   `rebuild/basecase_report.md`. Getting floating-point residue (a total saved as
   `...999999999`) to match digit-for-digit is strong evidence the rebuild does the
   arithmetic in the same order as the sheet.
2. **Many input sets against the `formulas` library** (already installed), which phase 0
   records as a *third-opinion engine, not the oracle*. Include the corner cases your
   ledger rows predict -- zeroed inputs, a duplicated lookup key, a negative quantity --
   so a static "this mechanism is fragile" row becomes a demonstrated effect. Write
   `rebuild/thirdopinion_report.md`. Agreement means two reimplementations agree; it is
   corroboration, not Excel fidelity. Where the two disagree, say so in the ledger and
   mark it settle-by-Excel rather than picking a winner silently.

`report` surfaces both under **Proof** when `prove` did not run, labelled as checks that
narrow the gap without closing it.
```

## 2. `references/excel-gotchas.md` - add under **Formulas**, after the "Rounding" bullet

```markdown
- **ROUND sees 15 significant digits, not the binary value.** `ROUND(2.675, 2)` is `2.68` in Excel even though 2.675 is stored as 2.67499999999999982...: to 15 significant digits that value *is* 2.675, which then rounds half away from zero. So a pre-rounded result sitting just below a `.xx5` boundary can still round up. Use `excelfn.xl_round`, which applies that rule (`_approx15`); do not round via `repr()`, which returns the shortest round-tripping string and makes "is this an exact midpoint" an accident of formatting -- two values in the same relation to their midpoint then round in opposite directions. The `formulas` library is inconsistent on these values, so a one-unit-in-the-last-place disagreement at a `.xx5` boundary is expected and must be settled in Excel, not by trusting the third opinion.
```
