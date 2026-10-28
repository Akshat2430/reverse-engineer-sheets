# Rebuild rounds

## Round 1 - `prove --n 300`, 2026-10-28

**Result: 99.9% of 3,000 input sets.** 301 fresh Excel cases 301/301 (100.0%); 2,699 replayed
probe runs 2,695/2,699. `proof_report.md` and `proof_failures.csv` hold this run as recorded.

**What failed.** Four replay cases, every one of them the `blank` variant on a key in the
base-rate band column: probe runs 2191 (`Rates!A2=blank`), 2193 (`A5`), 2197 (`A6`) and
2201 (`A10`). Nothing else in the workbook mismatched: no rounding-only rows, no date
problems, no error-propagation problems. `proof_report.md` attributes all four to the one
input block, `Rates!A2:A10`.

**What I learned.** The mismatch was mine, not the workbook's: I had assumed
`excelfn.vlookup` covered the lookup completely, and it does not cover a *blank* key.
`excelfn._rank` ranks a blank as the number 0, so a blanked band key reads as a real band at
zero. Excel does something else, and two plausible readings each explained only part of the
evidence:

| Hypothesis | A2 | A5 | A6 | A10 |
|---|---|---|---|---|
| blank ranks as 0 (`excelfn.vlookup`, round 1) | fail | fail | fail | fail |
| blank sorts after every number | pass | pass | **fail** | pass |
| blank-keyed rows dropped from the table | pass | pass | pass | **fail** |

Neither fits, so I stopped guessing and recovered Excel's actual per-deal base rate for the
two disputed runs. Because rows 40-61 are frozen constants (ledger L11) only rows 2-39 can
move a Summary rep line, and each rep has at most four such deals, so the rate each deal
actually used can be solved for from the twelve rep totals Excel recorded. That gave, for
example, `Rates!A6` blank: a 190,000 deal took 0.06 and a 1,038,000 deal took 0.1; and
`Rates!A10` blank: a 149,000 deal took 0.05 and a 1,038,000 deal took 0.11.

Those numbers pin the rule. Excel's binary search, on landing on an empty cell, **steps down
to the nearest non-empty key at or above the current floor and compares that one, while still
narrowing from the midpoint it originally picked**; if there is no non-empty key left in
reach, the lookup fails with `#N/A`. A blank therefore neither blocks the search from
descending past it (which is why 1,038,000 still reaches 0.1 in the `A6` case) nor vanishes
from the index arithmetic (which is why 149,000 still reaches 0.05 in the `A10` case).

**What changed.** Added `_base_rate()` to `rebuild/model.py` implementing that search, and
pointed `Calc!H` at it instead of `excelfn.vlookup`. It is the only approximate-match lookup
in the workbook, so nothing else is touched. Logged as issue 3 in `toolkit_issues.md`.
Added ledger row **L26** for the workbook behaviour this uncovered - clearing a band key
silently reprices mid-size deals, and clearing the top one turns the whole Summary into
`#N/A`.

## Round 2 - re-verification, 2026-10-28

`prove` could **not** be re-run: part-way through the follow-up experiments the bridge's Excel
stopped accepting automation (`EngineBlocked` on a dialog, then
`OSERROR: -1708 ... display_alerts.set(False)` on every later call), and no second engine is
available in this sandbox. See issue 4 in `toolkit_issues.md`. So round 2 was verified without
Excel, in the two halves the proof is made of:

- **Replay half, re-run in full offline.** Replaying a probe run needs no recalculation: the
  proof compares the rebuild against the answer Excel already recorded in
  `probe_moves.jsonl`. Re-ran all 2,699 through `prove.replay_cases` and `prove.compare` at
  the same tolerance: **2,699/2,699 cases match on every output**, the four failures included.
- **Fresh-Excel half, carried over.** The fix cannot change those 301 results, on two
  independent checks: (1) `_base_rate` is byte-for-byte equivalent to `excelfn.vlookup`
  whenever no key is blank - 40,000/40,000 fuzzed tables and lookup values agree, sorted and
  unsorted, hits, misses and duplicates; and (2) none of the 301 cases blanks anything at
  all - `prove.make_cases` with the same seed yields 0 blank values across all 301, because
  `_random_value` always returns a number for a numeric block. The saved-inputs case was
  re-checked directly against the cached workbook: **43/43**.

**Standing result: 3,000/3,000 input sets, every output.** The recorded `proof_report.md` is
still round 1's 99.9% and has deliberately not been edited or regenerated - a proof report
should say what the engine actually did. **The operator should rerun `prove` once Excel is
healthy to put a 100% report on record**, together with the published
`probe --coverage exhaustive` and `prove --n 1000`.

### Round 2 rerun of `prove`: attempted, blocked

Once `engine-check` reported `excel-bridge: usable` again, `prove` was rerun to put a clean
report on record. It got no further than opening the workbook:

```
EngineBlocked: bridge 'open': Excel did not answer 'open' within 240s. Excel is probably
showing a dialog (VBA error, Save As, update links, repair).
```

The bridge answers `hello` but not `open`, so a modal dialog is sitting on the workbook and
only a person at that machine can clear it. Round 1's `proof_report.md`,
`proof_summary.json`, `proof_runs.csv` and `proof_failures.csv` are therefore left in place,
unedited, as the recorded proof. **Close the dialog in Excel and rerun
`prove RUN` to put the 100% result on record.**
