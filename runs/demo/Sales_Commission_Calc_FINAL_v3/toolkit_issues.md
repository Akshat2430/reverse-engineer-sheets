# Toolkit / environment issues

## 1. `pgrep` unavailable in this sandbox, so the documented wait-loop fails

**What happened.** SKILL.md step "Time and scale" tells the operator to wait for a
background probe with `while pgrep -f "sheetre.cli probe" >/dev/null; do sleep 20; done`.
In this session `pgrep` cannot read the process list:

```
$ pgrep -f "sheetre.cli probe"
sysmon request failed with error: sysmond service not found
pgrep: Cannot get process list
exit 1
```

Because it exits non-zero, the documented loop falls straight through and would let the
session end (or a second stage start) while the probe was still running — exactly the
failure the instruction exists to prevent. `nice` is also blocked
(`(eval):1: nice(5) failed: operation not permitted`), harmless but noisy.

**Evidence.** Transcript of this run; `probe_moves.jsonl` was still growing
(341 -> 374 lines) at the moment `pgrep` reported "not running".

**Workaround used.** Waited on the stage's own output instead of the process table: poll
`probe_summary.json` (written only when probe finishes) and treat a stalled
`probe_moves.jsonl` line count as the liveness signal.

**Suggested fix.** Have each recalculating stage write a sentinel
(`RUN/.<stage>.running`, removed in a `finally:`), and change SKILL.md's wait-loop to poll
that file rather than `pgrep` — it works in sandboxes, needs no process-list access, and
cannot silently report "finished" when the stage is alive. Failing that, document a
fallback loop such as
`while [ ! -f RUN/probe_summary.json ]; do sleep 20; done`.

## 2. No usable local Excel/LibreOffice; the bridge carried the run

`engine-check` reported:

```
excel-bridge: usable
excel: installed but UNUSABLE - [Errno 1] Operation not permitted: '/tmp/sheetre-excel.lock'
libreoffice: installed but UNUSABLE - LibreOffice exited before its worker started (code 1)
```

Not a defect: the sandbox denies `/tmp` writes, which is what the bridge is for. Recorded
only so the proof's provenance is clear — every recalculation in this run went through the
operator's `sheetre.bridge` process, not a sandbox-local engine.

## 3. `excelfn.vlookup` does not reproduce Excel's approximate match when a key is blank

**What happened.** Round 1 of the proof failed on exactly four of 3,000 input sets, all four
the `blank` variant on a key in `Rates!A2:A10` (probe runs 2191, 2193, 2197, 2201).

`excelfn._rank` maps `None` to `(0, 0)`, so a blanked key is ranked as the number 0 and
`_approx_pos` treats it as a genuine band starting at zero. Excel does not. On landing on an
empty cell Excel steps down to the nearest non-empty key at or above the current floor and
compares that one, while still narrowing from the midpoint it originally picked; with no
non-empty key in reach the lookup returns `#N/A`.

**Evidence.** Excel's own recorded answers. With `Rates!A6` blank a 190,000 deal takes 0.06
and a 1,038,000 deal takes 0.1; with `Rates!A10` blank a 149,000 deal takes 0.05 and a
1,038,000 deal takes 0.11. Neither "blank sorts last" nor "drop the blank-keyed row"
reproduces both - the per-case table is in `rebuild/ROUNDS.md`. The docstring's claim of
700/700 random cases against Excel still stands; those cases evidently had no blank keys.

**Workaround used.** `_base_rate()` in `rebuild/model.py`, used for `Calc!H` only. Verified
equivalent to `excelfn.vlookup` on 40,000 fuzzed no-blank cases, so it changes nothing except
the blank-key path, and it takes all 2,699 replays to 2,699.

**Suggested fix.** Teach `_approx_pos` the skip-down step and give `_rank` a blank that is not
0 for lookup purposes (`eq`/`lt` can keep treating blank as 0 - that part matches Excel's
comparison operators). Worth a test fixture with blanks interleaved into a lookup column,
since the current suite appears not to cover them.

## 4. Excel wedged part-way through, and `prove` could not be rerun

**What happened.** `try` with a *range* in `--get` (`--get "Calc!H2:H61"`) did not return; the
bridge timed out after 240s with `EngineBlocked: Excel is probably showing a dialog`. Every
later call failed differently and permanently:

```
EngineError: bridge 'hello': EngineUnavailable: Excel is installed but will not accept
automation in this session (OSERROR: -1708 MESSAGE: Application could not handle this
command. COMMAND: app(pid=644).display_alerts.set(False))
```

Single-cell `--get` had worked all run (tries 1-9), so the range form looks like the trigger.
`engine-check` offers no second engine here (sandbox-local Excel cannot take its lock,
LibreOffice will not start), so no stage needing recalculation could be rerun. Round 2 was
verified offline instead - see `rebuild/ROUNDS.md`.

**Also:** after the bridge reported itself usable again, `prove` still could not open the workbook (`EngineBlocked` on `open`, 240s) - a dialog was left sitting on the file. The run ends with round 1's proof on record.

**Suggested fix.** Two things. (1) If `--get` is meant to take only single cells, reject a
range up front with a clear message instead of sending it to Excel; if it is meant to take
ranges, `cmd_try` needs to expand them. (2) `EngineBlocked` currently leaves Excel
unrecoverable for the rest of the session - a dialog-clearing or restart path in the bridge
(it already restarts its worker for other failures) would save the run, since at that point
every remaining stage is dead.
