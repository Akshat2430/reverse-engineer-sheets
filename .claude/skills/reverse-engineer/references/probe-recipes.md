# Probe recipes

`probe` runs these automatically on every cell of blocks with 8 cells or fewer, and on 4 sampled cells (first, middle, last, one random) of larger blocks. Each run changes one input, recalculates the copy in Excel and records every output that moved (`probes.csv`). The base case runs twice.

| Input type | Variants |
|---|---|
| Number | zero, blank, negative, x10, every downstream comparison threshold T-1 / T / T+1 (on up to 3 sampled cells per block), and a 12-point sweep across the observed range |
| Text | blank, value + trailing space, up to 3 other values seen in the same column |
| Date | blank, -31, +31 and +92 days |
| Dropdown cell | every value in its validation list |
| TRUE/FALSE | flipped |

## Reading `probe_summary.json`

- `thresholds`: for each (input, threshold) pair, outputs that step **exactly at T** (behaves like `>=`), step **just above T** (behaves like `>`), or **reverse** (input up by one step, output down). Different behaviour at the same threshold on different rows is a boundary inconsistency. A reversal next to a lookup usually means approximate match on unsorted keys.
- `sweeps`: outputs that are non-monotonic as one input rises. Some ratios are naturally non-monotonic; check the formula before calling it a problem.
- `nondeterministic_outputs`: differ between two identical base runs (RAND and the like).
- `date_or_random_functions`: TODAY()/NOW() cells. Both base runs fall on the same day, so they look deterministic. Treat them as date-dependent. Prove pins `today`.
- `inputs_not_probed`: true inputs the sampler skipped (large blocks are sampled, blocks of 8 cells or fewer are probed in full). Absent from `sensitivity` does not mean unimportant; use `try` on any that matter.
- `inputs_with_no_effect`: sampled inputs that moved nothing. Either dead inputs, or effects that only show in combination with another input's value.
- `recalculated_differs_from_cached`: the saved numbers are not what the file computes today. Stale values, volatile functions, or a manual-calculation workbook.
- `sensitivity`: largest output moves per input.

## Targeted experiments with `try`

```
$PY -m sheetre.cli try RUN --set "Data!F30=5000" --get "Model!K30" --get "Dashboard!C9"
$PY -m sheetre.cli try RUN --set "Data!B30=\"North\"" --set "Data!F30=5000" --get "Model!K30"
$PY -m sheetre.cli try RUN --set "Dashboard!D3=Mar"
```

- Text values: wrap them in quotes if they look like numbers. Dates: `2026-03-31`. Blank: `Sheet!A1=`.
- Every `try` is logged to `try_log.jsonl` and numbered. Cite it as `try N` in ledger evidence.
- Use `try` to turn an INFERRED mechanism into a SEEN effect. It never settles intent.

Patterns worth testing:

1. **Boundary:** the same threshold value on two rows that take different branches (e.g. different regions).
2. **Lookup fragility:** a key with a trailing space, a key in different case, a number stored as text.
3. **Dynamic references:** change the selector cell an OFFSET/INDIRECT depends on and see which outputs move.
4. **Overrides:** blank a cell on a hidden sheet and see whether a total changes.
5. **Stale values:** change an input that feeds a column of typed numbers inside a formula run. If the typed cells don't move, something (often a macro) wrote them.
6. **Circularity:** move an input of a circular group. Outputs should converge; compare against `iterate()` in the rebuild.

## No engine

If `probe` exits with code 3, no recalculation is possible. Static findings still stand. Anything that needed a recalculation becomes UNKNOWN with `settle_by: "recalculate in Excel"`.
