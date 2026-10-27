# Ledger rules

`RUN/ledger.csv`, one row per finding. Columns, in this order:

| Column | What goes in it |
|---|---|
| `id` | L01, L02, ... (stable once written) |
| `location` | Cells or objects: `Model!H2:H90`, `Lookup!A2:B12`, `name:TaxRate`, `VBA:Module1.Macro`, `sheet:Hidden1`, `externalLink1` |
| `finding` | One sentence a finance lead understands. What it does, not what you think of it. |
| `category` | Short slug, e.g. `lookup_unsorted_approx`, `hardcoded_literal`, `hidden_logic` (see list below) |
| `tag` | `SEEN`, `INFERRED` or `UNKNOWN` |
| `problem` | `Y` or `N`, see "What counts as a problem" below |
| `evidence` | What you quote: cell + formula text, probe run ids, `try N`, VBA lines, inventory or map entries |
| `settle_by` | What would confirm or settle it (required for INFERRED; for UNKNOWN unless `ask` is filled) |
| `ask` | Who could answer: role, not a guess at a name ("model owner", "finance controller") |

## The three tags

**SEEN**: it is in the file or a recalculation showed it. You can point at it.
- "The bonus formula pays 5% when sales `>= 1,000` but the next band starts at `> 5,000`" (formula text quoted): SEEN.
- "Setting units to 0 moves gross margin by -4,210" (`run 12`): SEEN.
- "Column K is hidden and its font is white" (inventory `hidden_text`, `hidden_columns`): SEEN.

**INFERRED**: the evidence points there but does not prove it. State the evidence and what would confirm it.
- "The typed 1.18 in the landed-cost formulas is probably import duty". The number is SEEN; its meaning is INFERRED. Settle by: procurement confirms the rate and its date.
- "The shipping-zone list was meant to be in distance order". Settle by: logistics owner or the source table.
- "The cleanup macro is run before each month's export". Its source is SEEN; when it runs is INFERRED.

**UNKNOWN**: the file cannot answer it. Say who or what could.
- Why an override exists, which of two contradictory rules is policy, where an external file's numbers came from, the date of a hard-coded rate.

Split one observation into two rows when its parts have different tags: the mechanism is SEEN, the intent is INFERRED or UNKNOWN. Scoring and readers both rely on this split.

## What counts as a problem

`problem = Y` only when, in the workbook **as it stands**:
- an output is wrong, or differs from what its label, note or comment says it is; or
- an output would go wrong for inputs inside the range the model is plainly used for (the values already in it, or the next row someone would add the same way); or
- logic that changes outputs is hidden, overridden or not reproducible from formulas (hidden sheets with overrides, white text, values a macro wrote, links to missing files); or
- a rule people rely on is not actually enforced (it lives only in formatting or a note).

`problem = N` for:
- neutral descriptions of the logic;
- design suggestions that don't make any current output wrong: no input validation, no error guard for impossible inputs (all headcounts zero), a fixed number of rows that are all filled correctly, no recorded source or owner for an assumption. These are useful; put the important ones in the narrative's "What would worry me" if you think they matter;
- anything about the environment or the tools (no engine, sandbox limits). Those go in `RUN/toolkit_issues.md`, not the ledger.

A clean, well-built workbook should come back with zero `Y` rows. Don't manufacture problems to look thorough.

## Rules that are easy to break

- Never write "there are no macros" because a tool printed nothing. Write what the tool said, and tag the gap UNKNOWN if the tool errored.
- "Probably", "likely" and "seems" belong in INFERRED rows, never in SEEN rows.
- A probe that shows an effect makes the effect SEEN. It does not make the cause SEEN unless you also quote the formula.
- If the proof fails on an output, the ledger has a gap: add a row before fixing the model.
- Neutral logic rows (`problem = N`) are welcome. "Group total = sum of the store totals" helps the reader and is SEEN.
- Keep `finding` free of jargon a non-Excel reader would trip on. Put formula text in `evidence`.

## Evidence formats the checker accepts for SEEN

`Sheet!A1` or `'Sheet name'!A1:B2` with formula text, `run 123` (probe_runs.csv), `try 4` (try_log.jsonl), `VBA` line quotes, `inventory` / `graph.json` / `map.md` entries, `defined name`, `externalLink`, `calcPr`, `conditional format`, `data validation`.

## Suggested categories

`lookup_unsorted_approx`, `lookup_duplicate_keys`, `lookup_whitespace`, `iferror_masking`, `nested_if_boundary`, `hardcoded_literal`, `hardcoded_value_in_formula_run`, `formula_pattern_break`, `hidden_logic`, `hidden_sheet_override`, `text_stored_numbers`, `text_dates_locale`, `dynamic_reference`, `vba_logic`, `vba_writes_values`, `comment_vs_formula`, `external_link`, `broken_reference`, `broken_name`, `circular_iterative`, `unit_scale`, `range_truncated`, `metric_definition`, `boolean_logic`, `stale_validation`, `volatile_function`, `autoconverted_identifier`, `rounding`, `format_only_rule`, `currency_mixing`, `logic` (neutral description).
