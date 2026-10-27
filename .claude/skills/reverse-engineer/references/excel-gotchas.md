# Excel gotchas: what the flags can mean

A checklist for steps 2 and 3. Each entry says how it shows up in `inventory.json` / `map.md` / `probe_summary.json` and what to check before you write a ledger row.

## Lookups

- **Approximate match on unsorted keys.** `VLOOKUP(..., TRUE)` (or the 4th argument left out), `MATCH(..., 1)` (or left out), and `LOOKUP` all binary-search and assume ascending keys. On unsorted keys they return wrong rows silently, and the result can even go down as the input goes up. Shows up as `map.md` "Lookup tables with problems: NOT sorted ascending" and threshold `reversal`s. Rebuild with `excelfn.vlookup/match`, which reproduce Excel's exact search on unsorted data.
- **Duplicate keys.** Exact-match lookups return the first occurrence. Check which duplicate is newer (effective dates, row order) and whether the stale one is first.
- **Whitespace and case.** A trailing space on either side makes an exact match fail. Matching ignores case. `inventory.whitespace_text` plus probe `trailing-space` variants.
- **Errors swallowed.** `IFERROR(lookup, 0)` or `IFERROR(..., "")` turns a failed lookup into a plausible number. `inventory.error_masking`: list what each wrapper hides and check whether any lookup inside actually fails today.
- **Identifiers changed on paste.** Codes like `DEC1`, `OCT-4` or `3E7` become dates or numbers when pasted. Shows up as `inventory.mixed_type_columns` (a mostly-text ID column containing numbers or dates). Lookups on those keys never match.

## Formulas

- **Inconsistent boundaries.** `>=` in one branch and `>` in another, or different thresholds per segment inside one nested IF. Probe `thresholds`: one row "steps at T", another "steps just above T".
- **Typed constants.** Rates, FX, uplifts or caps typed inside formulas (`*1.18`, `/0.92`, `MIN(x, 0.2)`). `inventory.literals` grouped by column. The number is SEEN; its meaning and date are usually INFERRED or UNKNOWN.
- **Pattern breaks.** One cell in a copied-down column with a different formula. `inventory.formula_pattern_breaks` and the `!` lines in `map.md`.
- **Typed values inside formula runs.** A number pasted over a formula never updates. Often a macro wrote it (look for `.Value =` in `RUN/vba/`). `inventory.constants_in_formula_runs`.
- **Boolean logic slips.** `AND`/`OR` nesting that doesn't match the written rule. Compare the formula with any stated policy text on the sheets and probe the corners.
- **Ranges that stop short.** `SUM(D4:D30)` while data runs to row 36 (rows added later). `map.md` "Aggregates that stop short of adjacent data".
- **Metric definitions.** Labels and notes that say one thing ("average", "%", "per month") while the formula does another (sum, ratio of sums, annual). Read labels and comments next to output formulas.
- **Mixed units.** Adding different currencies, or values in thousands to values in units. `inventory.scale_outliers` (a sheet whose values are about 1,000x smaller than its siblings), `unit_labels`.
- **Rounding.** Displayed to 2 dp, stored with more. One total rounds each line, another doesn't. "Match to the cent" is defined on stored values. Prove reports outputs that match only after rounding.

## Hidden things

- **Hidden and very hidden sheets.** Very hidden sheets don't appear in Excel's Unhide dialog. They often hold overrides or old tables. `inventory.hidden_sheets`. Check whether anything references them (`map.md` marks `[veryHidden sheet]`).
- **Hidden text.** White or near-white font on white, the `;;;` number format, hidden columns and rows. `inventory.hidden_text`, `hidden_columns`, `hidden_rows`.
- **Comments and notes** that state a rule. Compare them with the formula that implements it. A contradiction is SEEN on both sides; which one is policy is UNKNOWN.
- **Rules that live only in formatting.** A red conditional format at a threshold that no formula enforces. `inventory.cf_threshold_not_in_formulas`.
- **Data validation lists** that point at hidden or stale ranges. The dropdown may not offer values the model expects. Excel re-saves cross-sheet lists in an `x14` extension, which the inventory reads directly (openpyxl alone drops them).

## Dynamic and external

- **OFFSET / INDIRECT** pick their targets at run time. The static map lists candidates (`<~ dynamic`). Probe the driver cell to see which outputs move.
- **External links** to other workbooks keep only cached values. If the source file is gone, those inputs are frozen. `inventory.external_links`.
- **Broken names.** A defined name pointing to `#REF!`, often next to a near-duplicate valid name (`Rate` vs `Rate_`). Formulas using the broken one fail, or fall back via IFERROR. `inventory.broken_names`, `similar_names`.
- **`#REF!` inside formulas** left after rows were deleted. Usually wrapped in IFERROR, so the whole term silently becomes 0. `inventory.ref_errors_in_formulas`.

## Calculation

- **Circular references with iterative calculation on.** `calcPr iterate="1"` plus `map.md` "Circular references". Rebuild with `excelfn.iterate` and expect convergence-level differences (well under the tolerance).
- **Volatile functions.** TODAY/NOW change daily and RAND changes on every recalculation. OFFSET/INDIRECT recalculate on every change. Prove pins today's date; flag date-dependent outputs.
- **Manual calculation mode** (`calcPr calcMode="manual"`). The saved numbers may be stale. Probe's `recalculated_differs_from_cached` tells you.
- **Text dates.** `DATEVALUE("05/06/2028")` is 5 June in day-first locales (en-IN, en-GB) and 6 May in en-US. The engine's locale is in `probe_summary.json`; use `excelfn.datevalue(text, locale)`.
- **Text-stored numbers.** `SUM` and `SUMIF` skip text. `+`, `*` and `=A1*1` coerce text that looks numeric. `N()` turns text into 0. Indian grouping (`1,23,456`) and some other formats are **not** coerced on any locale. `inventory.text_numbers`.

## VBA

- Read every module in `RUN/vba/`. Look for writes (`.Value =`, `.Formula =`), deletes, hard-coded sheet names and multipliers.
- A macro that writes values leaves numbers no formula can reproduce. The rebuild treats them as inputs. Ledger: the mechanism is SEEN (quote the lines); when and how often it runs is INFERRED or UNKNOWN.
- A compile error anywhere in a module stops all its macros. A common cause is a variable named like a built-in function (`CDate`, `Val`, `Str`). The code is still readable.

## Also worth a look

Merged cells inside data ranges (break sorting and filling), array formulas and spills, `SUBTOTAL` vs `SUM` with filtered rows, percentages stored as whole numbers (15 vs 0.15), relative references that drift when copied across, the 1900 leap-year bug in serial dates before March 1900, and floating-point equality (`=0.1+0.2=0.3` is FALSE).
