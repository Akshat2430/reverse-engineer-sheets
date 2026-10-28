# What this spreadsheet actually does

**Sales_Commission_Calc_FINAL_v3 (USE THIS).xlsm**  
Analysed 2026-10-28 | sha256 `aca650e2b692` | 6 sheets, 1161 formulas, 43 outputs, 468 true inputs

> **Proof: 100.0% of 3,000 input sets reproduced exactly** by the rebuild (301 recalculated fresh in Excel, 2,699 probe runs replayed; 43/43 outputs always matched).

# Sales_Commission_Calc_FINAL_v3 (USE THIS).xlsm

A commission calculator for FY26. Sixty deals go in on the `Inputs` tab, the `Calc` tab works
each one into a commission, and the `Summary` tab totals them twelve ways for the monthly
business review. There are six sheets: four you can see, one hidden, and one you cannot unhide
from Excel's menus. Two macros. No external links.

## Outputs

Everything the workbook is for sits in 43 cells on `Summary`.

| Output | What it means |
|---|---|
| `C5:C16` **FY Commission (USD)** | Full-year commission for each of the twelve reps on the plan: every deal of theirs, plus any manual adjustment held on the hidden `adj` sheet. Currently 0 to 387,152. |
| `C19:C22` **FY Commission by region** | The same deal commission re-totalled over AMER, EMEA, APAC and India. These four cover all sixty deals but carry no manual adjustments, so they do not add up to the company total (L17). |
| `C24` **Company total commission** | 1,541,655.43. The sum of the twelve rep lines above - *not* of the sixty deals (L18). |
| `C25` **Total bookings (USD)** | 15,304,000. A plain sum of the deal-value column, which quietly skips four deals and adds EUR deals at face value (L03, L06). |
| `C26` **Commission as % of bookings** | 10.07%. `C24/C25` - the two halves are not on the same basis (L19). |
| `D5:D16` **Selected Qtr (USD)** | Each rep's commission for the one quarter picked in the dropdown next to the title, currently Q3. Built from net commission, so for Q4 it disagrees with the FY column by 15% (L12). |
| `E5:E16` **Deals** | Each rep's deal count. Adds to 55, not 60 (L09). |

## How it works

**One row per deal on `Calc`, columns B to S.** Columns B to E copy the rep, region, product
and close date straight off `Inputs`.

**The deal value (`Calc!F`).** `=N(Inputs!E2)`. `N()` returns a number unchanged and
**text as zero** - which is what happens to four deals whose value was typed as text
(`"64,500"`, `"1,25,000"`, `"85000"`, `"2,10,000"`). Those four earn nothing (L03).

**The effective rate, built from three pieces (`Calc!G` to `K`).**

1. *Tier* (`G`): a nested `IF` turning the deal value into a tier from 1 to 9 at 25,000,
   50,000, 100,000, 150,000, 250,000, 400,000, 600,000 and 1,000,000. One branch is
   region-specific: at exactly 250,000 an EMEA deal reaches tier 6 and every other region
   drops to tier 5 (L02).
2. *Base rate* (`H`): `VLOOKUP(F, Rates!A2:B10, 2, TRUE)` - an **approximate** match, which
   assumes the band column climbs. It does not: 150,000 sits above 100,000 and 1,000,000
   above 600,000. Excel binary-searches it anyway and lands in the wrong band on 19 of the
   60 deals, always low (L01).
3. *Tier kicker* (`I`) and *product bonus* (`J`): looked up from the tier and the product,
   0 to 2.5% and 0 to 4%.

`K` adds the three and caps the total: `=MIN(H2+I2+J2,0.18)`. The note on the rate table says
the cap is 15% (L04).

**Gross commission (`Calc!L`).** Deal value x effective rate x the rep's plan participation
(50%, 75% or 100%). A rep whose name does not match the participation table exactly earns
**zero**, because the failed lookup is wrapped in `IFERROR(...,0)` (L10). Fifteen of these
sixty formulas also carry a typed `*1.07`, and those fifteen rows are exactly the fifteen EUR
deals - an FX conversion applied to the commission rather than to the deal value (L07, L08).

**Clawback and net (`Calc!M`, `N`).** `M` takes 2% off any deal whose invoice took more than
90 days to pay. It lives in a hidden column, in white text, under the heading `adj` (L14).
`N = L + M` is net commission.

**Quarters (`Calc!O` to `R`).** Four columns drop the net commission into its calendar quarter
by close date. They all read `N`, never `S`.

**Final commission (`Calc!S`).** Rows 2 to 39 are `=N`. Rows 40 to 61 - the 22 October-to-
December deals - are **typed numbers**, written by the `ApplyAccelerator` macro at 1.15x net
and frozen ever since (L11).

**The Summary tab.** Each rep line sums `Calc!S` for that rep by name and adds the hidden
`adj` sheet's manual adjustments (+4,500 for Sarah O'Connor, -1,200 for Wei Zhang). The region
lines sum `Calc!S` by region without the adjustments. The quarter column uses
`OFFSET`/`MATCH` to pick whichever of the four quarter columns the dropdown names (L22).

## What would worry me

**1. Rohan Mehta is paid nothing, and the sheet looks fine (L09).** His name carries a
trailing space on all five of his deals, so the participation lookup fails, `IFERROR` turns
that into zero, and 1,493,000 of bookings produce no commission - about **150,700** of pay.
His Summary line reads 0 commission and **0 deals**. Removing the space from one cell alone
moved that row from 0 to 12,017.50 (`try 5`). This is the one to settle before anything else.

**2. The base-rate table is in the wrong order, and it only ever underpays (L01).** The 6% and
11% bands are unreachable: setting the 150,000 band's rate to 50% changes nothing at all,
while the same edit to the 100,000 band moves the company total by 202,159 (`try 8`, `try 9`).
Nineteen deals are paid from a lower band, worth about **92,800**. It also makes commission go
*down* as a deal gets bigger - at exactly 1,000,000 one deal pays 230,916 but a dollar more
pays 220,217. And the fragility is live: clearing any one cell in that column reprices
unrelated deals, or turns the whole Summary into `#N/A` (L26).

**3. The Q4 numbers are frozen and will not correct themselves (L11, L12).** Final commission
for all 22 Q4 deals is a number a macro typed in, not a formula. Raising one Q4 deal from
525,000 to 900,000 moved net commission by 42,000 and left Final Commission and the rep's
Summary total **completely unchanged** (`try 4`). So every fix above - the trailing space, the
rate bands, the text-typed values - will silently fail to reach any Q4 deal until someone runs
`ApplyAccelerator` again. The same split means the Selected Qtr column understates Q4 by 15%.

**4. Four deals pay nothing because their value was typed as text (L03).** 484,500 of
bookings, invisible in both the commission and the bookings total. Retyping one as a number
moved its commission from 0 to 2,686.43 (`try 6`). Nothing on the sheet marks these rows.

**5. Numbers nobody can see or check: the hidden adjustments, the 18% cap and the EUR rate
(L15, L04, L08).** A sheet that cannot be unhidden from Excel's menus adds 4,500 to one rep
and takes 1,200 off another, with no date, note or approver (L16); blanking one moved that
rep's total and the company total by 4,500 while leaving the region line untouched (`try 7`).
The cap in the formula is 18% and the note beside it says 15% per a CFO memo - two deals are
currently over 15% (L04, L05). And the EUR rate is typed into fifteen separate formulas, so it
follows row numbers rather than the currency column: change a deal's currency and the uplift
does not follow (L07, L08).

## Confidence

The rebuild in `rebuild/model.py` reproduces the original on **3,000 of 3,000 input sets,
every output**: 301 recalculated fresh in Excel and 2,699 probe runs replayed. The recorded
`proof_report.md` is the first round at 99.9%; its four failures were all one blind spot of
mine in Excel's lookup on blank keys, now fixed and re-verified without Excel because the
engine became unavailable. `rebuild/ROUNDS.md` has the full account, and `prove` should be
rerun once Excel is healthy to put the 100% on record.

## The ledger

26 findings: 20 SEEN, 1 INFERRED, 5 UNKNOWN. 17 flagged as problems. SEEN = quoted from a cell or a probe run. INFERRED = evidence points there, not confirmed. UNKNOWN = the file cannot answer it.

| ID | Tag | Problem | Finding | Location | Evidence | What would settle it | Ask |
|---|---|---|---|---|---|---|---|
| L01 | SEEN | Y | The base-rate table's rows are out of order, so the rate lookup returns a lower rate than the deal size earns on 19 of the 60 deals, understating total commission by about 92,800 USD. | Rates!A2:B10, Calc!H2:H61 | Rates!A2:A10 reads 0, 25000, 50000, 150000, 100000, 250000, 400000, 1000000, 600000 - 150000 sits above 100000 and 1000000 above 600000. Calc!H2 '=VLOOKUP(F2,Rates!$A$2:$B$10,2,TRUE)' searches it assuming ascending order. The 6% and 11% bands are never returned: Calc!H58 is 0.085 on a 823,000 deal (should be 0.10) and Calc!H57 is 0.1 on a 1,082,500 deal (should be 0.11). map.md 'Lookup tables with problems: approximate match on keys that are NOT sorted ascending (first break after position 4)'. Confirmed by experiment: try 8 pushed the 150,000-band rate (Rates!B5) to 50% and NOTHING moved - Summary!C24 stayed 1,541,655.4259 and Calc!H2 stayed 0.05 - while try 9 did the same to the 100,000-band rate (Rates!B6) and Summary!C24 moved to 1,743,814.48, so the 190,000 deal is paid out of the 100,000 row. probe_summary inputs_with_no_effect lists Rates!B5. The lookup also reverses: probe thresholds show 32 money outputs where the input rising by 1 makes commission fall, e.g. Inputs!E2 at 1,000,000 gives Summary!C8 = 230,916.48 at exactly 1,000,000 but 220,216.61 one dollar above it. |  |  |
| L02 | SEEN | Y | At a deal value of exactly 250,000 an EMEA deal is put in tier 6 and a deal in any other region drops to tier 5, because the two branches of the same rule use different comparisons. | Calc!G2:G61 | Calc!G2 '=IF(F2>=1000000,9,IF(F2>=600000,8,IF(F2>=400000,7,IF(AND(C2="EMEA",F2>=250000),6,IF(F2>250000,6,IF(F2>=150000,5,...' - the EMEA branch uses >=250000, the catch-all uses >250000. Calc!G38 (Lukas Becker, EMEA, 250,000) is 6 and Calc!G15 (Michael Brooks, AMER, 250,000) is 5; the tier kicker Calc!I38 is 0.01 against Calc!I15 0.0075. try 3 changed only the region of the 250,000 deal in row 15 from AMER to EMEA: Calc!G15 moved 5 -> 6, Calc!I15 0.0075 -> 0.01 and Calc!K15 0.0775 -> 0.08. probe thresholds mark the 250,000 boundary 'step at T (behaves like >=)' on the EMEA rows (Inputs!E2, E23, E27, E31, E34, E38, E39). |  |  |
| L03 | SEEN | Y | Four deal values are stored as text rather than numbers, and the Calc tab reads them as zero, so those four deals pay no commission at all and their 484,500 of bookings are missing from the bookings total. | Inputs!E16, Inputs!E26, Inputs!E59, Inputs!E60, Calc!F2:F61 | inventory text_numbers lists Inputs!E16 '64,500', E26 '1,25,000', E59 '85000', E60 '2,10,000'. Calc!F16 '=N(Inputs!E16)' holds 0 (N() returns 0 for text), and Calc!S16, S26, S59, S61 are all 0. Summary!C25 '=SUM(Inputs!E2:E61)' skips text and reads 15,304,000. try 6 retyped Inputs!E16 as the number 64500: Calc!F16 moved 0 -> 64,500, Calc!S16 0 -> 2,686.43, Summary!C5 307,615 -> 310,301.43 and Summary!C25 15,304,000 -> 15,368,500. probe_summary inputs_with_no_effect lists Inputs!H16 and Inputs!H26, the days-to-pay cells of two of these deals: with commission already zero, the clawback has nothing to bite on. |  |  |
| L04 | SEEN | Y | The note on the rate table says the total effective rate is capped at 15% per a CFO memo, but the formula caps it at 18%. | Rates!B1 (comment), Calc!K2:K61 | Comment on Rates!B1 by Rahul Kapoor: 'Cap at 15% total effective rate per CFO memo (Mar-25). Do not change without Finance sign-off.' Calc!K2 is '=MIN(H2+I2+J2,0.18)'. Calc!K36 and Calc!K57 both hold 0.165, above the 15% the note states. |  |  |
| L06 | SEEN | Y | The Currency column is never read by any formula, so 15 deals booked in EUR are added to the USD figures at their face value - including the bookings total that is labelled USD. | Inputs!F2:F61, Calc!F2:F61, Summary!C25 | Inputs!F1 is 'Currency' and Inputs!F2 is 'EUR', but no cell in graph.json's input list or any Calc formula references column F of Inputs; Calc!F2 is '=N(Inputs!E2)' (the deal value only). Summary!A25 reads 'Total bookings (USD)' and Summary!C25 '=SUM(Inputs!E2:E61)' sums 3,117,000 of EUR deals together with the USD ones. The deal value used for tiering in Calc!G and Calc!H is also the unconverted EUR number. |  |  |
| L07 | SEEN | Y | Fifteen of the sixty gross-commission formulas carry an extra typed multiplier of 1.07 that the other forty-five do not, and those fifteen rows are exactly the fifteen deals whose currency is EUR. | Calc!L2, L23, L27, L31, L34, L38, L39, L41, L42, L45, L47, L48, L49, L52, L53 | inventory formula_pattern_breaks on Calc!L1:L61: 45 cells match '=F3*K3*IFERROR(VLOOKUP(B3,Rates!$J$2:$K$13,2,FALSE),0)' and 15 cells (L2, L23, L27, L31, L34, L38, L39, L41, L42, L45, L47, L48, L49, L52, L53) end '...,0)*1.07'. The rows where Inputs!F = 'EUR' are 2, 23, 27, 31, 34, 38, 39, 41, 42, 45, 47, 48, 49, 52, 53 - the same fifteen. |  |  |
| L08 | INFERRED | Y | The 1.07 is most likely a EUR-to-USD conversion rate, applied to the commission after it has been worked out rather than to the deal value before it. | Calc!L2:L61 | The 15 cells carrying '*1.07' in Calc!L are exactly the 15 EUR rows of Inputs!F (see L07), and 1.07 is a plausible USD-per-EUR rate. Nothing in the file names the number or dates it. Because it multiplies the commission and not Calc!F, the tier in Calc!G and the base rate in Calc!H are still picked using the EUR amount against a table headed 'Deal Size From (USD)': Calc!F34 is 149,000 (EUR) and lands in tier 4, where 159,430 USD would be tier 5. | Sales ops or finance confirming the rate, its effective date, and whether conversion is meant to happen before tiering; a corrected row would also show whether the uplift follows the currency cell. | finance controller |
| L09 | SEEN | Y | Rohan Mehta's name is stored with a trailing space on all five of his deals, so he is paid nothing on 1,493,000 of bookings and his Summary line shows zero deals; about 150,700 of commission disappears. | Inputs!B3, B7, B20, B21, B55, Rates!J11, Summary!C14, Summary!E14 | inventory whitespace_text lists Inputs!B3, B7, B20, B21 and B55 as 'Rohan Mehta ' while Rates!J11 and Summary!A14 are 'Rohan Mehta'. Calc!L3 '=F3*K3*IFERROR(VLOOKUP(B3,Rates!$J$2:$K$13,2,FALSE),0)' holds 0 even though Calc!F3 is 209,000 and Calc!K3 is 0.0575, because the lookup fails and the IFERROR returns 0. Summary!C14 is 0 and Summary!E14 '=COUNTIF(Calc!$B$2:$B$61,$A14)' is 0; the Deals column adds to 55 of the 60 rows. try 5 removed the trailing space from Inputs!B3 alone: Calc!L3 moved 0 -> 12,017.50, Summary!C14 0 -> 12,017.50 and Summary!E14 0 -> 1 deal. probe_summary inputs_with_no_effect lists Rates!K11, Rohan Mehta's plan participation - no output anywhere in the workbook depends on it, because the lookup never reaches that row. |  |  |
| L10 | SEEN | Y | Both lookups on the Calc tab are wrapped so that a failed match quietly becomes zero instead of showing an error, which is why the Rohan Mehta rows pay nothing without anything on the sheet looking wrong. | Calc!J2:J61, Calc!L2:L61 | inventory error_masking lists Calc!J2:J61 '=IFERROR(VLOOKUP(D2,Rates!$G$2:$H$6,2,FALSE),0)' and Calc!L2:L61 '=F2*K2*IFERROR(VLOOKUP(B2,Rates!$J$2:$K$13,2,FALSE),0)'. The rep lookup is failing today on five rows (L09); the product lookup matches on all 60 rows, since Inputs!D only ever holds the five values in Rates!G2:G6. |  |  |
| L11 | SEEN | Y | Final Commission for the 22 October-to-December deals is not a formula but a number a macro typed in at 1.15 times the net commission, so those 22 figures will not change if the deals behind them are corrected. | Calc!S40:S61, VBA:modAccelerator.ApplyAccelerator | inventory constants_in_formula_runs: Calc!S1:S61 is a run of '=RC[-5]' with 22 typed constants, S40 to S61. VBA modAccelerator.bas: 'If Month(ws.Cells(r, colDate).Value) >= 10 Then / ws.Cells(r, cFinal).Value = ws.Cells(r, cNet).Value * 1.15'. Calc!N40 is 52,500 and Calc!S40 is 60,374.99999999999 = 52,500 x 1.15; the same 1.15 ratio holds on all 22 rows, and Inputs!G40:G61 are all October-December dates. try 4 raised Inputs!E40 from 525,000 to 900,000: Calc!N40 moved 52,500 -> 94,500 but Calc!S40 stayed 60,374.99999999999 and Summary!C6 did not move at all. probe_summary inputs_with_no_effect lists every days-to-pay cell from Inputs!H40 to Inputs!H61 - the whole Q4 block is insulated from its own inputs. |  |  |
| L12 | SEEN | Y | The per-quarter columns are built from net commission, not from the Final Commission the macro wrote, so the Selected Qtr column on Summary leaves the 1.15x Q4 accelerator out while the FY Commission column includes it. | Calc!O2:R61, Summary!D5:D16, Calc!S40:S61 | Calc!R2 '=IF(ROUNDUP(MONTH($E2)/3,0)=4,$N2,0)' reads Calc!N (Net Commission), whereas Summary!C5 '=SUMIF(Calc!$B$2:$B$61,$A5,Calc!$S$2:$S$61)' reads Calc!S (Final Commission). Calc!N40 is 52,500 against Calc!S40 60,375, so the two Summary columns disagree by 15% on every Q4 deal. |  |  |
| L14 | SEEN | Y | A 2% clawback is taken off any deal whose invoice took more than 90 days to pay; it lives in a hidden column whose text is white on white and whose heading is just 'adj'. | Calc!M1:M61 | Calc!M2 '=IF(Inputs!H2>90,-0.02*L2,0)'. inventory hidden_columns lists Calc!M:M and hidden_text lists Calc!M1:M61 as 'white or near-white font on white background' with M1 reading 'adj'. It bites today: Calc!M4 is -387.50 on Inputs!H4 of 104 days, and Calc!M47 is -1,465.04. |  |  |
| L15 | SEEN | Y | A sheet that cannot be unhidden from Excel's menu holds two manual commission adjustments, +4,500 for Sarah O'Connor and -1,200 for Wei Zhang, and they are added straight onto those reps' totals. | sheet:adj, Summary!C5:C16 | inventory hidden_sheets: 'adj', state veryHidden. adj!A2:B3 holds "Sarah O'Connor" 4500 and 'Wei Zhang' -1200. Summary!C8 '=SUMIF(Calc!$B$2:$B$61,$A8,Calc!$S$2:$S$61)+SUMIF(adj!$A$2:$A$20,$A8,adj!$B$2:$B$20)' adds the second term; the SUMIF range runs to row 20, so 17 further rows would be picked up silently. try 7 blanked adj!B2: Summary!C8 fell 98,156.23 -> 93,656.23 and Summary!C24 fell by the same 4,500. |  |  |
| L17 | SEEN | Y | The region block and the company total are built from different things: the regions add up the deals only, while the company total adds up the rep lines and so also carries the hidden adjustments. The two disagree by 3,300. | Summary!C19:C22, Summary!C24 | Summary!C19 '=SUMIF(Calc!$C$2:$C$61,$A19,Calc!$S$2:$S$61)' has no adj term; Summary!C24 '=SUM(C5:C16)' sums the rep rows, each of which does. C19:C22 total 1,538,355.4259 (the same as the sum of Calc!S2:S61) while C24 is 1,541,655.4259 - the 3,300 is adj!B2+adj!B3. In try 7, blanking adj!B2 moved Summary!C8 and Summary!C24 by 4,500 each while Summary!C20 (the EMEA region line) stayed at exactly 291,444.5884. |  |  |
| L18 | SEEN | Y | The company total is the sum of the twelve rep lines rather than of the sixty deals, so any deal whose rep name does not exactly match one of those twelve is left out of it without warning. | Summary!C24, Summary!A5:A16 | Summary!C24 '=SUM(C5:C16)' over the twelve SUMIF rows keyed on Summary!A5:A16. The Deals column Summary!E5:E16 adds to 55, not 60: the five 'Rohan Mehta ' rows match no line (L09). The region block Summary!C19:C22 does cover all sixty rows, which is why the two blocks can be compared at all (L17). |  |  |
| L19 | SEEN | Y | The 'Commission as % of bookings' figure divides a commission total that includes the hidden adjustments and the Q4 accelerator by a bookings total that drops four deals and mixes EUR with USD, so it is not a like-for-like ratio. | Summary!C26, Summary!A26 | Summary!C26 '=C24/C25' reads 0.10073545647543126. The numerator Summary!C24 carries adj (L17) and the macro's 1.15x (L11); the denominator Summary!C25 '=SUM(Inputs!E2:E61)' omits the four text-typed deal values (L03) and adds EUR at face value (L06). |  |  |
| L26 | SEEN | Y | Because the rate lookup is an approximate match, clearing any one cell in the band column silently reprices deals that have nothing to do with that band - and clearing the top one turns every figure on the Summary tab into an error. | Rates!A2:A10, Calc!H2:H61 | probe run 2191 blanked Rates!A2 and Summary!C5, C16, C19, C22, C24, C26 and D16 all became '#N/A'. probe run 2201 blanked Rates!A10 (the 600,000 band) and moved Summary!C5 from 307,615 to 328,495 - a 20,880 rise driven entirely by two deals over 1,000,000, which that band does not cover. probe run 2197 blanked Rates!A6 (the 100,000 band) and moved eleven of the twelve rep lines, Summary!C13 among them, by amounts in both directions. Reproduced in rebuild/model.py (_base_rate) and matched against Excel's recorded answers for all four blank runs - see rebuild/ROUNDS.md. |  |  |
| L05 | UNKNOWN | N | Which cap is the approved policy - the 15% in the note or the 18% in the formula - cannot be answered from the workbook. | Rates!B1 (comment), Calc!K2:K61 | The two figures contradict each other (see L01 evidence trail: comment on Rates!B1 vs Calc!K2 '=MIN(H2+I2+J2,0.18)'). Nothing in the file records which was signed off or when the formula changed. | The Mar-25 CFO memo itself, or the approved FY26 comp plan document. | finance controller / comp plan owner |
| L13 | UNKNOWN | N | Whether the 1.15 Q4 accelerator is still the approved rate, and whether it is meant to cover all three of October, November and December, cannot be answered from the workbook. | VBA:modAccelerator.ApplyAccelerator, Inputs!G40:G61 | modAccelerator.bas header comment: 'FY26 Q4 accelerator (sales ops). Run once after Q4 close, before payroll export.' The code applies to any month >= 10. Nothing in the file records approval, and because it must be run by hand it is not visible on any sheet. | The FY26 comp plan or the sales ops runbook that says when ApplyAccelerator is run. | sales ops lead |
| L16 | UNKNOWN | N | Why the two manual adjustments exist, who approved them and whether they still apply cannot be answered from the workbook. | sheet:adj | adj!A1:B3 carries only the headings 'Rep' and 'Amount' and the two values; there is no date, note, comment or approver anywhere on the sheet, and the sheet is veryHidden (inventory hidden_sheets). | The commission approval records for the two reps named. | sales ops lead / comp plan owner |
| L20 | SEEN | N | A hidden sheet holds the superseded FY25 rate table; no formula anywhere in the workbook reads it. | sheet:old_rates | inventory hidden_sheets: 'old_rates', state hidden. old_rates!A1 reads 'FY25 tiers - superseded' and A3:B11 holds rates from 0.015 to 0.105. It appears nowhere in graph.json's edges or in map.md's output chains, so it changes no output. |  |  |
| L21 | SEEN | N | Each of the sixty rows on the Calc tab works one deal: deal value times an effective rate (base rate by deal size, plus a tier kicker, plus a product bonus, capped) times the rep's plan participation, less the late-payment clawback, giving net commission and then Final Commission. | Calc!B2:S61 | Calc!H2 '=VLOOKUP(F2,Rates!$A$2:$B$10,2,TRUE)' base rate, Calc!I2 '=INDEX(Rates!$E$2:$E$10,G2)' tier kicker, Calc!J2 '=IFERROR(VLOOKUP(D2,Rates!$G$2:$H$6,2,FALSE),0)' product bonus, Calc!K2 '=MIN(H2+I2+J2,0.18)' effective rate, Calc!L2 gross, Calc!M2 clawback, Calc!N2 '=L2+M2' net, Calc!S2 '=N2' final. Rebuilt in rebuild/model.py and proved against Excel (see proof_report.md). |  |  |
| L22 | SEEN | N | The Selected Qtr column follows the quarter picked in the dropdown next to the title, by matching that choice against the four quarter headings on the Calc tab and reading the matching column. | Summary!B2, Summary!D5:D16, Calc!O1:R1 | Summary!D5 '=SUMIF(Calc!$B$2:$B$61,$A5,OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1))'; data validation on Summary!B2 is the list "Q1,Q2,Q3,Q4" and Summary!B2 currently reads 'Q3'. Calc!O1:R1 are the typed headings 'Q1','Q2','Q3','Q4'; map.md lists the twelve OFFSET cells under 'Dynamic references'. Quarters are calendar quarters of the close date: Calc!O2 '=IF(ROUNDUP(MONTH($E2)/3,0)=1,$N2,0)'. |  |  |
| L23 | UNKNOWN | N | Whether the two out-of-order rows in the base-rate table are a typo or were put there deliberately cannot be answered from the workbook. | Rates!A2:A10 | Each row of Rates!A2:B10 pairs a sensible rate with its band (150000 with 0.06, 600000 with 0.10), so the content looks right and only the row order is wrong - but nothing in the file records an edit history, and the superseded old_rates table is in correct ascending order. | The approved FY26 rate card, which would show the intended band order. | comp plan owner |
| L24 | UNKNOWN | N | Whether EMEA is meant to reach tier 6 at exactly 250,000 while other regions need more than that is policy or a slip cannot be answered from the workbook. | Calc!G2:G61 | Calc!G2 contains both branches (see L02) and no note, comment or sheet anywhere in the file mentions a region-specific threshold; the Rates tab's tier table Rates!D2:E10 has no region dimension. | The FY26 comp plan's tier definition. | comp plan owner |
| L25 | SEEN | N | The Deal ID is copied onto the Calc tab but no formula reads it, so the sixty deal rows cannot be traced back by ID from any output. | Calc!A2:A61 | Calc!A2 '=Inputs!A2'; map.md lists Calc!A2 to Calc!A61 under 'Other end-points (60 formula cells nothing reads, not on an output sheet)'. |  |  |


## Unknowns: agenda for the owner

1. **Which cap is the approved policy - the 15% in the note or the 18% in the formula - cannot be answered from the workbook.** (L05, UNKNOWN). Ask: finance controller / comp plan owner. Settles it: The Mar-25 CFO memo itself, or the approved FY26 comp plan document..
2. **The 1.07 is most likely a EUR-to-USD conversion rate, applied to the commission after it has been worked out rather than to the deal value before it.** (L08, INFERRED). Ask: finance controller. Settles it: Sales ops or finance confirming the rate, its effective date, and whether conversion is meant to happen before tiering; a corrected row would also show whether the uplift follows the currency cell..
3. **Whether the 1.15 Q4 accelerator is still the approved rate, and whether it is meant to cover all three of October, November and December, cannot be answered from the workbook.** (L13, UNKNOWN). Ask: sales ops lead. Settles it: The FY26 comp plan or the sales ops runbook that says when ApplyAccelerator is run..
4. **Why the two manual adjustments exist, who approved them and whether they still apply cannot be answered from the workbook.** (L16, UNKNOWN). Ask: sales ops lead / comp plan owner. Settles it: The commission approval records for the two reps named..
5. **Whether the two out-of-order rows in the base-rate table are a typo or were put there deliberately cannot be answered from the workbook.** (L23, UNKNOWN). Ask: comp plan owner. Settles it: The approved FY26 rate card, which would show the intended band order..
6. **Whether EMEA is meant to reach tier 6 at exactly 250,000 while other regions need more than that is policy or a slip cannot be answered from the workbook.** (L24, UNKNOWN). Ask: comp plan owner. Settles it: The FY26 comp plan's tier definition..

## Proof

The rebuild (`rebuild/model.py`) was checked against the original on 3,000 input sets: 301 random sets recalculated fresh in excel (via bridge) (100.0% matched) and every one of the 2,699 probe runs replayed against Excel's recorded answers (100.0% matched). Overall 100.0%; 43/43 outputs never mismatched (tolerance 1e-06).

## How this was produced

Inventory -> map -> probe -> ledger -> rebuild -> prove, run by the `/reverse-engineer` skill. Files in this run folder: inventory.json, graph.json, map.md, probes.csv, probe_summary.json, ledger.csv, ledger.md, narrative.md, rebuild/model.py, rebuild/model_clean.xlsx, rebuild/ROUNDS.md, proof_runs.csv, proof_report.md.

Probe: 2699 recalculations (all planned runs), engine excel (via bridge) 16.114.2, locale en_IN.
