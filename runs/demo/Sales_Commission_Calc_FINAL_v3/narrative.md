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
