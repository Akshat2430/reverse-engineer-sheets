"""Rebuild of Sales_Commission_Calc_FINAL_v3 (USE THIS).xlsm -- as it is, bugs included.

The workbook pays sales commission on 60 deals (Inputs rows 2-61), works each deal through
the Calc tab column by column, and totals it on Summary by rep, by region and for the company.

This module reproduces Calc and Summary exactly as the formulas are written, including the
defects in the evidence ledger. Nothing here is "fixed": where the workbook understates a
base rate or drops a rep, so does this code, and the ledger row is cited at that line.

run(inputs, today) -> {"Summary!C5": value, ...} for every cell in graph.json["outputs"].
"""
from sheetre import excelfn as X

DEAL_ROWS = range(2, 62)                 # Calc rows 2..61, one per deal
REP_ROWS = range(5, 17)                  # Summary!A5:A16, the 12 reps on the plan
REGION_ROWS = range(19, 23)              # Summary!A19:A22

# Calc!L rows whose formula carries a typed "*1.07" (L2, L23, L27, ...). These are exactly the
# 15 rows whose Inputs!F says EUR, but the uplift is typed into the formula, not driven by the
# currency cell -- so it stays on these row numbers whatever the data says.  Ledger L06, L07.
FX_UPLIFT_ROWS = frozenset({2, 23, 27, 31, 34, 38, 39, 41, 42, 45, 47, 48, 49, 52, 53})

# Calc!S rows holding a number the ApplyAccelerator macro wrote (1.15x), not a formula.
# graph.json lists them as true inputs; we read them straight through.  Ledger L10, L11.
FROZEN_FINAL_ROWS = range(40, 62)


def _month_of(close_date):
    """MONTH(Calc!E_r) on a close date.

    A blank close date reads as serial 0, which Excel calls 00-Jan-1900 and reports as
    month 1 -- so blank-dated deals land in the Q1 column.  excelfn.month maps serial 0 to
    1899-12-30 (month 12) instead, so serials below 1 are handled here.  Confirmed by try 1
    (Inputs!G2 blanked: Calc!O2 took the net commission, Calc!R2 stayed 0).
    """
    serial = X.num(close_date)
    if serial < 1:
        return 1
    return X.month(serial)


def _base_rate(deal_value, table):
    """Calc!H_r =VLOOKUP(F_r,Rates!$A$2:$B$10,2,TRUE), Excel's search, blank keys included.

    excelfn.vlookup reproduces Excel's binary search on unsorted keys, but it ranks a blank
    key as 0, and Excel does not: when the search lands on an empty cell it steps down to the
    nearest non-empty key at or above the current floor and compares that, while still
    narrowing from the midpoint it originally picked. With no non-empty key in reach the
    lookup fails with #N/A. Derived from Excel's own answers to probe runs 2191, 2193, 2197
    and 2201 (the four runs that blank a key in Rates!A2:A10) and recorded in
    toolkit_issues.md; it only ever differs from excelfn.vlookup when a key is blank.
    """
    keys = [row[0] for row in table]
    lo, hi, ans = 0, len(keys) - 1, -1
    while lo <= hi:
        mid = (lo + hi) // 2
        j = mid
        while j >= lo and keys[j] is None:
            j -= 1
        if j < lo:                                  # nothing but blanks down to the floor
            hi = mid - 1
            continue
        if X.eq(keys[j], deal_value):
            while j + 1 < len(keys) and X.eq(keys[j + 1], deal_value):
                j += 1
            return table[j][1]
        if X.lt(keys[j], deal_value):
            ans, lo = j, mid + 1
        else:
            hi = j - 1
    if ans < 0:
        raise X.NA
    return table[ans][1]


def _calc_row(r, c):
    """One row of the Calc tab: Calc!B_r .. Calc!S_r for deal row r.

    Returns a dict of column letter -> value. An Excel error anywhere in the chain is
    stored as an XLError and left to propagate into the Summary aggregates, exactly as the
    workbook does.
    """
    out = {}

    # Calc!B:E are straight pass-throughs of the Inputs tab.  Inputs!F (Currency) is read by
    # no formula anywhere in the workbook -- that is ledger L06.
    out["B"] = c[f"Inputs!B{r}"]                                    # =Inputs!B_r   Sales Rep
    out["C"] = c[f"Inputs!C{r}"]                                    # =Inputs!C_r   Region
    out["D"] = c[f"Inputs!D{r}"]                                    # =Inputs!D_r   Product
    out["E"] = c[f"Inputs!G{r}"]                                    # =Inputs!G_r   Close Date

    # Calc!F_r =N(Inputs!E_r).  N() turns text into 0, so a deal value typed as text
    # ("64,500", "1,25,000", "85000") becomes a zero-value deal.  Ledger L03.
    out["F"] = X.n(c[f"Inputs!E{r}"])
    f = out["F"]

    # Calc!G_r: tier index 1-9 from a nested IF.  The EMEA branch uses >=250000 while the
    # branch for every other region uses >250000, so at exactly 250,000 EMEA gets tier 6 and
    # everyone else drops to tier 5.  Ledger L02.
    region_is_emea = X.eq(out["C"], "EMEA")
    if f >= 1000000:
        out["G"] = 9
    elif f >= 600000:
        out["G"] = 8
    elif f >= 400000:
        out["G"] = 7
    elif region_is_emea and f >= 250000:
        out["G"] = 6
    elif f > 250000:
        out["G"] = 6
    elif f >= 150000:
        out["G"] = 5
    elif f >= 100000:
        out["G"] = 4
    elif f >= 50000:
        out["G"] = 3
    elif f >= 25000:
        out["G"] = 2
    else:
        out["G"] = 1

    # Calc!H_r =VLOOKUP(F_r,Rates!$A$2:$B$10,2,TRUE).  Approximate match, and Rates!A2:A10 is
    # NOT sorted ascending (150000 sits above 100000, 1000000 above 600000), so Excel's binary
    # search lands on the wrong band for many deal sizes.  excelfn.vlookup reproduces Excel's
    # search on unsorted keys rather than the intended bracket.  Ledger L01.
    rate_table = c.table("Rates!A2:B10")
    try:
        out["H"] = _base_rate(f, rate_table)
    except X.XLError as e:
        out["H"] = e

    # Calc!I_r =INDEX(Rates!$E$2:$E$10,G_r): tier kicker, picked by the tier index.
    kicker_col = c.col("Rates!E2:E10")
    try:
        out["I"] = X.index(kicker_col, out["G"])
    except X.XLError as e:
        out["I"] = e

    # Calc!J_r =IFERROR(VLOOKUP(D_r,Rates!$G$2:$H$6,2,FALSE),0): product bonus, 0 if the
    # product is not in the table (the IFERROR hides a failed match).  Ledger L08.
    product_table = c.table("Rates!G2:H6")
    out["J"] = X.iferror(lambda: X.vlookup(out["D"], product_table, 2, approx=False), 0)

    # Calc!K_r =MIN(H_r+I_r+J_r,0.18): effective rate, capped at 18%.  The note on Rates!B1
    # says the cap is 15% per a CFO memo; the formula says 18%.  Ledger L04, L05.
    out["K"] = _arith(lambda: min(X.num(out["H"]) + X.num(out["I"]) + X.num(out["J"]), 0.18))

    # Calc!L_r: gross commission = deal value x effective rate x the rep's plan participation,
    # with a failed rep lookup swallowed to 0 by IFERROR.  A rep name that does not match
    # Rates!J2:J13 exactly therefore earns nothing -- which is what the trailing space on
    # "Rohan Mehta " does.  Ledger L08, L09.  The 15 FX_UPLIFT_ROWS multiply by a typed 1.07.
    rep_table = c.table("Rates!J2:K13")
    participation = X.iferror(lambda: X.vlookup(out["B"], rep_table, 2, approx=False), 0)
    out["L"] = _arith(lambda: _gross(f, out["K"], participation, r))

    # Calc!M_r =IF(Inputs!H_r>90,-0.02*L_r,0): a 2% clawback when the invoice took more than
    # 90 days to pay.  Column M is hidden and its text is white.  Ledger L12.
    days_to_pay = c[f"Inputs!H{r}"]
    late = X.lt(90, days_to_pay)                                    # Excel's  Inputs!H_r > 90
    out["M"] = _arith(lambda: -0.02 * X.num(out["L"]) if late else 0)

    # Calc!N_r =L_r+M_r: net commission.
    out["N"] = _arith(lambda: X.num(out["L"]) + X.num(out["M"]))

    # Calc!O_r..R_r =IF(ROUNDUP(MONTH($E_r)/3,0)=q,$N_r,0): the net commission dropped into its
    # calendar quarter column.  Note these are built from N (net), never from S (final), so the
    # Q4 accelerator never reaches them.  Ledger L11.
    out.update(_quarter_columns(out["E"], out["N"]))

    # Calc!S_r: Final Commission.  Rows 2-39 are "=N_r".  Rows 40-61 hold numbers the
    # ApplyAccelerator macro typed over the formula (net x 1.15); we read them as inputs.
    if r in FROZEN_FINAL_ROWS:
        out["S"] = c[f"Calc!S{r}"]
    else:
        out["S"] = out["N"]
    return out


def _gross(deal_value, rate, participation, r):
    """Calc!L_r, in Excel's left-to-right order: ((F*K)*participation) and then *1.07 if typed."""
    gross = deal_value * X.num(rate) * X.num(participation)
    if r in FX_UPLIFT_ROWS:
        gross = gross * 1.07
    return gross


def _quarter_columns(close_date, net):
    """Calc!O_r..R_r. ROUNDUP(MONTH(E)/3,0) is 1..4; the matching column takes N, the rest 0."""
    try:
        q = int(X.roundup(_month_of(close_date) / 3, 0))
    except X.XLError as e:
        return {letter: e for letter in ("O", "P", "Q", "R")}
    cols = {}
    for i, letter in enumerate(("O", "P", "Q", "R"), start=1):
        cols[letter] = _arith(lambda: X.num(net)) if q == i else 0
    return cols


def _arith(fn):
    """Run an arithmetic step, returning the Excel error instead of raising (errors propagate
    down the row and into the Summary aggregates, as they do in the workbook)."""
    try:
        return fn()
    except X.XLError as e:
        return e
    except ZeroDivisionError:
        return X.DIV0


def _calc_sheet(c):
    """Every Calc row, then the columns the Summary tab aggregates over."""
    rows = {r: _calc_row(r, c) for r in DEAL_ROWS}
    cols = {letter: [rows[r][letter] for r in DEAL_ROWS] for letter in "BCDEFGHIJKLMNOPQRS"}
    return rows, cols


def _rep_commission(c, cols, rep):
    """Summary!C_r =SUMIF(Calc!B,rep,Calc!S)+SUMIF(adj!A2:A20,rep,adj!B2:B20).

    The second SUMIF reads the veryHidden `adj` sheet, which carries two manual
    adjustments (+4,500 and -1,200).  Ledger L13, L14.
    """
    from_deals = _agg(lambda: X.sumif(cols["B"], rep, cols["S"]))
    adjustment = _agg(lambda: X.sumif(c.col("adj!A2:A20"), rep, c.col("adj!B2:B20")))
    return _arith(lambda: X.num(from_deals) + X.num(adjustment))


def _agg(fn):
    try:
        return fn()
    except X.XLError as e:
        return e


def _selected_quarter_column(c, cols):
    """The column OFFSET picks for Summary!D5:D16.

    =OFFSET(Calc!$O$2:$O$61,0,MATCH($B$2,Calc!$O$1:$R$1,0)-1) -- the quarter chosen in the
    Summary!B2 dropdown is matched against the typed headers Calc!O1:R1.  Ledger L15.
    """
    headers = c.row("Calc!O1:R1")
    try:
        i = X.match(c["Summary!B2"], headers, 0)
    except X.XLError as e:
        return e
    return cols["OPQR"[i - 1]]


def run(inputs, today=None):
    c = X.Cells(inputs)
    rows, cols = _calc_sheet(c)
    out = {}

    # --- Summary C5:C16 / D5:D16 / E5:E16: one line per rep on the plan -------------------
    quarter_col = _selected_quarter_column(c, cols)
    for sr in REP_ROWS:
        rep = c[f"Summary!A{sr}"]
        out[f"Summary!C{sr}"] = _rep_commission(c, cols, rep)
        # Summary!D_r: the selected quarter's net commission for this rep.
        if isinstance(quarter_col, X.XLError):
            out[f"Summary!D{sr}"] = quarter_col
        else:
            out[f"Summary!D{sr}"] = _agg(lambda: X.sumif(cols["B"], rep, quarter_col))
        # Summary!E_r =COUNTIF(Calc!B,rep): deal count. Exact text match, so a rep name that
        # does not match the Calc tab exactly counts zero deals.  Ledger L09.
        out[f"Summary!E{sr}"] = _agg(lambda: X.countif(cols["B"], rep))

    # --- Summary C19:C22: the same final commission re-totalled by region ------------------
    # These read Calc!S only: the adj sheet's manual adjustments are NOT in them, so the
    # region block and the company total below disagree by the size of adj.  Ledger L14.
    for sr in REGION_ROWS:
        region = c[f"Summary!A{sr}"]
        out[f"Summary!C{sr}"] = _agg(lambda: X.sumif(cols["C"], region, cols["S"]))

    # --- Summary C24:C26: company totals --------------------------------------------------
    # C24 =SUM(C5:C16): the sum of the rep lines, so it covers only deals whose rep name
    # matches one of the 12, plus the adj amounts.  Ledger L09, L14.
    out["Summary!C24"] = _agg(lambda: X.xl_sum([out[f"Summary!C{sr}"] for sr in REP_ROWS]))
    # C25 =SUM(Inputs!E2:E61), labelled "Total bookings (USD)".  SUM skips the text-typed deal
    # values, and the EUR deals go in unconverted.  Ledger L03, L06, L16.
    out["Summary!C25"] = _agg(lambda: X.xl_sum(c.col("Inputs!E2:E61")))
    # C26 =C24/C25, labelled "Commission as % of bookings".  Ledger L16.
    out["Summary!C26"] = _arith(lambda: X.num(out["Summary!C24"]) / X.num(out["Summary!C25"]))
    return out
