"""Write rebuild/model_clean.xlsx: the same logic as the original, laid out so it can be read.

Same numbers, same bugs -- this is a readable restatement of the workbook, not a fixed version.
What changes is the layout: every rate, cap and multiplier that the original types inside a
formula is pulled out onto one Assumptions sheet, each step of the calculation gets its own
labelled column, and a Notes sheet lists what the evidence ledger found.

Run from the repo root:  .venv/bin/python <this file>
"""
import csv
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))

RUN = Path(__file__).resolve().parent.parent
SRC = RUN / "work" / "Sales_Commission_Calc_FINAL_v3__USE_THIS_.xlsm"
OUT = Path(__file__).resolve().parent / "model_clean.xlsx"

HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="31506E")
TITLE = Font(bold=True, size=13)
NOTE = Font(italic=True, color="666666")
MONEY = '#,##0.00'
PCT = '0.00%'

# Calc!L rows carrying the typed *1.07 (ledger L07, L08) and Calc!S rows the macro froze (L11).
FX_UPLIFT_ROWS = frozenset({2, 23, 27, 31, 34, 38, 39, 41, 42, 45, 47, 48, 49, 52, 53})
FROZEN_FINAL_ROWS = range(40, 62)


def _header(ws, row, labels, widths=None):
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row, i, text)
        c.font, c.fill = HEAD, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = ws.cell(row + 1, 1)


def assumptions(wb, src):
    """One sheet for every number the original types inside a formula or keeps on the Rates tab."""
    ws = wb.create_sheet("Assumptions")
    rates, adj = src["Rates"], src["adj"]
    ws["A1"] = "Assumptions - every rate, cap and multiplier the model uses"
    ws["A1"].font = TITLE
    ws["A2"] = ("Named here so they can be checked in one place. In the original, the ones marked "
                "'typed into the formula' appear only inside Calc formulas.")
    ws["A2"].font = NOTE

    r = 4
    _header(ws, r, ["Base rate by deal size", "Band from (USD)", "Base rate", "Note"],
            [34, 18, 12, 62])
    r += 1
    band_rows = [(rates.cell(i, 1).value, rates.cell(i, 2).value) for i in range(2, 11)]
    for frm, rate in band_rows:
        ws.cell(r, 2, frm)
        ws.cell(r, 3, rate).number_format = PCT
        r += 1
    ws.cell(r, 1, "Rates!A2:B10, kept in the original's row order")
    ws.cell(r, 4, "Ledger L01: rows 150,000 and 1,000,000 sit out of ascending order, which is "
                  "what breaks the approximate lookup. Order preserved on purpose.").font = NOTE
    r += 3

    _header(ws, r, ["Tier kicker by tier", "Tier", "Kicker", "Note"], [34, 18, 12, 62])
    r += 1
    for i in range(2, 11):
        ws.cell(r, 2, rates.cell(i, 4).value)
        ws.cell(r, 3, rates.cell(i, 5).value).number_format = PCT
        r += 1
    ws.cell(r, 1, "Rates!D2:E10")
    r += 3

    _header(ws, r, ["Product bonus", "Product", "Bonus", "Note"], [34, 22, 12, 62])
    r += 1
    for i in range(2, 7):
        ws.cell(r, 2, rates.cell(i, 7).value)
        ws.cell(r, 3, rates.cell(i, 8).value).number_format = PCT
        r += 1
    ws.cell(r, 1, "Rates!G2:H6")
    r += 3

    _header(ws, r, ["Plan participation", "Sales rep", "Participation", "Note"], [34, 22, 14, 62])
    r += 1
    for i in range(2, 14):
        ws.cell(r, 2, rates.cell(i, 10).value)
        ws.cell(r, 3, rates.cell(i, 11).value).number_format = PCT
        r += 1
    ws.cell(r, 1, "Rates!J2:K13")
    ws.cell(r, 4, "Ledger L09: matched exactly, so 'Rohan Mehta ' on the Inputs tab never "
                  "finds 'Rohan Mehta' here.").font = NOTE
    r += 3

    _header(ws, r, ["Single values", "Value", "Where it lives in the original", "Note"],
            [34, 14, 46, 62])
    r += 1
    singles = [
        ("Effective rate cap", 0.18, "typed into the formula - Calc!K2 MIN(...,0.18)",
         "Ledger L04: the note on Rates!B1 says 15%."),
        ("Late-payment clawback", -0.02, "typed into the formula - Calc!M2",
         "Ledger L14: hidden column, white text, headed 'adj'."),
        ("Days-to-pay threshold", 90, "typed into the formula - Calc!M2", ""),
        ("EUR uplift on commission", 1.07, "typed into 15 of the 60 Calc!L formulas",
         "Ledger L07, L08: applied to the commission, not to the deal value."),
        ("Q4 accelerator", 1.15, "VBA modAccelerator.ApplyAccelerator, typed into Calc!S40:S61",
         "Ledger L11: a number, not a formula, so it does not recalculate."),
        ("EMEA tier-6 threshold", 250000, "typed into the formula - Calc!G2, AND(C2=\"EMEA\",...)",
         "Ledger L02: >= here, > for every other region."),
    ]
    for label, val, where, note in singles:
        ws.cell(r, 1, label)
        v = ws.cell(r, 2, val)
        v.number_format = PCT if isinstance(val, float) and abs(val) < 1.5 else MONEY
        ws.cell(r, 3, where)
        ws.cell(r, 4, note).font = NOTE
        r += 1
    r += 2

    _header(ws, r, ["Manual adjustments (the veryHidden 'adj' sheet)", "Rep", "Amount", "Note"],
            [44, 22, 14, 62])
    r += 1
    for i in range(2, 4):
        ws.cell(r, 2, adj.cell(i, 1).value)
        ws.cell(r, 3, adj.cell(i, 2).value).number_format = MONEY
        r += 1
    ws.cell(r, 4, "Ledger L15, L17: added to the rep lines and the company total, "
                  "but not to the region lines.").font = NOTE
    return ws


def deals(wb, src, model_rows):
    """One row per deal, one labelled column per step, in the order the calculation happens."""
    ws = wb.create_sheet("Deals")
    ws["A1"] = "Deals - one row per deal, one column per step"
    ws["A1"].font = TITLE
    ws["A2"] = ("Steps 1-9 are the Calc tab, renamed. The two 'as typed' columns are where the "
                "original stops being a formula: see the Notes sheet.")
    ws["A2"].font = NOTE
    cols = ["Deal ID", "Sales rep", "Region", "Product", "Close date", "Currency (not used)",
            "1. Deal value used", "2. Tier", "3. Base rate", "4. Tier kicker", "5. Product bonus",
            "6. Effective rate (capped)", "7. Plan participation", "8. EUR uplift applied",
            "9. Gross commission", "10. Late-payment clawback", "11. Net commission",
            "12. Final commission (as typed)", "Calendar quarter"]
    _header(ws, 4, cols, [12, 18, 10, 18, 12, 16, 15, 7, 11, 11, 12, 14, 13, 12, 15, 15, 15, 17, 11])
    inp = src["Inputs"]
    out = 5
    for r, row in model_rows.items():
        ws.cell(out, 1, inp.cell(r, 1).value)
        ws.cell(out, 2, row["B"])
        ws.cell(out, 3, row["C"])
        ws.cell(out, 4, row["D"])
        d = ws.cell(out, 5, row["E"]); d.number_format = "yyyy-mm-dd"
        ws.cell(out, 6, inp.cell(r, 6).value)
        ws.cell(out, 7, row["F"]).number_format = MONEY
        ws.cell(out, 8, row["G"])
        ws.cell(out, 9, row["H"]).number_format = PCT
        ws.cell(out, 10, row["I"]).number_format = PCT
        ws.cell(out, 11, row["J"]).number_format = PCT
        ws.cell(out, 12, row["K"]).number_format = PCT
        ws.cell(out, 13, row["participation"]).number_format = PCT
        ws.cell(out, 14, 1.07 if r in FX_UPLIFT_ROWS else 1)
        ws.cell(out, 15, row["L"]).number_format = MONEY
        ws.cell(out, 16, row["M"]).number_format = MONEY
        ws.cell(out, 17, row["N"]).number_format = MONEY
        s = ws.cell(out, 18, row["S"]); s.number_format = MONEY
        if r in FROZEN_FINAL_ROWS:
            s.font = Font(color="B00020")
        ws.cell(out, 19, row["quarter"])
        out += 1
    ws.cell(out + 1, 1, "Final commission in red is a number the accelerator macro typed over the "
                        "formula (ledger L11); the rest is net commission unchanged.").font = NOTE
    return ws


def summary(wb, outputs, src):
    ws = wb.create_sheet("Summary", 0)
    ws["A1"] = "FY26 Sales Commission Summary - rebuilt"
    ws["A1"].font = TITLE
    ws["A2"] = (f"Same numbers as the original workbook. Selected quarter: "
                f"{src['Summary']['B2'].value}. Every figure traces to the Deals sheet.")
    ws["A2"].font = NOTE
    _header(ws, 4, ["Sales rep", "Region", "FY commission (USD)", "Selected qtr (USD)", "Deals"],
            [20, 10, 20, 18, 8])
    s = src["Summary"]
    for i, r in enumerate(range(5, 17)):
        ws.cell(5 + i, 1, s.cell(r, 1).value)
        ws.cell(5 + i, 2, s.cell(r, 2).value)
        ws.cell(5 + i, 3, outputs[f"Summary!C{r}"]).number_format = MONEY
        ws.cell(5 + i, 4, outputs[f"Summary!D{r}"]).number_format = MONEY
        ws.cell(5 + i, 5, outputs[f"Summary!E{r}"])
    _header(ws, 18, ["Region", "", "FY commission (USD)", "", ""])
    for i, r in enumerate(range(19, 23)):
        ws.cell(19 + i, 1, s.cell(r, 1).value)
        ws.cell(19 + i, 3, outputs[f"Summary!C{r}"]).number_format = MONEY
    rows = [("Company total commission", outputs["Summary!C24"], MONEY,
             "Sum of the rep lines above, so it carries the adj amounts - ledger L17, L18."),
            ("Total bookings (USD)", outputs["Summary!C25"], MONEY,
             "Four text-typed deal values are skipped and EUR goes in unconverted - L03, L06."),
            ("Commission as % of bookings", outputs["Summary!C26"], PCT,
             "Numerator and denominator are not like for like - ledger L19.")]
    for i, (label, val, fmt, note) in enumerate(rows):
        ws.cell(24 + i, 1, label)
        ws.cell(24 + i, 3, val).number_format = fmt
        ws.cell(24 + i, 6, note).font = NOTE
    ws.cell(29, 1, f"Region lines total {sum(outputs[f'Summary!C{r}'] for r in range(19, 23)):,.2f}, "
                   f"which is {outputs['Summary!C24'] - sum(outputs[f'Summary!C{r}'] for r in range(19, 23)):,.2f} "
                   f"less than the company total (ledger L17).").font = NOTE
    return ws


def notes(wb, run):
    """The ledger's problems, worst first, so a reader of the clean model sees them too."""
    ws = wb.create_sheet("Notes")
    ws["A1"] = "What the evidence ledger found"
    ws["A1"].font = TITLE
    ws["A2"] = ("Problems first, then the neutral descriptions and the open questions. "
                "Full detail, with evidence, in ledger.md.")
    ws["A2"].font = NOTE
    _header(ws, 4, ["ID", "Problem?", "Tag", "Finding", "Where", "Who could answer it"],
            [7, 10, 11, 96, 34, 26])
    with open(run / "ledger.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    rows.sort(key=lambda r: (r["problem"].upper() != "Y", r["id"]))
    for i, r in enumerate(rows):
        out = 5 + i
        ws.cell(out, 1, r["id"])
        ws.cell(out, 2, "problem" if r["problem"].upper() == "Y" else "-")
        ws.cell(out, 3, r["tag"])
        c = ws.cell(out, 4, r["finding"]); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(out, 5, r["location"]).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(out, 6, r["ask"]).alignment = Alignment(wrap_text=True, vertical="top")
        if r["problem"].upper() == "Y":
            ws.cell(out, 2).font = Font(bold=True, color="B00020")
    return ws


def main():
    import json
    import importlib.util
    from sheetre import excelfn as X

    graph = json.loads((RUN / "graph.json").read_text())
    src = openpyxl.load_workbook(SRC, data_only=True)
    inputs = {a: src[a.split("!")[0]][a.split("!")[1]].value for a in graph["inputs"]}

    spec = importlib.util.spec_from_file_location("m", Path(__file__).parent / "model.py")
    model = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(model)
    outputs = model.run(inputs, None)

    # Re-walk the rows so the Deals sheet can show every intermediate step, not just the outputs.
    c = X.Cells(inputs)
    rate_table = c.table("Rates!J2:K13")
    model_rows = {}
    for r in model.DEAL_ROWS:
        row = model._calc_row(r, c)
        row["participation"] = X.iferror(
            lambda: X.vlookup(row["B"], rate_table, 2, approx=False), 0)
        row["quarter"] = "Q%d" % int(X.roundup(model._month_of(row["E"]) / 3, 0))
        model_rows[r] = row

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary(wb, outputs, src)
    assumptions(wb, src)
    deals(wb, src, model_rows)
    notes(wb, RUN)
    wb.save(OUT)
    print(f"wrote {OUT} ({', '.join(wb.sheetnames)})")


if __name__ == "__main__":
    main()
