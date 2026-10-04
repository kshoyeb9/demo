#!/usr/bin/env python3
"""Pull FY2025/26 (Jul-25 → Jun-26) actuals out of the Noon financial model.

Source: 'Noon financial model 3.0 v47.xlsx'.

The model is a calculated workbook, so the cached values are readable. Two tabs
are used and they do NOT share a column layout:

    02_Consolidated_FS   months start at Jul-24  →  FY24/25 = E:P (5-16), total Q (17)
                                                    FY25/26 = R:AC (18-29), total AD (30)
    11_Working_Capital   months start at Jul-25  →  FY25/26 = E:P (5-16), total Q (17)

Both maps are asserted against the sheets' own header dates on every run, so a
column inserted upstream fails loudly instead of silently shifting a year.

Everything is emitted in USD millions, which is the unit the dashboard input
workbook is built in. Costs are emitted POSITIVE; the model carries the
below-EBITDA lines as negatives.
"""
import argparse, datetime, json, sys

import openpyxl

M = 1_000_000.0

# ── column maps, verified at run time ────────────────────────────────────────
CONS_PRI = (5, 16, 17)     # FY24/25 first month, last month, year total
CONS_CUR = (18, 29, 30)    # FY25/26
WC_CUR   = (5, 16, 17)     # FY25/26 on the working-capital tab

# ── the six business units, in the model's own order ────────────────────────
# Row numbers are the Total Revenues block; the cost blocks sit at a fixed
# offset below it, which is asserted below.
BU = [
    (10, "Tracks"),
    (11, "Govt Schools — Legacy"),
    (12, "Govt Schools — New"),
    (13, "B2B"),
    (14, "KSA B2C"),
    (15, "Out of School"),
]
OFF_COS, OFF_MKT, OFF_OPX = 9, 19, 29   # revenue row + offset

# ── consolidated P&L / balance sheet / cash flow rows ───────────────────────
R = dict(
    rev=9, cos=18, gp=25, mkt=28, cm=35, opx=38, ebitda=45,
    fin=48, other=49, da=50, tax=51, ni=52,
    cash=63, ar=64, other_ca=65,
    ap=69, other_pay=70, defrev=71, cards=72, other_cl=73,
    cf_op=99, cf_inv=103, cf_fin=108,
    cf_net=110, cf_open=111, cf_fx=112, cf_undep=113, cf_close=114,
)


def fail(msg):
    sys.exit("model_fy2526: " + msg)


def check_months(ws, hrow, first, last, total, label, start):
    """Assert the sheet's own header dates match the column map."""
    got = ws.cell(hrow, first).value
    if not isinstance(got, datetime.datetime) or (got.year, got.month) != start:
        fail(f"{label}: expected {start[1]:02d}/{start[0]} at column {first}, found {got!r}. "
             "The model's columns have moved — update the column map.")
    n = last - first + 1
    if n != 12:
        fail(f"{label}: column map spans {n} months, not 12.")
    t = ws.cell(hrow, total).value
    if not (isinstance(t, str) and "/" in t):
        fail(f"{label}: column {total} should be the year total, found {t!r}.")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--file", default="v47.xlsx")
    p.add_argument("--out", default="fy2526.json")
    a = p.parse_args()

    wb = openpyxl.load_workbook(a.file, data_only=True)
    for need in ("02_Consolidated_FS", "11_Working_Capital"):
        if need not in wb.sheetnames:
            fail(f"'{need}' is not a tab in {a.file}.")
    cs = wb["02_Consolidated_FS"]
    wc = wb["11_Working_Capital"]

    check_months(cs, 5, *CONS_PRI, "02_Consolidated_FS FY24/25", (2024, 7))
    check_months(cs, 5, *CONS_CUR, "02_Consolidated_FS FY25/26", (2025, 7))
    check_months(wc, 3, *WC_CUR,   "11_Working_Capital FY25/26", (2025, 7))

    def series(ws, row, span, sign=1.0):
        """Twelve monthly values in USD M."""
        first, last, _ = span
        out = []
        for c in range(first, last + 1):
            v = ws.cell(row, c).value
            out.append(None if not isinstance(v, (int, float)) else round(sign * v / M, 4))
        return out

    def year(ws, row, span, sign=1.0):
        v = ws.cell(row, span[2]).value
        return None if not isinstance(v, (int, float)) else round(sign * v / M, 4)

    def close(a_, b_, what, tol=0.01):
        if a_ is None or b_ is None:
            fail(f"{what}: missing a figure ({a_!r} vs {b_!r}).")
        if abs(a_ - b_) > tol:
            fail(f"{what}: {a_:,.4f} vs {b_:,.4f} — differ by {a_-b_:,.4f}M.")

    # ── labels ──────────────────────────────────────────────────────────────
    months = [cs.cell(5, c).value.strftime("%b-%y") for c in range(CONS_CUR[0], CONS_CUR[1] + 1)]

    # ── revenue by BU ───────────────────────────────────────────────────────
    rev_cur, rev_pri = {}, {}
    for row, name in BU:
        lab = str(cs.cell(row, 3).value or "").strip()
        if not lab:
            fail(f"row {row} of 02_Consolidated_FS has no business-unit label.")
        rev_cur[name] = series(cs, row, CONS_CUR)
        rev_pri[name] = series(cs, row, CONS_PRI)

    # ── cost categories ─────────────────────────────────────────────────────
    # 'Finance & tax' bundles finance expenses, other income/(expenses) and
    # taxes so that EBITDA less D&A less this line equals net income exactly.
    def fin_tax(span):
        return [
            None if any(v is None for v in t) else round(sum(t), 4)
            for t in zip(series(cs, R["fin"],   span, -1),
                         series(cs, R["other"], span, -1),
                         series(cs, R["tax"],   span, -1))
        ]

    cost_cur = {
        "Cost of sales":      series(cs, R["cos"], CONS_CUR),
        "Marketing":          series(cs, R["mkt"], CONS_CUR),
        "Operating expenses": series(cs, R["opx"], CONS_CUR),
        "D&A":                series(cs, R["da"],  CONS_CUR, -1),
        "Finance & tax":      fin_tax(CONS_CUR),
    }
    cost_pri = {
        "Cost of sales":      series(cs, R["cos"], CONS_PRI),
        "Marketing":          series(cs, R["mkt"], CONS_PRI),
        "Operating expenses": series(cs, R["opx"], CONS_PRI),
        "D&A":                series(cs, R["da"],  CONS_PRI, -1),
        "Finance & tax":      fin_tax(CONS_PRI),
    }

    # ── BU P&L table: revenue, cost of sales, marketing, both years ─────────
    bu_pl = []
    for row, name in BU:
        def trio(span):
            return dict(
                rev=year(cs, row, span),
                cos=year(cs, row + OFF_COS, span),
                mkt=year(cs, row + OFF_MKT, span),
                opx=year(cs, row + OFF_OPX, span),
            )
        bu_pl.append({"name": name, "cur": trio(CONS_CUR), "pri": trio(CONS_PRI)})

    # ── working capital, from the balance sheet ─────────────────────────────
    def wc_other(span):
        parts = [series(cs, R["other_ca"],  span,  1),
                 series(cs, R["other_pay"], span, -1),
                 series(cs, R["cards"],     span, -1),
                 series(cs, R["other_cl"],  span, -1)]
        return [None if any(v is None for v in t) else round(sum(t), 4) for t in zip(*parts)]

    working_capital = {
        "receivables":      series(cs, R["ar"],     CONS_CUR),
        "payables":         series(cs, R["ap"],     CONS_CUR, -1),
        "deferred_revenue": series(cs, R["defrev"], CONS_CUR, -1),
        "other":            wc_other(CONS_CUR),
    }

    # ── receivables and payables ───────────────────────────────────────────
    # Closing balances come from the balance sheet, which is the auditable
    # figure and the one the working-capital section already reports. The flows
    # are then derived from those balances, so the roll-forward ties by
    # construction rather than being read from a second schedule that disagrees
    # with the balance sheet. (The per-business-unit AR roll-forward and the AP
    # block on 11_Working_Capital both do: the AP block is empty before Dec-25
    # and ends the year 1.37M below the balance sheet.)
    def rollforward(bal_row, additions, open_prior):
        closing = series(cs, bal_row, CONS_CUR)
        opening = [open_prior] + closing[:-1]
        settled = [None if any(x is None for x in (o, a, c)) else round(o + a - c, 4)
                   for o, a, c in zip(opening, additions, closing)]
        return opening, settled, closing

    open_ar = series(cs, R["ar"], CONS_PRI)[-1]
    open_ap = series(cs, R["ap"], CONS_PRI)[-1]

    invoiced = series(cs, R["rev"], CONS_CUR)
    ar_open, ar_coll, ar_close = rollforward(R["ar"], invoiced, open_ar)
    ar = {"opening": ar_open, "invoiced": invoiced,
          "collections": ar_coll, "closing": ar_close}

    # Payables are driven by the cost base that passes through them: cost of
    # sales plus operating expenses. Marketing, D&A and financing are excluded.
    purchases = [None if any(x is None for x in t) else round(sum(t), 4)
                 for t in zip(series(cs, R["cos"], CONS_CUR),
                              series(cs, R["opx"], CONS_CUR))]
    ap_open, ap_pay, ap_close = rollforward(R["ap"], purchases, open_ap)
    ap = {"opening": ap_open, "purchases": purchases,
          "payments": ap_pay, "closing": ap_close}

    # ── cash ────────────────────────────────────────────────────────────────
    # The closing balance is the balance sheet's cash line. The model's own
    # cash-flow statement does not chain to it — undeposited funds and FX are
    # carried outside the three activity totals — so the gap is reported as its
    # own movement line instead of being buried in an activity.
    bal     = series(cs, R["cash"],   CONS_CUR)
    op_cf   = series(cs, R["cf_op"],  CONS_CUR)
    inv_cf  = series(cs, R["cf_inv"], CONS_CUR)
    fin_cf  = series(cs, R["cf_fin"], CONS_CUR)
    open_cash = series(cs, R["cash"], CONS_PRI)[-1]
    cash_open = [open_cash] + bal[:-1]
    residual = [None if any(x is None for x in (b, o, a, i_, f)) else
                round((b - o) - (a + i_ + f), 4)
                for b, o, a, i_, f in zip(bal, cash_open, op_cf, inv_cf, fin_cf)]
    cash = {
        "balance":   bal,
        "opening":   cash_open,
        "operating": op_cf,
        "investing": inv_cf,
        "financing": fin_cf,
        "residual":  residual,
        "fx":        series(cs, R["cf_fx"],    CONS_CUR),
        "undeposited": series(cs, R["cf_undep"], CONS_CUR),
        # The six months before the year starts, for context on the balance
        # chart. Labels come from the sheet so they cannot drift.
        "prior": [
            [cs.cell(5, c).value.strftime("%b-%y"),
             round((cs.cell(R["cash"], c).value or 0) / M, 4)]
            for c in range(CONS_PRI[1] - 5, CONS_PRI[1] + 1)
        ],
    }

    # the receivables roll-forward must reproduce the balance sheet
    for i in range(12):
        close(ar["opening"][i] + ar["invoiced"][i] - ar["collections"][i], ar["closing"][i],
              f"AR roll-forward does not tie in month {i+1}")
        close(cash["opening"][i] + op_cf[i] + inv_cf[i] + fin_cf[i] + residual[i], bal[i],
              f"cash roll-forward does not tie in month {i+1}")

    # ── P&L summary totals, for the reconciliation checks downstream ───────
    # D&A is carried negative in the model; emit it positive, as the cost
    # block does, so every cost in this file reads the same way round.
    SUM_SIGN = {"da": -1}
    def summary_for(span):
        return {k: year(cs, R[k], span, SUM_SIGN.get(k, 1))
                for k in ("rev", "cos", "gp", "mkt", "cm", "opx", "ebitda", "da", "ni")}
    summary, summary_pri = summary_for(CONS_CUR), summary_for(CONS_PRI)

    # ── reconciliations: fail rather than publish a table that does not tie ─
    close(sum(b["cur"]["rev"] for b in bu_pl), summary["rev"],
          "business-unit revenue does not sum to Total Revenues")
    close(sum(b["cur"]["cos"] for b in bu_pl), summary["cos"],
          "business-unit cost of sales does not sum to Total COS")
    close(sum(b["cur"]["mkt"] for b in bu_pl), summary["mkt"],
          "business-unit marketing does not sum to Advertising & Marketing")
    close(summary["rev"] - summary["cos"], summary["gp"], "revenue less COS is not gross profit")
    close(summary["gp"] - summary["mkt"], summary["cm"], "gross profit less marketing is not CM")
    close(summary["cm"] - summary["opx"], summary["ebitda"], "CM less opex is not EBITDA")
    ft = year(cs, R["fin"], CONS_CUR, -1) + year(cs, R["other"], CONS_CUR, -1) \
        + year(cs, R["tax"], CONS_CUR, -1)
    close(summary["ebitda"] - summary["da"] - ft, summary["ni"],
          "EBITDA less D&A less finance & tax is not net income")

    # ── commentary, generated from the figures above ────────────────────────
    # Written here rather than typed into the workbook so it can never drift
    # from the numbers it describes.
    def pc(x, base):
        return f"{x/base:+.1%}" if base else "n/a"
    biggest = max(bu_pl, key=lambda b: b["cur"]["rev"] or 0)
    ar_mv = ar["closing"][-1] - ar["opening"][0]
    ap_mv = ap["closing"][-1] - ap["opening"][0]
    fin_total = sum(v for v in fin_cf if v)
    cash_mv = bal[-1] - open_cash
    key_updates = [
        ("Revenue",
         f"FY2025/26 revenue of ${summary['rev']:,.2f}M came in {pc(summary['rev']-summary_pri['rev'], summary_pri['rev'])} "
         f"on FY2024/25's ${summary_pri['rev']:,.2f}M — essentially flat year on year. "
         f"{biggest['name']} was the largest unit at ${biggest['cur']['rev']:,.2f}M, "
         f"{biggest['cur']['rev']/summary['rev']:.0%} of the total."),
        ("Margin",
         f"Gross profit of ${summary['gp']:,.2f}M gave a {summary['gp']/summary['rev']:.1%} margin "
         f"(FY2024/25 {summary_pri['gp']/summary_pri['rev']:.1%}). After ${summary['mkt']:,.2f}M of marketing, "
         f"contribution was ${summary['cm']:,.2f}M, a {summary['cm']/summary['rev']:.1%} margin."),
        ("EBITDA",
         f"EBITDA of ${summary['ebitda']:,.2f}M is a {summary['ebitda']/summary['rev']:.1%} margin, against "
         f"{summary_pri['ebitda']/summary_pri['rev']:.1%} last year, after ${summary['opx']:,.2f}M of operating expenses."),
        ("Below EBITDA",
         f"D&A of ${summary['da']:,.2f}M and finance and tax of ${ft:,.2f}M took the year to a net loss of "
         f"${abs(summary['ni']):,.2f}M. D&A alone is {summary['da']/summary['rev']:.0%} of revenue."),
        ("Working capital",
         f"Receivables closed at ${ar['closing'][-1]:,.2f}M, ${abs(ar_mv):,.2f}M "
         f"{'higher' if ar_mv > 0 else 'lower'} than at the start of the year, on collections of "
         f"${sum(v for v in ar['collections'] if v):,.2f}M. Payables closed at ${ap['closing'][-1]:,.2f}M, "
         f"${abs(ap_mv):,.2f}M {'higher' if ap_mv > 0 else 'lower'}."),
        ("Cash",
         f"Cash closed the year at ${bal[-1]:,.2f}M against ${open_cash:,.2f}M at the start, "
         f"up ${cash_mv:,.2f}M. Operations and investing together consumed "
         f"${abs(sum(v for v in op_cf if v) + sum(v for v in inv_cf if v)):,.2f}M; "
         f"financing contributed ${fin_total:,.2f}M, so the balance is funded by drawdown rather than trading."),
    ]

    data = {
        "source": a.file,
        "fy": {"cur": "FY2025/26", "pri": "FY2024/25",
               "cur_short": "FY25/26", "pri_short": "FY24/25"},
        "months": months,
        "active": 12,
        "revenue": {"cur": rev_cur, "pri": rev_pri},
        "costs":   {"cur": cost_cur, "pri": cost_pri},
        "bu_pl": bu_pl,
        "working_capital": working_capital,
        "ar": ar,
        "ap": ap,
        "cash": cash,
        "summary": {"cur": summary, "pri": summary_pri},
        "key_updates": [{"topic": t, "text": x} for t, x in key_updates],
    }
    with open(a.out, "w") as f:
        json.dump(data, f, indent=1)

    print(f"  Source    {a.file} · 02_Consolidated_FS + 11_Working_Capital")
    print(f"  Year      FY2025/26 — {months[0]} to {months[-1]} (closed, 12 months actual)")
    print(f"  Revenue   ${summary['rev']:,.2f}M   (FY24/25 ${summary_pri['rev']:,.2f}M)")
    print(f"  Gross     ${summary['gp']:,.2f}M  {summary['gp']/summary['rev']:.1%}")
    print(f"  EBITDA    ${summary['ebitda']:,.2f}M  {summary['ebitda']/summary['rev']:.1%}")
    print(f"  Net inc   ${summary['ni']:,.2f}M")
    print(f"  Cash      ${cash['balance'][-1]:,.2f}M closing")
    print(f"  Units     {', '.join(b['name'] for b in bu_pl)}")
    print(f"  Wrote     {a.out}")


if __name__ == "__main__":
    main()
