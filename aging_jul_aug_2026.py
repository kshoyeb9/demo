#!/usr/bin/env python3
"""Receivables and payables ageing for Jul-26 and Aug-26.

Payables come from the A/P Aging Summary workbook, vendor by vendor, for both
months. Receivables come from the reporting pack's AR ageing and AR-by-contract
blocks for the same two months, which are transcribed below — July is held
there to two decimals, August to full precision.

The two sides behave differently and the dashboard says so:

  · the AR ageing ties exactly to receivables on the balance sheet
  · the AP subledger does not tie to balance-sheet payables, being short by the
    accruals and other payables that never pass through a vendor account; the
    gap is reported on the page rather than adjusted away

Usage:  python aging_jul_aug_2026.py <AP_Aging.xlsx> <ar_inputs.json> <out.json>
"""
import json, sys

import openpyxl

M = 1_000_000.0
MONTHS = ["Jul-26", "Aug-26"]
TOP_N = 10

AP_BUCKETS = ["Current", "1–30 days", "31–60 days", "61–90 days", "90+ days"]


def main(xlsx, ar_inputs, out="aging_2026.json"):
    # The receivables ageing, the counterparty balances and the balance-sheet
    # positions they are checked against all live in an input file rather than
    # in this script, which is committed to a public repository.
    _in = json.load(open(ar_inputs))
    BS = _in["balance_sheet"]
    AR_AGING = {m: [tuple(x) for x in v] for m, v in _in["ar_aging"].items()}
    AR_CONTRACT = {m: [tuple(x) for x in v] for m, v in _in["ar_by_contract"].items()}

    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb[wb.sheetnames[0]]

    hdr = [ws.cell(4, c).value for c in range(1, 14)]
    want = ["Vendor", "Current", "1-30", "31-60", "61-90", ">90", "Total",
            "Current", "1-30", "31-60", "61-90", ">90", "Total"]
    if hdr != want:
        sys.exit(f"unexpected header row: {hdr}\nexpected: {want}")
    if str(ws.cell(3, 2).value).strip() != "July 2026" or \
       str(ws.cell(3, 8).value).strip() != "August 2026":
        sys.exit("the two month blocks are not July 2026 and August 2026.")

    # vendor rows, with the two month blocks at columns B:G and H:M
    vendors, stated = [], {}
    for r in range(5, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        name = str(name).strip()
        cells = [ws.cell(r, c).value or 0 for c in range(2, 14)]
        if name.lower() == "total":
            stated = {"Jul-26": cells[5], "Aug-26": cells[11]}
            continue
        vendors.append((name, {"Jul-26": cells[0:6], "Aug-26": cells[6:12]}))
    if not vendors or not stated:
        sys.exit("no vendor rows or no Total row found.")

    by_month = {}
    for mi, mon in enumerate(MONTHS):
        rows = [(n, v[mon]) for n, v in vendors]
        total = sum(v[5] for _, v in rows)
        if abs(total - stated[mon]) > 1.0:
            sys.exit(f"{mon}: vendor rows sum to {total:,.2f} but the workbook's "
                     f"Total reads {stated[mon]:,.2f}.")
        buckets = [sum(v[i] for _, v in rows) for i in range(5)]
        if abs(sum(buckets) - total) > 1.0:
            sys.exit(f"{mon}: the ageing buckets do not sum to the vendor total.")

        ap_aging = [{"bucket": b, "amount": round(buckets[i] / M, 4),
                     "share": round(buckets[i] / total, 6) if total else None}
                    for i, b in enumerate(AP_BUCKETS)]

        ranked = sorted(rows, key=lambda x: -x[1][5])
        top, rest = ranked[:TOP_N], ranked[TOP_N:]
        ap_vendor = [{"vendor": n, "amount": round(v[5] / M, 4)} for n, v in top]
        if rest:
            ap_vendor.append({"vendor": f"Other vendors ({len(rest)})",
                              "amount": round(sum(v[5] for _, v in rest) / M, 4)})

        ar_ag = [{"bucket": b, "amount": a} for b, a in AR_AGING[mon]]
        ar_tot = round(sum(x["amount"] for x in ar_ag), 4)
        for x in ar_ag:
            x["share"] = round(x["amount"] / ar_tot, 6) if ar_tot else None
        if abs(ar_tot - BS[mon]["ar"]) > 0.02:
            sys.exit(f"{mon}: the AR ageing sums to {ar_tot:,.4f}M but receivables "
                     f"on the balance sheet are {BS[mon]['ar']:,.4f}M.")

        ar_con = [{"contract": c, "amount": a} for c, a in AR_CONTRACT[mon]]
        con_tot = round(sum(x["amount"] for x in ar_con), 4)
        if abs(con_tot - BS[mon]["ar"]) > 0.02:
            sys.exit(f"{mon}: AR by contract sums to {con_tot:,.4f}M against "
                     f"{BS[mon]['ar']:,.4f}M on the balance sheet.")

        ap_tot = round(total / M, 4)
        by_month[mon] = {
            "ar_aging": ar_ag, "ar_aging_total": ar_tot,
            "ar_by_contract": ar_con, "ar_by_contract_total": con_tot,
            "ap_aging": ap_aging, "ap_aging_total": ap_tot,
            "ap_by_vendor": ap_vendor,
            # how far the vendor subledger sits from the balance sheet
            "ap_unreconciled": round(BS[mon]["ap"] - ap_tot, 4),
            "ap_balance_sheet": BS[mon]["ap"],
        }
        print(f"  {mon}")
        print(f"    AR ageing     ${ar_tot:,.4f}M  ties to the balance sheet")
        print(f"    AR contracts  ${con_tot:,.4f}M  across {len(ar_con)} counterparties")
        print(f"    AP subledger  ${ap_tot:,.4f}M  vs ${BS[mon]['ap']:,.4f}M on the "
              f"balance sheet — ${BS[mon]['ap']-ap_tot:,.4f}M outside it")
        print(f"    AP vendors    {len(vendors)} rows, top {TOP_N} shown")

    with open(out, "w") as f:
        json.dump({"months": MONTHS, "by_month": by_month}, f,
                  ensure_ascii=False, separators=(",", ":"))
    print(f"  Wrote {out}")


if __name__ == "__main__":
    main(*sys.argv[1:])
