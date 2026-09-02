"""
etl_pl.py — Monthly refresh for the P&L Performance HTML dashboard.

Usage:
    python etl_pl.py --file "pl_only_final.xlsx" [--out index.html]

What it does:
    1. Reads the single "P&L Performance" sheet in the workbook.
    2. Extracts every structured section (P&L, costs, WC, AR, AP, cash, key updates).
    3. Builds the DATA JSON object expected by the HTML dashboard template.
    4. Writes a fresh index.html by replacing the DATA constant in the template.

After running, open index.html in any browser — no server required.
"""

import argparse
import json
import math
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl not installed — run: pip install openpyxl")

# ── CLI ───────────────────────────────────────────────────────────────────────
ap = argparse.ArgumentParser()
ap.add_argument("--file", default="pl_only_final.xlsx", help="Path to the P&L workbook")
ap.add_argument("--out",  default="index.html",         help="Output HTML file")
ap.add_argument("--template", default=None,             help="HTML template (defaults to existing --out file)")
args = ap.parse_args()

xl_path  = Path(args.file)
out_path = Path(args.out)
tpl_path = Path(args.template) if args.template else out_path

if not xl_path.exists():
    sys.exit(f"Workbook not found: {xl_path}")
if not tpl_path.exists():
    sys.exit(f"HTML template not found: {tpl_path}\n"
             f"Place the original dashboard HTML at {tpl_path} before running.")

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(xl_path, data_only=True)
if "P&L Performance" not in wb.sheetnames:
    sys.exit("Sheet 'P&L Performance' not found in the workbook.")
ws = wb["P&L Performance"]

# Read all rows into a list (0-indexed)
rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

def row(n):          # 1-indexed access
    if n < 1 or n > len(rows):
        return ()
    return rows[n - 1]

def v(n, col):       # safe cell read
    r = row(n)
    return r[col] if col < len(r) else None

def nz(x):           # None → 0
    return x if x is not None else 0

def rnd(x, d=4):
    return round(x, d) if x is not None else None

# ── Parse meta from header row 5 ─────────────────────────────────────────────
# "Key Metrics — June 2026  (Month)"
header_m = str(v(5, 10) or "")
period_full = re.search(r"—\s+(.+?)\s+\(", header_m)
period_full = period_full.group(1).strip() if period_full else "Unknown"

# "Key Metrics — Year-to-Date  (Jan–Jun 2026)"
header_y = str(v(5, 20) or "")
ytd_match = re.search(r"\((.+?)\)", header_y)
ytd_label = ytd_match.group(1).strip() if ytd_match else ""

# Extract months list from YTD label e.g. "Jan–Jun 2026"
month_abbrs = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
ytd_months_match = re.search(r"(\w+)[–-](\w+)\s+(\d{4})", ytd_label)
if ytd_months_match:
    m_start, m_end, year_str = ytd_months_match.groups()
    i0 = month_abbrs.index(m_start[:3]) if m_start[:3] in month_abbrs else 0
    i1 = month_abbrs.index(m_end[:3])   if m_end[:3]   in month_abbrs else 5
    months = month_abbrs[i0:i1+1]
    year   = year_str
else:
    months = month_abbrs[:6]
    year   = str(datetime.now().year)

period_index = len(months) - 1  # 0-indexed

# Fiscal year label (Jul start)
fy_start = int(year) - 1 if int(months[0][:3] in ["Jan","Feb","Mar","Apr","May","Jun"] and True or False) else int(year)
fiscal_year = f"FY{str(fy_start)[2:]}/{str(fy_start+1)[2:]}"

meta = {
    "period":        period_full,
    "ytd_label":     ytd_label,
    "generated":     datetime.now().strftime("%d %b %Y"),
    "source":        xl_path.name,
    "currency":      "USD M",
    "months":        months,
    "year":          year,
    "period_index":  period_index,
    "fy_start_month": 7,
    "fiscal_year":   fiscal_year,
}

# ── Helper: parse an actual/budget/var/pct row ────────────────────────────────
def parse_avb(row_n, col_offset=0):
    """Read 8 values: act_mo, bgt_mo, var_mo, pct_mo, act_ytd, bgt_ytd, var_ytd, pct_ytd."""
    c = col_offset
    a_m = nz(v(row_n, c+1)); b_m = nz(v(row_n, c+2))
    a_y = nz(v(row_n, c+5)); b_y = nz(v(row_n, c+6))
    return {
        "month": {"actual": rnd(a_m,4), "budget": rnd(b_m,4),
                  "var": rnd(a_m-b_m,4), "pct": rnd(a_m/b_m,4) if b_m else None},
        "ytd":   {"actual": rnd(a_y,4), "budget": rnd(b_y,4),
                  "var": rnd(a_y-b_y,4), "pct": rnd(a_y/b_y,4) if b_y else None},
    }

def _sum_rows(rows_list, key1, key2):
    """Sum actual/budget across non-total rows — used when total row contains formulas."""
    return (
        round(sum(r[key1][key2] or 0 for r in rows_list if not r["total"]), 4)
        or 0
    )

def _fix_total(entry, detail_rows):
    """If the total row has zero actual (formula not cached), derive from detail rows."""
    for period in ("month", "ytd"):
        a = entry[period]["actual"]
        b = entry[period]["budget"]
        if not a:
            a = round(sum(r[period]["actual"] or 0 for r in detail_rows if not r["total"]), 4)
            entry[period]["actual"] = a
        if not b:
            b = round(sum(r[period]["budget"] or 0 for r in detail_rows if not r["total"]), 4)
            entry[period]["budget"] = b
        entry[period]["var"] = round(a - b, 4)
        entry[period]["pct"] = round(a / b, 4) if b else None

# ── Revenue by BU (rows 15-20) ────────────────────────────────────────────────
revenue = []
r_n = 15
while True:
    name = v(r_n, 0)
    if not name or str(name).strip() == "":
        break
    data = parse_avb(r_n)
    is_total = "total" in str(name).lower()
    revenue.append({
        "name":  str(name),
        **data,
        "total": is_total,
        "monthly": None,
    })
    if is_total:
        break
    r_n += 1

# If total row had uncached formulas, derive from detail rows
if revenue and revenue[-1]["total"]:
    _fix_total(revenue[-1], revenue[:-1])
elif revenue:
    # Add synthetic total if the scan never hit a "total" row
    am = round(sum(r["month"]["actual"] or 0 for r in revenue), 4)
    bm = round(sum(r["month"]["budget"] or 0 for r in revenue), 4)
    ay = round(sum(r["ytd"]["actual"]   or 0 for r in revenue), 4)
    by = round(sum(r["ytd"]["budget"]   or 0 for r in revenue), 4)
    revenue.append({"name": "Total Revenue", "total": True, "monthly": None,
                    "month": {"actual":am,"budget":bm,"var":round(am-bm,4),"pct":round(am/bm,4) if bm else None},
                    "ytd":   {"actual":ay,"budget":by,"var":round(ay-by,4),"pct":round(ay/by,4) if by else None}})

# ── Costs by category (rows 24-29) ───────────────────────────────────────────
costs = []
r_n = 24
while True:
    name = v(r_n, 0)
    if not name or str(name).strip() == "":
        break
    data = parse_avb(r_n)
    is_total = "total" in str(name).lower()
    costs.append({
        "name":  str(name),
        **data,
        "total": is_total,
        "monthly": None,
    })
    if is_total:
        break
    r_n += 1

if costs and costs[-1]["total"]:
    _fix_total(costs[-1], costs[:-1])

# ── P&L summary rows (rows 34-41) ────────────────────────────────────────────
# Fixed structure: Revenue, Direct costs, Gross profit, Operating expenses, EBITDA
# Plus margin % rows

# Revenue and Total costs from above
rev_row  = next((r for r in revenue if r["total"]), revenue[-1]) if revenue else {}
cost_row = next((r for r in costs   if r["total"]), costs[-1])   if costs   else {}

def _pct_row(num_name, den_name, pl_rows):
    num = next((r for r in pl_rows if r["name"] == num_name), None)
    den = next((r for r in pl_rows if r["name"] == den_name), None)
    if not num or not den:
        return None
    am = num["month"]["actual"]; bm = num["month"]["budget"]
    ay = num["ytd"]["actual"];   by = num["ytd"]["budget"]
    drm = (den["month"]["actual"] or 1); dry = (den["ytd"]["actual"] or 1)
    dbm = (den["month"]["budget"] or 1); dby = (den["ytd"]["budget"] or 1)
    return {
        "month": {"actual": rnd(am/drm,4), "budget": rnd(bm/dbm,4),
                  "var": rnd(am/drm - bm/dbm, 4), "pct": None},
        "ytd":   {"actual": rnd(ay/dry,4), "budget": rnd(by/dby,4),
                  "var": rnd(ay/dry - by/dby, 4), "pct": None},
    }

# Build P&L from parsed rows 34-38 and margin rows 40-41
pl = []
pl_map = {}  # name → parse_avb result, for rows 34-38
r_n = 34
while r_n <= 41:
    name = v(r_n, 0)
    if name and str(name).strip():
        data = parse_avb(r_n)
        pl_map[str(name)] = data
    r_n += 1

def _pl_entry(name, kind="value", cost=False, total=False):
    d = pl_map.get(name)
    if not d:
        return None
    return {"name": name, "kind": kind, **d, "cost": cost, "total": total, "monthly": None}

# Gross margin % from rows 40-41; if formula not cached, derive from P&L rows
gm_act_m  = rnd(nz(v(40, 1)), 4)
gm_act_y  = rnd(nz(v(40, 5)), 4)
ebi_act_m = rnd(nz(v(41, 1)), 4)
ebi_act_y = rnd(nz(v(41, 5)), 4)

gp_row_fallback  = pl_map.get("Gross profit / contribution") or pl_map.get("Gross profit")
ebi_row_fallback = pl_map.get("EBITDA")
rev_act_m_fb = rev_row.get("month",{}).get("actual") or 1
rev_act_y_fb = rev_row.get("ytd",{}).get("actual")   or 1

if not gm_act_m and gp_row_fallback:
    gm_act_m = rnd((gp_row_fallback["month"]["actual"] or 0) / rev_act_m_fb, 4)
if not gm_act_y and gp_row_fallback:
    gm_act_y = rnd((gp_row_fallback["ytd"]["actual"]   or 0) / rev_act_y_fb, 4)
if not ebi_act_m and ebi_row_fallback:
    ebi_act_m = rnd((ebi_row_fallback["month"]["actual"] or 0) / rev_act_m_fb, 4)
if not ebi_act_y and ebi_row_fallback:
    ebi_act_y = rnd((ebi_row_fallback["ytd"]["actual"]   or 0) / rev_act_y_fb, 4)

# Budget margins: gp_budget / rev_budget
rev_bgt_m = rev_row.get("month",{}).get("budget") or 1
rev_bgt_y = rev_row.get("ytd",{}).get("budget")   or 1
gp_bgt_m  = nz(pl_map.get("Gross profit / contribution",{}).get("month",{}).get("budget") if "Gross profit / contribution" in pl_map
               else next((r["month"]["budget"] for r in revenue if r["total"]),0) * gm_act_m)
ebi_bgt_m = pl_map.get("EBITDA",{}).get("month",{}).get("budget") or 0
ebi_bgt_y = pl_map.get("EBITDA",{}).get("ytd",{}).get("budget")   or 0

# Try to get gross profit budget directly
gp_row_d = pl_map.get("Gross profit / contribution") or pl_map.get("Gross profit")
if gp_row_d:
    gm_bgt_m = rnd((gp_row_d["month"]["budget"] or 0) / rev_bgt_m, 4)
    gm_bgt_y = rnd((gp_row_d["ytd"]["budget"]   or 0) / rev_bgt_y, 4)
else:
    gm_bgt_m = gm_act_m
    gm_bgt_y = gm_act_y

ebi_bgt_pct_m = rnd(ebi_bgt_m / rev_bgt_m, 4) if rev_bgt_m else ebi_act_m
ebi_bgt_pct_y = rnd(ebi_bgt_y / rev_bgt_y, 4) if rev_bgt_y else ebi_act_y

for name, kind, cost, total in [
    ("Revenue",                    "value", False, False),
    ("Direct / operator costs",    "value", True,  False),
    ("Gross profit / contribution","value", False, True ),
    ("Operating expenses (staff, facilities, mktg, central)", "value", True, False),
    ("EBITDA",                     "value", False, True ),
]:
    e = _pl_entry(name, kind, cost, total)
    if e:
        pl.append(e)

# Rename for display
for r in pl:
    if r["name"] == "Gross profit / contribution":
        r["name"] = "Gross profit"
    if r["name"] == "Operating expenses (staff, facilities, mktg, central)":
        r["name"] = "Operating expenses"

pl.append({
    "name": "Gross profit margin", "kind": "pct", "cost": False, "total": False, "monthly": None,
    "month": {"actual": gm_act_m,  "budget": gm_bgt_m, "var": rnd(gm_act_m-gm_bgt_m,4),   "pct": None},
    "ytd":   {"actual": gm_act_y,  "budget": gm_bgt_y, "var": rnd(gm_act_y-gm_bgt_y,4),   "pct": None},
})
pl.append({
    "name": "Contribution margin", "kind": "pct", "cost": False, "total": False, "monthly": None,
    "month": {"actual": gm_act_m,  "budget": gm_bgt_m, "var": rnd(gm_act_m-gm_bgt_m,4),   "pct": None},
    "ytd":   {"actual": gm_act_y,  "budget": gm_bgt_y, "var": rnd(gm_act_y-gm_bgt_y,4),   "pct": None},
})
pl.append({
    "name": "EBITDA margin", "kind": "pct", "cost": False, "total": False, "monthly": None,
    "month": {"actual": ebi_act_m, "budget": ebi_bgt_pct_m, "var": rnd(ebi_act_m-ebi_bgt_pct_m,4), "pct": None},
    "ytd":   {"actual": ebi_act_y, "budget": ebi_bgt_pct_y, "var": rnd(ebi_act_y-ebi_bgt_pct_y,4), "pct": None},
})

# Revenue mix
rev_mix = []
for r in revenue:
    if r["total"]:
        continue
    tot_m = rev_row["month"]["actual"] or 1
    tot_y = rev_row["ytd"]["actual"]   or 1
    rev_mix.append({
        "name": r["name"],
        "month": {"rev": r["month"]["actual"], "share": rnd(r["month"]["actual"]/tot_m, 4)},
        "ytd":   {"rev": r["ytd"]["actual"],   "share": rnd(r["ytd"]["actual"]/tot_y,   4)},
    })

margins = {
    "gm_month":    gm_act_m,
    "gm_ytd":      gm_act_y,
    "ebitda_month":ebi_act_m,
    "ebitda_ytd":  ebi_act_y,
}

# ── Working capital monthly (rows 61-69) ─────────────────────────────────────
wc_monthly = []
ACTUAL_MONTHS = set(months)
r_n = 62
while True:
    month = v(r_n, 0)
    if not month or str(month).strip() == "" or str(month).startswith("Jan–"):
        break
    if not re.match(r"[A-Z][a-z]+-\d{2}", str(month)):
        break
    ar  = rnd(nz(v(r_n,1)),4); ap_v = rnd(nz(v(r_n,2)),4)
    def_v = rnd(nz(v(r_n,3)),4); other = rnd(nz(v(r_n,4)),4)
    nwc  = rnd(nz(v(r_n,5)),4); mov = rnd(nz(v(r_n,6)),4)
    m_abbr = str(month)[:3]
    forecast = m_abbr not in ACTUAL_MONTHS
    entry = {"month": str(month), "ar": ar, "ap": ap_v, "deferred": def_v,
             "other": other, "nwc": nwc, "movement": mov}
    if forecast:
        entry["forecast"] = True
    wc_monthly.append(entry)
    r_n += 1

# ── Working capital YTD (rows 74-78) ─────────────────────────────────────────
wc_ytd = []
r_n = 74
while True:
    item = v(r_n, 0)
    if not item or str(item).strip() == "":
        break
    if str(item).startswith("Cash impact"):
        break
    open_ = rnd(nz(v(r_n,1)),4); close = rnd(nz(v(r_n,2)),4)
    mov   = rnd(nz(v(r_n,3)),4); cash  = rnd(nz(v(r_n,4)),4)
    wc_ytd.append({"item": str(item), "open": open_, "close": close,
                   "movement": mov, "cash_impact": cash})
    r_n += 1

# ── AR monthly (rows 84-91) ───────────────────────────────────────────────────
ar_monthly = []
r_n = 84
while True:
    month = v(r_n, 0)
    if not month or not re.match(r"[A-Z][a-z]+-\d{2}", str(month)):
        break
    opening = rnd(nz(v(r_n,1)),4); inv = rnd(nz(v(r_n,2)),4)
    coll    = rnd(nz(v(r_n,3)),4); closing = rnd(nz(v(r_n,4)),4)
    rate    = rnd(nz(v(r_n,5)),4)
    ar_monthly.append({"month": str(month), "opening": opening, "invoiced": inv,
                       "collected": coll, "closing": closing, "rate": rate})
    r_n += 1

# AR YTD row
ar_ytd = {"invoiced": 0, "collected": 0, "closing": 0, "rate": 0}
for r_n2 in range(r_n, r_n+3):
    label = v(r_n2, 0)
    if label and "ytd" in str(label).lower():
        ar_ytd = {
            "invoiced":  rnd(nz(v(r_n2,2)),4),
            "collected": rnd(nz(v(r_n2,3)),4),
            "closing":   rnd(nz(v(r_n2,4)),4),
            "rate":      rnd(nz(v(r_n2,5)),4),
        }
        break

# AR aging (rows 96-99)
ar_aging = []
ar_aging_total = 0
r_n = 96
while True:
    bucket = v(r_n, 0)
    if not bucket or str(bucket).strip() == "":
        break
    amt = rnd(nz(v(r_n,1)),4); share = rnd(nz(v(r_n,2)),4)
    if "total" in str(bucket).lower():
        ar_aging_total = amt
        break
    ar_aging.append({"bucket": str(bucket), "amount": amt, "share": share})
    r_n += 1

# ── AP monthly (rows 105-112) ─────────────────────────────────────────────────
ap_monthly = []
r_n = 105
while True:
    month = v(r_n, 0)
    if not month or not re.match(r"[A-Z][a-z]+-\d{2}", str(month)):
        break
    opening = rnd(nz(v(r_n,1)),4); purch = rnd(nz(v(r_n,2)),4)
    pays    = rnd(nz(v(r_n,3)),4); closing = rnd(nz(v(r_n,4)),4)
    dpo     = rnd(nz(v(r_n,5)),4)
    ap_monthly.append({"month": str(month), "opening": opening, "purchases": purch,
                       "payments": pays, "closing": closing, "dpo": dpo})
    r_n += 1

# AP YTD
ap_ytd = {"purchases": 0, "payments": 0, "closing": 0, "dpo": 0}
for r_n2 in range(r_n, r_n+3):
    label = v(r_n2, 0)
    if label and "ytd" in str(label).lower():
        ap_ytd = {
            "purchases": rnd(nz(v(r_n2,2)),4),
            "payments":  rnd(nz(v(r_n2,3)),4),
            "closing":   rnd(nz(v(r_n2,4)),4),
            "dpo":       rnd(nz(v(r_n2,5)),4),
        }
        break

# AP aging (rows 118-121)
ap_aging = []
ap_aging_total = 0
r_n = 118
while True:
    bucket = v(r_n, 0)
    if not bucket or str(bucket).strip() == "":
        break
    amt = rnd(nz(v(r_n,1)),4); share = rnd(nz(v(r_n,2)),4)
    if "total" in str(bucket).lower():
        ap_aging_total = amt
        break
    ap_aging.append({"bucket": str(bucket), "amount": amt, "share": share})
    r_n += 1

# ── Cash tiles (rows 133-139) ─────────────────────────────────────────────────
cash_val_m   = rnd(nz(v(135,10)),4)
coll_val_m   = rnd(nz(v(135,15)),4)
nwc_val_m    = rnd(nz(v(138,10)),4)
ar_val_m     = rnd(nz(v(138,15)),4)
cash_note_m  = str(v(136,10) or "")
coll_note_m  = str(v(136,15) or "")
nwc_note_m   = str(v(139,10) or "")
ar_note_m    = str(v(139,15) or "")

cash_val_y   = rnd(nz(v(135,20)),4)
coll_val_y   = rnd(nz(v(135,25)),4)
nwc_val_y    = rnd(nz(v(138,20)),4)
ar_val_y     = rnd(nz(v(138,25)),4)
cash_note_y  = str(v(136,20) or "")
coll_note_y  = str(v(136,25) or "")
nwc_note_y   = str(v(139,20) or "")
ar_note_y    = str(v(139,25) or "")

cash_tiles = {
    "month": {
        "cash":        {"value": cash_val_m,  "note": cash_note_m},
        "collections": {"value": coll_val_m,  "note": coll_note_m},
        "nwc":         {"value": nwc_val_m,   "note": nwc_note_m},
        "ar":          {"value": ar_val_m,    "note": ar_note_m},
    },
    "ytd": {
        "cash":        {"value": cash_val_y,  "note": cash_note_y},
        "collections": {"value": coll_val_y,  "note": coll_note_y},
        "nwc":         {"value": nwc_val_y,   "note": nwc_note_y},
        "ar":          {"value": ar_val_y,    "note": ar_note_y},
    },
}

# ── Cash bridges ─────────────────────────────────────────────────────────────
def parse_bridge(start_row, end_row):
    steps = []
    for r_n in range(start_row, end_row+1):
        name = v(r_n, 0)
        if not name or str(name).strip() == "" or str(name).startswith("The "):
            continue
        col4 = nz(v(r_n,4)); col5 = nz(v(r_n,5)); col6 = nz(v(r_n,6))
        running = rnd(nz(v(r_n,2)),4)
        if col4 > 0:
            steps.append({"step": str(name), "kind": "total", "value": rnd(col4,4)})
        elif col5 > 0:
            steps.append({"step": str(name), "kind": "delta", "value": rnd(col5,4),  "running": running})
        elif col6 > 0:
            steps.append({"step": str(name), "kind": "delta", "value": rnd(-col6,4), "running": running})
    return steps

bridge_month = parse_bridge(143, 150)
bridge_ytd   = parse_bridge(154, 161)

# ── Runway (rows 166-169) ─────────────────────────────────────────────────────
burn_month = rnd(nz(v(166,1)),4)
burn_3m    = rnd(nz(v(167,1)),4)
ytd_move   = rnd(nz(v(168,1)),4)
runway_mo  = rnd(nz(v(169,1)),4)
runway = {"burn_month": burn_month, "burn_3m": burn_3m,
          "ytd_move": ytd_move, "months": runway_mo}

# ── Cash monthly time series ──────────────────────────────────────────────────
# Derive from AR closing balances mapped to month labels.
# Only actuals from ar_monthly closing cash is not directly available;
# use a simple linear interpolation from opening (FY start) to current close.
# Mark all as non-forecast for closed months.
ar_closed_months = {r["month"][:3]: r for r in ar_monthly}
cash_m_labels, cash_m_values, cash_m_forecast, cash_m_illustrative = [], [], [], []

open_cash = cash_val_y  # Opening cash = year-start cash tile
close_cash = cash_val_m
n_months = len(months)
for i, m in enumerate(months):
    label = f"{m}-{year[2:]}"
    cash_m_labels.append(label)
    # Linear interpolation as placeholder
    interp = rnd(open_cash + (close_cash - open_cash) * (i / max(n_months-1,1)), 3)
    cash_m_values.append(interp)
    cash_m_forecast.append(False)
    cash_m_illustrative.append(i < n_months - 1)  # only last month is confirmed

cash_monthly = {
    "labels":      cash_m_labels,
    "values":      cash_m_values,
    "forecast":    cash_m_forecast,
    "illustrative":cash_m_illustrative,
}

# ── Key updates (rows 174-179) ────────────────────────────────────────────────
key_updates = []
r_n = 174
while True:
    num   = v(r_n, 0)
    topic = v(r_n, 1)
    text  = v(r_n, 3)
    if not topic and not text:
        break
    if topic and text:
        key_updates.append({"topic": str(topic), "text": str(text)})
    r_n += 1
    if r_n > 185:
        break

# ── Placeholder contract / vendor breakdowns ──────────────────────────────────
# Not in this Excel tab — keep as empty lists (the [DUMMY] cards will show empty bars)
ar_by_contract = []
ap_by_vendor   = []

# ── Build final DATA object ───────────────────────────────────────────────────
DATA = {
    "meta":           meta,
    "revenue":        revenue,
    "costs":          costs,
    "pl":             pl,
    "margins":        margins,
    "rev_mix":        rev_mix,
    "wc_monthly":     wc_monthly,
    "wc_ytd":         wc_ytd,
    "ar_monthly":     ar_monthly,
    "ar_ytd":         ar_ytd,
    "ar_aging":       ar_aging,
    "ar_aging_total": ar_aging_total,
    "ar_by_contract": ar_by_contract,
    "ap_monthly":     ap_monthly,
    "ap_ytd":         ap_ytd,
    "ap_aging":       ap_aging,
    "ap_aging_total": ap_aging_total,
    "ap_by_vendor":   ap_by_vendor,
    "cash_tiles":     cash_tiles,
    "bridge_month":   bridge_month,
    "bridge_ytd":     bridge_ytd,
    "cash_monthly":   cash_monthly,
    "runway":         runway,
    "key_updates":    key_updates,
}

# ── Inject into HTML template ─────────────────────────────────────────────────
html = tpl_path.read_text(encoding="utf-8")

data_json = json.dumps(DATA, ensure_ascii=False, separators=(",", ":"))

# Replace existing DATA constant (the line: const DATA = {...};)
new_html, count = re.subn(
    r'const DATA\s*=\s*\{.*?\};',
    f'const DATA = {data_json};',
    html,
    count=1,
    flags=re.DOTALL,
)
if count == 0:
    sys.exit("Could not find 'const DATA = {...};' in the template. Make sure the template is the correct dashboard HTML.")

out_path.write_text(new_html, encoding="utf-8")

print(f"✓ DATA injected from {xl_path.name}")
print(f"  Period : {meta['period']}")
print(f"  YTD    : {meta['ytd_label']}")
print(f"  Revenue: ${rev_row.get('ytd',{}).get('actual',0):.2f}M YTD")
print(f"  EBITDA : {ebi_act_y*100:.1f}% margin YTD")
print(f"  Cash   : ${cash_val_m:.2f}M · {runway_mo:.1f} mo runway")
print(f"✓ Written → {out_path}")
print(f"\nOpen {out_path} in your browser to view the dashboard.")
