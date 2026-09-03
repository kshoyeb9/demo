"""
etl_pl.py — Rebuilds the Noon P&L dashboard from the input workbook.

Usage:
    python etl_pl.py --file noon_dashboard_input.xlsx [--out index.html]

All data lives on the 'Dashboard Input' tab. Tables are located by their HEADER
TEXT, not by row number, so you can insert or delete rows in the workbook freely. Only raw input cells are read; every derived
figure (variances, margins, shares, running balances) is computed here, so the
workbook's own formulas never need to have been recalculated.
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl not installed — run: pip install openpyxl")

ap = argparse.ArgumentParser()
ap.add_argument("--file", default="noon_dashboard_input.xlsx")
ap.add_argument("--out",  default="index.html")
ap.add_argument("--template", default=None,
                help="HTML template to inject into (defaults to --out)")
ap.add_argument("--publish-out", default="dashboard.publish.html",
                help="Also write a bare-content copy for publishing as an Artifact")
args = ap.parse_args()

xl_path  = Path(args.file)
out_path = Path(args.out)
tpl_path = Path(args.template) if args.template else out_path

if not xl_path.exists():
    sys.exit(f"Workbook not found: {xl_path}")
if not tpl_path.exists():
    sys.exit(f"HTML template not found: {tpl_path}")

wb = openpyxl.load_workbook(xl_path, data_only=True)

MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]

# ── Generic table reader ──────────────────────────────────────────────────────
DATA_SHEET = "Dashboard Input"

def sheet(name=DATA_SHEET):
    """All data lives on one tab; 'Instructions' is prose and is never read."""
    if name in wb.sheetnames:
        return wb[name]
    candidates = [n for n in wb.sheetnames if n.strip().lower() != "instructions"]
    if len(candidates) == 1:
        return wb[candidates[0]]
    sys.exit(f"Sheet '{name}' missing from {xl_path.name}. "
             f"Found: {', '.join(wb.sheetnames)}")

def grid(ws):
    return list(ws.iter_rows(values_only=True))

def find_header(rows, header_text, col=0):
    """Row index (0-based) of the row whose `col` cell equals header_text."""
    want = str(header_text).strip().lower()
    for i, r in enumerate(rows):
        if col < len(r) and r[col] is not None and str(r[col]).strip().lower() == want:
            return i
    return None

def read_table(ws, header_text, ncols, *, col=0, stop_on_total=True, stop_labels=()):
    """Rows under `header_text` until a blank first cell. Total rows are dropped
    (they are recomputed here). Returns a list of value-tuples of width ncols.

    `stop_labels` adds table-specific terminators. Keep these per-call, never
    global: a label that ends one table can be a legitimate data row in another
    (e.g. 'Net working capital' terminates the WC YTD table but is a cash tile)."""
    rows = grid(ws)
    h = find_header(rows, header_text, col)
    if h is None:
        sys.exit(f"Header '{header_text}' not found on sheet '{ws.title}'. "
                 f"Do not rename or delete a table's header row.")
    extra = {s.strip().lower() for s in stop_labels}
    out = []
    for r in rows[h + 1:]:
        first = r[col] if col < len(r) else None
        if first is None or str(first).strip() == "":
            break
        label = str(first).strip().lower()
        if label in extra:
            break
        if stop_on_total and label.startswith("total"):
            break
        out.append(tuple(r[i] if i < len(r) else None for i in range(col, col + ncols)))
    return out

def num(x, default=0.0):
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        s = x.strip().replace("$", "").replace(",", "").replace("%", "")
        try:
            return float(s)
        except ValueError:
            return default
    return default

def yes(x):
    return str(x).strip().lower() in ("y", "yes", "true", "1")

def rnd(x, d=4):
    return None if x is None else round(float(x), d)

def blk(actual, budget):
    """One actual-vs-budget cell block, with variance and % of budget."""
    a, b = float(actual or 0), float(budget or 0)
    return {"actual": rnd(a), "budget": rnd(b),
            "var": rnd(a - b), "pct": rnd(a / b) if b else None}


# ══ SETUP ═════════════════════════════════════════════════════════════════════
ws = sheet()
rows = grid(ws)

def setting(label, default=""):
    want = str(label).strip().lower()
    for r in rows:
        if r and r[0] is not None and str(r[0]).strip().lower() == want:
            return r[1] if len(r) > 1 and r[1] is not None else default
    return default

period_full = str(setting("Period label (month)", "Unknown"))
ytd_label   = str(setting("YTD label", ""))
year        = str(setting("Reporting year", datetime.now().year)).strip()
fiscal_year = str(setting("Fiscal year label", ""))
fy_start    = int(num(setting("Fiscal year start month", 7), 7))
currency    = str(setting("Currency label", "USD M"))

month_rows = read_table(ws, "Setup Month", 2)
months      = [str(r[1]).strip() for r in month_rows if r[1]]
months_full = [str(r[0]).strip() for r in month_rows if r[0]]
if not months:
    sys.exit("No months listed on the Setup tab. Add at least one row under 'Month'.")

meta = {
    "period":         period_full,
    "ytd_label":      ytd_label,
    "generated":      datetime.now().strftime("%d %b %Y"),
    "source":         xl_path.name,
    "currency":       currency,
    "months":         months,
    "year":           year,
    "period_index":   len(months) - 1,
    "fy_start_month": fy_start,
    "fiscal_year":    fiscal_year,
}

# ══ P&L MONTHLY  (read first — the P&L blocks attach these series) ════════════
ws_m = ws
n_months = len(months)

def monthly_block(header):
    """{line name: [values aligned to meta.months]}"""
    out = {}
    for r in read_table(ws_m, header, 1 + n_months):
        name = str(r[0]).strip()
        out[name] = [rnd(num(r[1 + i])) for i in range(n_months)]
    return out

rev_act_m  = monthly_block("Revenue — Actual")
rev_bgt_m  = monthly_block("Revenue — Budget")
cost_act_m = monthly_block("Costs — Actual")
cost_bgt_m = monthly_block("Costs — Budget")

def attach_monthly(name, act_map, bgt_map):
    a, b = act_map.get(name), bgt_map.get(name)
    if a is None or b is None:
        return None
    return {"actual": a, "budget": b}

def sum_series(maps, names):
    """Element-wise sum of several named series — used for group totals."""
    tot = [0.0] * n_months
    for nm in names:
        s = maps.get(nm)
        if s:
            for i, v in enumerate(s):
                tot[i] += v or 0
    return [rnd(v) for v in tot]


# ══ P&L TAB ═══════════════════════════════════════════════════════════════════
ws_p = ws

def avb_rows(header):
    return [(str(r[0]).strip(), num(r[1]), num(r[2]), num(r[3]), num(r[4]))
            for r in read_table(ws_p, header, 5) if r[0]]

rev_src  = avb_rows("Business Unit")
cost_src = avb_rows("Cost Category")

revenue = [{"name": n, "month": blk(am, bm), "ytd": blk(ay, by),
            "total": False, "monthly": attach_monthly(n, rev_act_m, rev_bgt_m)}
           for n, am, bm, ay, by in rev_src]
costs   = [{"name": n, "month": blk(am, bm), "ytd": blk(ay, by),
            "total": False, "monthly": attach_monthly(n, cost_act_m, cost_bgt_m)}
           for n, am, bm, ay, by in cost_src]

def add_total(items, label, act_map, bgt_map):
    am = sum(i["month"]["actual"] for i in items)
    bm = sum(i["month"]["budget"] for i in items)
    ay = sum(i["ytd"]["actual"]   for i in items)
    by = sum(i["ytd"]["budget"]   for i in items)
    names = [i["name"] for i in items]
    mon = None
    if all(i["monthly"] for i in items):
        mon = {"actual": sum_series(act_map, names), "budget": sum_series(bgt_map, names)}
    items.append({"name": label, "month": blk(am, bm), "ytd": blk(ay, by),
                  "total": True, "monthly": mon})
    return items[-1]

rev_total  = add_total(revenue, "Total Revenue", rev_act_m, rev_bgt_m)
cost_total = add_total(costs,   "Total costs",   cost_act_m, cost_bgt_m)

# P&L summary — read the labels/flags, recompute every value from the blocks above
pl_src = [(str(r[0]).strip(), str(r[5] or "value").strip(), str(r[6] or "No").strip())
          for r in read_table(ws_p, "P&L Line", 7) if r[0]]

cost_by_name = {c["name"]: c for c in costs if not c["total"]}

def pl_value_row(name, kind, is_cost):
    """Resolve one P&L line to its source block."""
    if name == "Revenue":
        return rev_total
    if name in cost_by_name:
        return cost_by_name[name]
    # 'Marketing expense' ↔ 'Marketing', etc.
    stripped = re.sub(r"\s+(expense|costs?)$", "", name, flags=re.I).strip()
    for cname, c in cost_by_name.items():
        if cname.lower() == stripped.lower() or cname.lower() == name.lower():
            return c
    return None

def combine(base, minus, label):
    """base − sum(minus), across month/ytd and the monthly series."""
    out = {"name": label, "kind": "value", "cost": False, "total": True}
    for period in ("month", "ytd"):
        a = base[period]["actual"] - sum(m[period]["actual"] for m in minus)
        b = base[period]["budget"] - sum(m[period]["budget"] for m in minus)
        out[period] = blk(a, b)
    if base.get("monthly") and all(m.get("monthly") for m in minus):
        out["monthly"] = {
            k: [rnd(base["monthly"][k][i] - sum(m["monthly"][k][i] for m in minus))
                for i in range(n_months)]
            for k in ("actual", "budget")
        }
    else:
        out["monthly"] = None
    return out

pl = []
resolved = {}
for name, kind, is_cost in pl_src:
    lname = name.lower()
    if lname == "gross profit":
        entry = combine(resolved["Revenue"], [resolved["Direct costs"]], name)
    elif lname == "contribution profit":
        entry = combine(resolved["Gross profit"], [resolved["Marketing expense"]], name)
    elif lname == "ebitda":
        opex = [resolved[n] for n in ("BU salaries", "Noon HQ", "Other operating expenses")
                if n in resolved]
        entry = combine(resolved["Contribution profit"], opex, name)
    else:
        srcblk = pl_value_row(name, kind, is_cost)
        if srcblk is None:
            sys.exit(f"P&L line '{name}' has no matching row in the revenue or cost blocks.")
        entry = {"name": name, "kind": "value",
                 "month": dict(srcblk["month"]), "ytd": dict(srcblk["ytd"]),
                 "monthly": srcblk.get("monthly"),
                 "cost": yes(is_cost), "total": False}
    resolved[name] = entry
    pl.append(entry)

# Margin rows, derived from the resolved lines
rev_e = resolved.get("Revenue", rev_total)
def margin_entry(label, numer):
    e = {"name": label, "kind": "pct", "cost": False, "total": False, "monthly": None}
    for period in ("month", "ytd"):
        da = rev_e[period]["actual"] or 1
        db = rev_e[period]["budget"] or 1
        a  = (numer[period]["actual"] or 0) / da
        b  = (numer[period]["budget"] or 0) / db
        e[period] = {"actual": rnd(a), "budget": rnd(b), "var": rnd(a - b), "pct": None}
    return e

margin_specs = [("Gross profit margin", "Gross profit"),
                ("Contribution margin", "Contribution profit"),
                ("EBITDA margin",       "EBITDA")]
for label, src_name in margin_specs:
    if src_name in resolved:
        entry = margin_entry(label, resolved[src_name])
        anchor = next((i for i, p in enumerate(pl) if p["name"] == src_name), None)
        pl.insert(anchor + 1 if anchor is not None else len(pl), entry)

def m_of(label, period, field="actual"):
    e = next((p for p in pl if p["name"] == label), None)
    return e[period][field] if e else None

margins = {
    "gm_month":     m_of("Gross profit margin", "month"),
    "gm_ytd":       m_of("Gross profit margin", "ytd"),
    "ebitda_month": m_of("EBITDA margin",       "month"),
    "ebitda_ytd":   m_of("EBITDA margin",       "ytd"),
}

rev_mix = []
tm = rev_total["month"]["actual"] or 1
ty = rev_total["ytd"]["actual"]   or 1
for r in revenue:
    if r["total"]:
        continue
    rev_mix.append({
        "name":  r["name"],
        "month": {"rev": r["month"]["actual"], "share": rnd(r["month"]["actual"] / tm)},
        "ytd":   {"rev": r["ytd"]["actual"],   "share": rnd(r["ytd"]["actual"]   / ty)},
    })


# ══ WORKING CAPITAL ═══════════════════════════════════════════════════════════
ws_w = ws
wc_monthly, prev_nwc = [], None
for r in read_table(ws_w, "WC Month", 8):
    ar_, ap_, dfr, oth = num(r[1]), num(r[2]), num(r[3]), num(r[4])
    nwc = ar_ + ap_ + dfr + oth
    entry = {"month": str(r[0]).strip(), "ar": rnd(ar_), "ap": rnd(ap_),
             "deferred": rnd(dfr), "other": rnd(oth), "nwc": rnd(nwc),
             "movement": rnd(0 if prev_nwc is None else nwc - prev_nwc)}
    if yes(r[7]):
        entry["forecast"] = True
    prev_nwc = nwc
    wc_monthly.append(entry)

wc_ytd = []
for r in read_table(ws_w, "WC Item", 5, stop_labels=("Net working capital",)):
    op, cl = num(r[1]), num(r[2])
    wc_ytd.append({"item": str(r[0]).strip(), "open": rnd(op), "close": rnd(cl),
                   "movement": rnd(cl - op), "cash_impact": rnd(-(cl - op))})
if wc_ytd:
    o = sum(x["open"] for x in wc_ytd); c = sum(x["close"] for x in wc_ytd)
    wc_ytd.append({"item": "Net working capital", "open": rnd(o), "close": rnd(c),
                   "movement": rnd(c - o), "cash_impact": rnd(-(c - o))})


# ══ AR ════════════════════════════════════════════════════════════════════════
ws_ar = ws
ar_monthly = []
for r in read_table(ws_ar, "AR Month", 6):
    op, inv, coll = num(r[1]), num(r[2]), num(r[3])
    ar_monthly.append({"month": str(r[0]).strip(), "opening": rnd(op),
                       "invoiced": rnd(inv), "collected": rnd(coll),
                       "closing": rnd(op + inv - coll),
                       "rate": rnd(coll / inv) if inv else None})

actual_ar = [x for x in ar_monthly if x["month"][:3] in months]
inv_y  = sum(x["invoiced"]  for x in actual_ar)
coll_y = sum(x["collected"] for x in actual_ar)
ar_ytd = {"invoiced": rnd(inv_y), "collected": rnd(coll_y),
          "closing": actual_ar[-1]["closing"] if actual_ar else 0,
          "rate": rnd(coll_y / inv_y) if inv_y else None}

ar_aging_src   = [(str(r[0]).strip(), num(r[1])) for r in read_table(ws_ar, "AR Aging Bucket", 3)]
ar_aging_total = rnd(sum(a for _, a in ar_aging_src))
ar_aging = [{"bucket": b, "amount": rnd(a),
             "share": rnd(a / ar_aging_total) if ar_aging_total else None}
            for b, a in ar_aging_src]

ar_by_contract = [{"contract": str(r[0]).strip(), "amount": rnd(num(r[1]))}
                  for r in read_table(ws_ar, "Contract", 3)]


# ══ AP ════════════════════════════════════════════════════════════════════════
ws_ap = ws
ap_monthly = []
for r in read_table(ws_ap, "AP Month", 6):
    op, pur, pay = num(r[1]), num(r[2]), num(r[3])
    ap_monthly.append({"month": str(r[0]).strip(), "opening": rnd(op),
                       "purchases": rnd(pur), "payments": rnd(pay),
                       "closing": rnd(op + pur - pay), "dpo": rnd(num(r[5]))})

actual_ap = [x for x in ap_monthly if x["month"][:3] in months]
ap_ytd = {
    "purchases": rnd(sum(x["purchases"] for x in actual_ap)),
    "payments":  rnd(sum(x["payments"]  for x in actual_ap)),
    "closing":   actual_ap[-1]["closing"] if actual_ap else 0,
    "dpo":       rnd(sum(x["dpo"] for x in actual_ap) / len(actual_ap)) if actual_ap else 0,
}

ap_aging_src   = [(str(r[0]).strip(), num(r[1])) for r in read_table(ws_ap, "AP Aging Bucket", 3)]
ap_aging_total = rnd(sum(a for _, a in ap_aging_src))
ap_aging = [{"bucket": b, "amount": rnd(a),
             "share": rnd(a / ap_aging_total) if ap_aging_total else None}
            for b, a in ap_aging_src]

ap_by_vendor = [{"vendor": str(r[0]).strip(), "amount": rnd(num(r[1]))}
                for r in read_table(ws_ap, "Vendor", 3)]


# ══ CASH ══════════════════════════════════════════════════════════════════════
ws_c = ws

TILE_KEY = {"cash balance": "cash", "collections": "collections",
            "net working capital": "nwc", "accounts receivable": "ar"}
cash_tiles = {"month": {}, "ytd": {}}
for r in read_table(ws_c, "Cash Tile", 8):
    key = TILE_KEY.get(str(r[0]).strip().lower())
    if not key:
        continue
    # Note cells are merged across three columns, so YTD starts at index 5
    cash_tiles["month"][key] = {"value": rnd(num(r[1])), "note": str(r[2] or "")}
    cash_tiles["ytd"][key]   = {"value": rnd(num(r[5])), "note": str(r[6] or "")}

missing_tiles = set(TILE_KEY.values()) - set(cash_tiles["month"])
if missing_tiles:
    sys.exit(f"Cash tiles missing: {', '.join(sorted(missing_tiles))}. "
             f"The 'Cash Tile' table needs one row per tile, with no blank rows "
             f"between them: {', '.join(TILE_KEY)}.")

def read_bridge(header):
    """Opening/Inflow/Outflow/Closing rows → waterfall steps with running balances."""
    steps, running = [], 0.0
    for r in read_table(ws_c, header, 4):
        name = str(r[0]).strip()
        kind = str(r[1] or "").strip().lower()
        amt  = num(r[2])
        if kind == "opening":
            running = amt
            steps.append({"step": name, "kind": "total", "value": rnd(amt)})
        elif kind == "closing":
            steps.append({"step": name, "kind": "total", "value": rnd(running)})
        elif kind == "inflow":
            running += amt
            steps.append({"step": name, "kind": "delta", "value": rnd(amt),  "running": rnd(running)})
        elif kind == "outflow":
            running -= amt
            steps.append({"step": name, "kind": "delta", "value": rnd(-amt), "running": rnd(running)})
        else:
            sys.exit(f"Bridge row '{name}' has Type '{r[1]}'. "
                     f"Use one of: Opening, Inflow, Outflow, Closing.")
    return steps

bridge_month = read_bridge("Bridge Step (Month)")
bridge_ytd   = read_bridge("Bridge Step (YTD)")

cm_rows = read_table(ws_c, "Cash Month", 4)
cash_monthly = {
    "labels":       [str(r[0]).strip() for r in cm_rows],
    "values":       [rnd(num(r[1]), 3) for r in cm_rows],
    "forecast":     [yes(r[2]) for r in cm_rows],
    "illustrative": [yes(r[3]) for r in cm_rows],
}

RW_KEY = {"monthly cash burn": "burn_month",
          "trailing 3-month average burn": "burn_3m",
          "ytd net cash movement": "ytd_move",
          "cash runway (months)": "months"}
runway = {}
for r in read_table(ws_c, "Runway Metric", 3):
    k = RW_KEY.get(str(r[0]).strip().lower())
    if k:
        runway[k] = rnd(num(r[1]))
runway.setdefault("burn_month", 0); runway.setdefault("burn_3m", 0)
runway.setdefault("ytd_move", 0);   runway.setdefault("months", 0)


# ══ KEY UPDATES ═══════════════════════════════════════════════════════════════
ws_k = ws
key_updates = [{"topic": str(r[1]).strip(), "text": str(r[2]).strip()}
               for r in read_table(ws_k, "Update #", 3) if r[1] and r[2]]


# ══ ASSEMBLE & INJECT ═════════════════════════════════════════════════════════
DATA = {
    "meta": meta, "revenue": revenue, "costs": costs, "pl": pl,
    "margins": margins, "rev_mix": rev_mix,
    "wc_monthly": wc_monthly, "wc_ytd": wc_ytd,
    "ar_monthly": ar_monthly, "ar_ytd": ar_ytd,
    "ar_aging": ar_aging, "ar_aging_total": ar_aging_total,
    "ar_by_contract": ar_by_contract,
    "ap_monthly": ap_monthly, "ap_ytd": ap_ytd,
    "ap_aging": ap_aging, "ap_aging_total": ap_aging_total,
    "ap_by_vendor": ap_by_vendor,
    "cash_tiles": cash_tiles,
    "bridge_month": bridge_month, "bridge_ytd": bridge_ytd,
    "cash_monthly": cash_monthly, "runway": runway,
    "key_updates": key_updates,
}

html = tpl_path.read_text(encoding="utf-8")
payload = json.dumps(DATA, ensure_ascii=False, separators=(",", ":"))
new_html, n = re.subn(r"const DATA\s*=\s*\{.*?\};",
                      lambda _m: f"const DATA = {payload};",
                      html, count=1, flags=re.DOTALL)
if n == 0:
    sys.exit("Could not find 'const DATA = {...};' in the template.")
out_path.write_text(new_html, encoding="utf-8")

# Artifact publishing wraps page content in its own document shell, so the
# published copy must not carry a doctype/html/head/body of its own. Strip the
# outer document, keeping everything between <body> and </body>.
publish_path = Path(args.publish_out) if args.publish_out else None
if publish_path:
    try:
        start = new_html.index("<body>") + len("<body>")
        end   = new_html.rindex("</body>")
        publish_path.write_text(new_html[start:end].strip() + "\n", encoding="utf-8")
    except ValueError:
        publish_path = None   # template was already bare content; nothing to strip

# ── Report ────────────────────────────────────────────────────────────────────
def warn(msg):
    print(f"  ! {msg}")

print(f"✓ Rebuilt {out_path} from {xl_path.name}")
print(f"  Period    {meta['period']}   ·   YTD {meta['ytd_label']}   ·   {n_months} months")
print(f"  Revenue   ${rev_total['ytd']['actual']:.2f}M YTD "
      f"({rev_total['ytd']['pct']*100:.1f}% of budget)" if rev_total['ytd']['pct'] else "")
print(f"  EBITDA    {(margins['ebitda_ytd'] or 0)*100:.1f}% margin YTD")
print(f"  Cash      ${(bridge_month[-1]['value'] if bridge_month else 0):.2f}M "
      f"· {runway['months']:.1f} mo runway")
print(f"  Tables    {len(revenue)-1} BUs · {len(costs)-1} cost lines · {len(pl)} P&L rows · "
      f"{len(ar_by_contract)} contracts · {len(ap_by_vendor)} vendors · "
      f"{len(cash_monthly['labels'])} cash months · {len(key_updates)} updates")

# Consistency checks the reader would otherwise have to eyeball
for r in revenue + costs:
    if r.get("monthly"):
        s = sum(r["monthly"]["actual"])
        if abs(s - r["ytd"]["actual"]) > 0.05:
            warn(f"'{r['name']}' monthly actuals sum to {s:.2f} "
                 f"but its YTD actual is {r['ytd']['actual']:.2f}")
if not ar_by_contract:
    warn("AR by contract is empty — that chart will render blank.")
if not ap_by_vendor:
    warn("AP by vendor is empty — that chart will render blank.")
if not cash_monthly["labels"]:
    warn("Monthly closing cash is empty — the cash line chart will render blank.")
