"""
etl_cash.py — Rebuilds the daily Cash Position dashboard straight from the
treasury team's own 13-week cash flow workbook.

    python etl_cash.py --file Weekly_Report_CF_2026_Final.xlsx

No separate input file: the workbook that is already updated daily IS the
source, so nothing is keyed twice.

Two things about that workbook drive how this reads it.

1.  Which rows to add up.  Category rows carry SUBTOTAL formulas over their
    detail rows, and 'Other' sits *inside* the Financing range, so adding up
    every labelled row would double-count. Only the category rows below are
    summed, and they reconcile exactly to the model's own Cash Inflows (row 11)
    and Cash Outflows (row 53).

2.  Where actuals stop.  'Daily CF' row 3 flags each day Actual or Forecast and
    is the authority here, because it is the granular sheet. The weekly sheet
    tags two further weeks lowercase "actual" even though the days inside them
    are flagged Forecast — and those weeks contain a large investor receipt that
    has not landed. Trusting the weekly flag would overstate cash materially, so
    the boundary is taken from the daily sheet and the disagreement is reported.
"""

import argparse, json, re, sys, warnings
from datetime import datetime, date, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")
try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl not installed — run: pip install openpyxl")

ap = argparse.ArgumentParser()
ap.add_argument("--file", default="cf_source.xlsx")
ap.add_argument("--out",  default="cash.html")
ap.add_argument("--template", default=None)
ap.add_argument("--publish-out", default="cash.publish.html")
ap.add_argument("--weeks", type=int, default=13, help="forecast weeks to show")
ap.add_argument("--floor", type=float, default=0.75, help="operating cash floor, USD M")
args = ap.parse_args()

xl, out_path = Path(args.file), Path(args.out)
tpl = Path(args.template) if args.template else out_path
if not xl.exists():  sys.exit(f"Workbook not found: {xl}")
if not tpl.exists(): sys.exit(f"Page template not found: {tpl} — run make_cash_page.py first")

wb = openpyxl.load_workbook(xl, data_only=True)
for need in ("Daily CF", "Weekly CF (USD)"):
    if need not in wb.sheetnames:
        sys.exit(f"Sheet '{need}' not in {xl.name}. Found: {', '.join(wb.sheetnames)}")
day, wk = wb["Daily CF"], wb["Weekly CF (USD)"]

R_TYPE, R_FROM, R_BEGIN, R_IN, R_OUT = 3, 6, 9, 11, 53
M = 1_000_000.0                       # the workbook is in units; the dashboard is in millions

# Category rows. These carry the SUBTOTAL formulas and together reconcile to
# rows 11 and 53 — summing the detail rows as well would double-count.
# (row, chart label, short label for the table, direction)
CATS = [
    (13, "Tracks",             "Tracks",  "in"),
    (18, "B2B",                "B2B",     "in"),
    (35, "Partnership",        "Partner", "in"),
    (37, "B2C",                "B2C",     "in"),
    (40, "Financing in",       "Fin in",  "in"),
    (48, "Grants & other",     "Grants",  "in"),
    (54, "Payroll & benefits", "Payroll", "out"),
    (74, "Taxes & government", "Tax",     "out"),
    (80, "AP payments",        "AP",      "out"),
    (87, "Financing out",      "Fin out", "out"),
]
# Inflow detail rows, for the by-source chart (blank rows and sub-headers omitted)
SRC_ROWS = list(range(14, 18)) + list(range(19, 35)) + [36] + list(range(38, 40)) \
         + list(range(41, 48)) + [49, 50, 52]

def val(sh, r, c):
    v = sh.cell(row=r, column=c).value
    return float(v) if isinstance(v, (int, float)) else 0.0

def as_date(x):
    if isinstance(x, datetime): return x.date()
    if isinstance(x, date): return x
    return None

def rnd(x, d=4): return None if x is None else round(float(x), d)

def periods(sh):
    out = []
    for c in range(3, sh.max_column + 1):
        d = as_date(sh.cell(row=R_FROM, column=c).value)
        if d is None: continue
        out.append((c, d, str(sh.cell(row=R_TYPE, column=c).value or "").strip().lower()))
    return out

# ── where do actuals stop?  the daily sheet decides ──────────────────────────
dcols = periods(day)
actual_days = [(c, d) for c, d, t in dcols if t == "actual"]
if not actual_days:
    sys.exit("No days flagged 'Actual' in row 3 of 'Daily CF'.")
asat_col, asat = actual_days[-1]
first_fc = next((d for c, d, t in dcols if t == "forecast" and d > asat), None)

# ── daily ledger ─────────────────────────────────────────────────────────────
led = [(d, val(day, R_IN, c)/M, val(day, R_OUT, c)/M) for c, d in actual_days]
led.sort(key=lambda x: x[0])
opening = val(day, R_BEGIN, actual_days[0][0]) / M
net = [r - p for _, r, p in led]
bal, run = [], opening
for n in net:
    run += n; bal.append(run)
balance = rnd(bal[-1], 4)

labels = [d.strftime("%d %b") for d, _, _ in led]
step   = max(1, round(len(labels)/10))
sparse = [l if (i % step == 0 or i == len(labels)-1) else "" for i, l in enumerate(labels)]

# ── weekly actuals, for the two bar charts ───────────────────────────────────
wcols = periods(wk)
wk_act = [(c, d) for c, d, _ in wcols if d <= asat and
          (abs(val(wk, R_IN, c)) + abs(val(wk, R_OUT, c))) > 0][-8:]
weekly = {"recent":   [d.strftime("%d %b") for _, d in wk_act],
          "receipts": [rnd(val(wk, R_IN,  c)/M, 3) for c, _ in wk_act],
          "payments": [rnd(val(wk, R_OUT, c)/M, 3) for c, _ in wk_act],
          "net":      [rnd((val(wk, R_IN, c)-val(wk, R_OUT, c))/M, 3) for c, _ in wk_act]}

# ── forecast: every week beginning after the last actual day ─────────────────
fc_cols = [(c, d) for c, d, _ in wcols if d > asat][:args.weeks]
if not fc_cols:
    sys.exit(f"No forecast weeks found after {asat:%d %b %Y} in 'Weekly CF (USD)'.")

# Only show categories that actually move over the horizon — carrying columns of
# zeros pushes Net and Closing, the two that matter, off the side of the table.
active = {r for r, _, _, _ in CATS
          if any(abs(val(wk, r, c)) > 0 for c, _ in
                 [(cc, dd) for cc, dd, _ in wcols if dd > asat][:args.weeks])}
cats_meta = [{"key": f"c{r}", "label": lab, "short": sh, "dir": dr}
             for r, lab, sh, dr in CATS if r in active]
forecast, proj = [], balance
for c, d in fc_cols:
    vals = {f"c{r}": rnd(val(wk, r, c)/M, 3) for r, _, _, _ in CATS if r in active}
    n = sum(val(wk, r, c) for r, _, _, dr in CATS if dr == "in") / M \
      - sum(val(wk, r, c) for r, _, _, dr in CATS if dr == "out") / M
    proj += n
    forecast.append({"week": d.strftime("%d %b"), "vals": vals,
                     "net": rnd(n, 3), "closing": rnd(proj, 3)})

fc_in  = sum(val(wk, r, c) for r, _, _, dr in CATS if dr == "in"  for c, _ in fc_cols) / M
fc_out = sum(val(wk, r, c) for r, _, _, dr in CATS if dr == "out" for c, _ in fc_cols) / M
trough = min(forecast, key=lambda w: w["closing"])

# reconcile the chosen category rows against the model's own totals
model_in  = sum(val(wk, R_IN,  c) for c, _ in fc_cols) / M
model_out = sum(val(wk, R_OUT, c) for c, _ in fc_cols) / M

# ── composition over the horizon ─────────────────────────────────────────────
def label_of(r):
    a = wk.cell(row=r, column=1).value
    b = wk.cell(row=r, column=2).value
    t = str(a).strip() if a else (str(b).strip() if b else f"Row {r}")
    return re.sub(r"\s+", " ", t)[:44]

srcs = [(label_of(r), sum(val(wk, r, c) for c, _ in fc_cols)/M) for r in SRC_ROWS]
srcs = sorted([(s, a) for s, a in srcs if a > 0], key=lambda x: -x[1])[:8]
inflow_sources = [{"source": s, "amount": rnd(a, 3)} for s, a in srcs]
outflow_cats = sorted(
    [{"cat": lab, "amount": rnd(sum(val(wk, r, c) for c, _ in fc_cols)/M, 3)}
     for r, lab, sh, dr in CATS if dr == "out"],
    key=lambda x: -x["amount"])

# ── KPIs ─────────────────────────────────────────────────────────────────────
recent   = net[-20:]
burn_day = sum(recent)/len(recent) if recent else 0.0
# Weeks of cover come from the forecast, not from extrapolating the daily burn:
# actual days are lumpy (a single payment run can dominate a 20-day window), so
# the extrapolation swings wildly. The forecast is what treasury actually steers on.
breach = next((i+1 for i, w in enumerate(forecast) if w["closing"] <= args.floor), None)
cover_weeks = breach
net_wtd  = sum(n for (d,_,_), n in zip(led, net) if d > asat - timedelta(days=7))
rec_mtd  = sum(r for d, r, _ in led if d.year == asat.year and d.month == asat.month)

notes = [
    {"topic": "Position", "text":
     f"Cash stood at ${balance:.2f}M on {asat:%d %B %Y}, the last day marked actual in the "
     f"daily cash flow. Everything from {first_fc:%d %B} onward is forecast."
     if first_fc else f"Cash stood at ${balance:.2f}M on {asat:%d %B %Y}."},
    {"topic": "Forecast", "text":
     f"The {len(forecast)}-week view expects ${fc_in:.2f}M in and ${fc_out:.2f}M out, "
     f"reaching its low of ${trough['closing']:.2f}M in the week of {trough['week']}."},
    {"topic": "Inflows", "text":
     f"{inflow_sources[0]['source']} is the largest expected receipt at "
     f"${inflow_sources[0]['amount']:.2f}M, {inflow_sources[0]['amount']/fc_in*100:.0f}% of "
     f"forecast inflows. The projection depends heavily on it arriving as timed."
     if inflow_sources and fc_in else "No inflows are forecast over the horizon."},
    {"topic": "Outflows", "text":
     f"{outflow_cats[0]['cat']} dominates at ${outflow_cats[0]['amount']:.2f}M, "
     f"{outflow_cats[0]['amount']/fc_out*100:.0f}% of forecast payments."
     if outflow_cats and fc_out else "No outflows are forecast over the horizon."},
]

DATA = {
    "meta": {"asat": asat.strftime("%d %b %Y"), "generated": datetime.now().strftime("%d %b %Y"),
             "source": xl.name, "currency": "USD M",
             "ledger_from": led[0][0].strftime("%d %b %Y"),
             "forecast_from": first_fc.strftime("%d %b %Y") if first_fc else fc_cols[0][1].strftime("%d %b %Y")},
    "kpi": {"balance": balance, "burn_day": rnd(burn_day,4), "cover_weeks": cover_weeks, "horizon": len(forecast),
            "net_wtd": rnd(net_wtd,4), "receipts_mtd": rnd(rec_mtd,4), "floor": rnd(args.floor,3),
            "trough": trough["closing"], "trough_week": "week of " + trough["week"],
            "forecast_in": rnd(fc_in,3), "forecast_out": rnd(fc_out,3)},
    "ledger": {"labels": labels, "sparse": sparse,
               "receipts": [rnd(r,3) for _,r,_ in led], "payments": [rnd(p,3) for _,_,p in led],
               "net": [rnd(n,3) for n in net], "balance": [rnd(b,3) for b in bal]},
    "weekly": weekly, "forecast": forecast, "cats": cats_meta,
    "inflow_sources": inflow_sources, "outflow_cats": outflow_cats, "notes": notes,
}

html = tpl.read_text(encoding="utf-8")
payload = json.dumps(DATA, ensure_ascii=False, separators=(",",":"))
new, n = re.subn(r"const DATA\s*=\s*\{.*?\};", lambda _m: f"const DATA = {payload};",
                 html, count=1, flags=re.DOTALL)
if n == 0: sys.exit("Could not find 'const DATA = {...};' in the page template.")
out_path.write_text(new, encoding="utf-8")
if args.publish_out:
    try:
        s = new.index("<body>")+len("<body>"); e = new.rindex("</body>")
        Path(args.publish_out).write_text(new[s:e].strip()+"\n", encoding="utf-8")
    except ValueError:
        Path(args.publish_out).write_text(new, encoding="utf-8")

def warn(m): print(f"  ! {m}")
print(f"✓ Rebuilt {out_path} from {xl.name}")
print(f"  As at     {asat:%d %b %Y} (last day flagged Actual in 'Daily CF')")
print(f"  Balance   ${balance:,.3f}M  ·  20-day burn ${burn_day:+,.4f}M/day")
print(f"  Cover     " + (f"{cover_weeks} week(s) before the floor is breached"
                         if cover_weeks else f"clears the floor for all {len(forecast)} weeks"))
print(f"  Forecast  {len(forecast)} weeks from {DATA['meta']['forecast_from']} · "
      f"in ${fc_in:,.2f}M / out ${fc_out:,.2f}M")
dropped = [lab for r, lab, _, _ in CATS if r not in active]
if dropped: print(f"  Omitted   {', '.join(dropped)} — nothing forecast over the horizon")
print(f"  Low point ${trough['closing']:,.2f}M in week of {trough['week']} "
      f"(floor ${args.floor:.2f}M)")

if abs(fc_in - model_in) > 0.01 or abs(fc_out - model_out) > 0.01:
    warn(f"category rows do not reconcile to the workbook's own totals: "
         f"in {fc_in:,.2f} vs {model_in:,.2f}, out {fc_out:,.2f} vs {model_out:,.2f}")
else:
    print(f"  Reconciled category rows tie to Cash Inflows / Cash Outflows exactly")

horizon_end = fc_cols[-1][1]
stale = [d for c, d, t in wcols if t == "actual" and asat < d <= horizon_end]
if stale:
    warn(f"'Weekly CF (USD)' flags {len(stale)} week(s) from {min(stale):%d %b} as actual, but the "
         f"days inside them are flagged Forecast in 'Daily CF'. Treated as forecast. Those weeks "
         f"carry the large investor receipt, so trusting the weekly flag would overstate cash.")
if trough["closing"] <= args.floor:
    warn(f"forecast falls to ${trough['closing']:,.2f}M in week of {trough['week']}, "
         f"at or below the ${args.floor:.2f}M floor.")
if fc_out > fc_in:
    warn(f"forecast outflows exceed inflows by ${fc_out-fc_in:,.2f}M over {len(forecast)} weeks.")
