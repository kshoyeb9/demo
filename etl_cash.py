"""
etl_cash.py — Rebuilds the DAILY cash dashboard from the daily workbook.

Usage:
    python etl_cash.py --file noon_cash_daily.xlsx [--out cash.html]

Like the monthly ETL, this reads only the typed input cells and re-derives every
balance, total and ratio itself. The workbook's own formulas are for the person
reading it in Excel; nothing here depends on them having been calculated.
"""

import argparse, json, re, sys
from datetime import datetime, date, timedelta
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl not installed — run: pip install openpyxl")

ap = argparse.ArgumentParser()
ap.add_argument("--file", default="noon_cash_daily.xlsx")
ap.add_argument("--out",  default="cash.html")
ap.add_argument("--template", default=None, help="HTML template (defaults to --out)")
ap.add_argument("--publish-out", default="cash.publish.html")
args = ap.parse_args()

xl, out_path = Path(args.file), Path(args.out)
tpl = Path(args.template) if args.template else out_path
if not xl.exists():  sys.exit(f"Workbook not found: {xl}")
if not tpl.exists(): sys.exit(f"Page template not found: {tpl} — run make_cash_page.py first")

wb = openpyxl.load_workbook(xl, data_only=True)
SHEET = "Cash Input"
if SHEET not in wb.sheetnames:
    cand = [n for n in wb.sheetnames if n.strip().lower() != "instructions"]
    if len(cand) != 1: sys.exit(f"Sheet '{SHEET}' missing. Found: {', '.join(wb.sheetnames)}")
    SHEET = cand[0]
ws   = wb[SHEET]
rows = list(ws.iter_rows(values_only=True))

def find(header, col=0):
    want = str(header).strip().lower()
    for i, r in enumerate(rows):
        if col < len(r) and r[col] is not None and str(r[col]).strip().lower() == want:
            return i
    return None

def table(header, ncols, stop_labels=()):
    h = find(header)
    if h is None:
        sys.exit(f"Header '{header}' not found on '{SHEET}'. Do not rename a header row.")
    extra = {s.lower() for s in stop_labels}; out = []
    for r in rows[h+1:]:
        f = r[0] if r else None
        if f is None or str(f).strip() == "": break
        lab = str(f).strip().lower()
        if lab in extra or lab.startswith("total"): break
        out.append(tuple(r[i] if i < len(r) else None for i in range(ncols)))
    return out

def setting(label, default=None):
    want = str(label).strip().lower()
    for r in rows:
        if r and r[0] is not None and str(r[0]).strip().lower() == want:
            return r[1] if len(r) > 1 else default
    return default

def num(x, d=0.0):
    if isinstance(x, (int, float)): return float(x)
    if isinstance(x, str):
        try: return float(x.strip().replace("$","").replace(",",""))
        except ValueError: return d
    return d

def as_date(x):
    if isinstance(x, datetime): return x.date()
    if isinstance(x, date): return x
    if isinstance(x, str):
        try: return date.fromisoformat(x.strip()[:10])
        except ValueError: return None
    return None

def rnd(x, d=4): return None if x is None else round(float(x), d)

# ── setup ─────────────────────────────────────────────────────────────────────
asat    = as_date(setting("As-at date"))
opening = num(setting("Opening balance"))
currency= str(setting("Currency label", "USD M"))
floor   = num(setting("Runway floor"), 0.0)
if asat is None: sys.exit("'As-at date' in Section 1 is missing or not a date.")

# ── daily ledger ──────────────────────────────────────────────────────────────
led = []
for r in table("Ledger Date", 5):
    d = as_date(r[0])
    if d is None or d > asat:      # ignore anything typed beyond the as-at date
        continue
    led.append((d, num(r[1]), num(r[2])))
led.sort(key=lambda x: x[0])
if not led: sys.exit("The daily ledger is empty — add at least one row under 'Ledger Date'.")

bal, run = [], opening
net = [rec - pay for _, rec, pay in led]
for n in net:
    run += n; bal.append(run)
balance = rnd(bal[-1], 4)

labels = [d.strftime("%d %b") for d, _, _ in led]
# Show roughly ten x-labels; the rest render as blanks so the axis stays readable
step   = max(1, round(len(labels) / 10))
sparse = [l if (i % step == 0 or i == len(labels)-1) else "" for i, l in enumerate(labels)]

# ── weekly roll-up of actuals ─────────────────────────────────────────────────
wk = {}
for (d, rec, pay), n in zip(led, net):
    key = d - timedelta(days=d.weekday())
    a = wk.setdefault(key, [0.0, 0.0, 0.0])
    a[0] += rec; a[1] += pay; a[2] += n
weeks = sorted(wk)[-8:]
weekly = {"recent":   [w.strftime("%d %b") for w in weeks],
          "receipts": [rnd(wk[w][0], 3) for w in weeks],
          "payments": [rnd(wk[w][1], 3) for w in weeks],
          "net":      [rnd(wk[w][2], 3) for w in weeks]}

# ── 13-week forecast ──────────────────────────────────────────────────────────
FC_COLS = ["collections","other_receipts","payroll","suppliers","tax","debt","other_payments"]
forecast, proj = [], balance
for r in table("Week Commencing", 10):
    d = as_date(r[0])
    if d is None: continue
    vals = {k: num(r[1+i]) for i, k in enumerate(FC_COLS)}
    n = vals["collections"] + vals["other_receipts"] - sum(
        vals[k] for k in ("payroll","suppliers","tax","debt","other_payments"))
    proj += n
    forecast.append({"week": d.strftime("%d %b"), **{k: rnd(v,3) for k,v in vals.items()},
                     "net": rnd(n,3), "closing": rnd(proj,3)})

trough_w = min(forecast, key=lambda w: w["closing"]) if forecast else None

# ── receivables ───────────────────────────────────────────────────────────────
def share(header, key):
    src = [(str(r[0]).strip(), num(r[1])) for r in table(header, 3)]
    tot = rnd(sum(a for _, a in src), 4)
    return [{key: b, "amount": rnd(a,3), "share": rnd(a/tot) if tot else None}
            for b, a in src], tot
ar_aging, ar_total = share("AR Aging Bucket", "bucket")
ar_by_contract, _  = share("Contract", "contract")
over60 = next((b["share"] for b in ar_aging if "60+" in b["bucket"]), 0) or 0

notes = [{"topic": str(r[1]).strip(), "text": str(r[2]).strip()}
         for r in table("Note #", 3) if r[1] and r[2]]

# ── KPIs ──────────────────────────────────────────────────────────────────────
recent    = net[-20:]
burn_day  = sum(recent)/len(recent) if recent else 0.0
BUS_MONTH = 21                       # business days in an average month
runway_m  = rnd(balance / (-burn_day * BUS_MONTH), 2) if burn_day < 0 else None
wk_start  = asat - timedelta(days=asat.weekday())
net_wtd   = sum(n for (d,_,_), n in zip(led, net) if d >= wk_start)
coll_mtd  = sum(rec for d, rec, _ in led if d.year == asat.year and d.month == asat.month)

kpi = {"balance": balance, "burn_day": rnd(burn_day,4), "runway_months": runway_m,
       "net_wtd": rnd(net_wtd,4), "collections_mtd": rnd(coll_mtd,4),
       "ar_total": ar_total, "ar_over60_share": rnd(over60,4), "floor": rnd(floor,3),
       "trough": trough_w["closing"] if trough_w else balance,
       "trough_week": ("week of " + trough_w["week"]) if trough_w else None}

DATA = {
    "meta": {"asat": asat.strftime("%d %b %Y"), "generated": datetime.now().strftime("%d %b %Y"),
             "source": xl.name, "currency": currency,
             "ledger_from": led[0][0].strftime("%d %b %Y"),
             "forecast_from": forecast[0]["week"] if forecast else "—"},
    "kpi": kpi,
    "ledger": {"labels": labels, "sparse": sparse,
               "receipts": [rnd(r,3) for _,r,_ in led], "payments": [rnd(p,3) for _,_,p in led],
               "net": [rnd(n,3) for n in net], "balance": [rnd(b,3) for b in bal]},
    "weekly": weekly, "forecast": forecast,
    "ar_aging": ar_aging, "ar_aging_total": ar_total, "ar_by_contract": ar_by_contract,
    "notes": notes,
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
print(f"  As at     {DATA['meta']['asat']}  ·  {len(led)} business days from {DATA['meta']['ledger_from']}")
print(f"  Balance   ${balance:.3f}M  ·  daily burn ${burn_day:+.3f}M  ·  "
      f"runway {runway_m if runway_m is not None else '—'} mo")
print(f"  Forecast  {len(forecast)} weeks · low ${kpi['trough']:.3f}M "
      f"({kpi['trough_week'] or '—'}) · floor ${floor:.2f}M")
print(f"  AR        ${ar_total:.2f}M · {over60*100:.0f}% over 60 days")

gap = (asat - led[-1][0]).days
if gap > 0:  warn(f"the ledger's last row is {led[-1][0]:%d %b} but the as-at date is "
                  f"{asat:%d %b} — {gap} day(s) missing.")
if forecast and abs(forecast[0]["closing"] - (balance + forecast[0]["net"])) > 0.001:
    warn("the forecast does not open from the ledger's closing balance.")
if kpi["trough"] <= floor:
    warn(f"forecast dips to ${kpi['trough']:.3f}M in {kpi['trough_week']}, "
         f"at or below the ${floor:.2f}M floor.")
if not forecast: warn("no forecast weeks — the 13-week section will be empty.")
if burn_day >= 0: warn("cash is building over the last 20 days, so no runway is shown.")
