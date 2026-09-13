"""
make_cash_template.py — Generates the DAILY cash input workbook.

Separate from the monthly P&L workbook on purpose: this one is touched every
day, so it stays small. Two tabs: 'Instructions' and 'Cash Input'.

  Section 1  Setup — as-at date and opening balance
  Section 2  Daily ledger — one row per business day: receipts, payments
  Section 3  13-week forecast — one row per week, by category
  Section 4  AR aging at the as-at date
  Section 5  AR by contract
  Section 6  Commentary

Everything on the dashboard is derived from these; nothing is typed twice.

Usage:
    python make_cash_template.py [--out noon_cash_daily.xlsx]
"""

import argparse
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ap = argparse.ArgumentParser()
ap.add_argument("--out", default="noon_cash_daily.xlsx")
ap.add_argument("--asat", default="2026-09-11", help="last actual day (YYYY-MM-DD)")
args = ap.parse_args()

ASAT = date.fromisoformat(args.asat)
# The daily ledger picks up where the monthly workbook's June close left off.
LEDGER_START = date(2026, 7, 1)
OPENING      = 3.78

FONT = "Verdana"
C_HEADER, C_SUBHDR = "11203A", "D9E1F2"
C_INPUT, C_CALC, C_TOTAL, C_WHITE = "FFF3CD", "F2F2F2", "BDD7EE", "FFFFFF"
BLUE, BLACK = "0000FF", "000000"
FMT_M   = '$#,##0.000;($#,##0.000);-'
FMT_MS  = '+$#,##0.000;-$#,##0.000;-'
FMT_PCT = '0.0%'
FMT_DT  = 'dd mmm yyyy'

def F(bold=False, color=BLACK, size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
def fill(h): return PatternFill("solid", fgColor=h)
def A(h="left", wrap=False, v="center"): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
THIN = Side(style="thin", color="BFBFBF")
BOX  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook(); wb.remove(wb.active)
WIDTH = 10

def bar(ws, r, text, width=WIDTH):
    c = ws.cell(row=r, column=1, value=text); c.font = F(bold=True, color=C_WHITE); c.alignment = A()
    for cc in range(1, width+1): ws.cell(row=r, column=cc).fill = fill(C_HEADER)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)
    ws.row_dimensions[r].height = 20

def note(ws, r, text, width=WIDTH):
    c = ws.cell(row=r, column=1, value=text)
    c.font = F(italic=True, color="595959"); c.alignment = A()
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)

def headers(ws, r, labels, merges=None):
    merges = merges or {}; col = 1
    for i, t in enumerate(labels):
        ws.cell(row=r, column=col, value=t); span = merges.get(i, 1)
        for cc in range(col, col+span):
            x = ws.cell(row=r, column=cc)
            x.font = F(bold=True); x.fill = fill(C_SUBHDR); x.border = BOX; x.alignment = A("center", wrap=True)
        if span > 1: ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+span-1)
        col += span
    ws.row_dimensions[r].height = 28

def put(ws, r, c, v, *, kind="input", fmt=None, bold=False, wrap=False, halign=None, span=1):
    bg = {"input": C_INPUT, "calc": C_CALC, "total": C_TOTAL}.get(kind, C_WHITE)
    for cc in range(c, c+span):
        x = ws.cell(row=r, column=cc); x.fill = fill(bg); x.border = BOX
    cell = ws.cell(row=r, column=c, value=v)
    isnum = isinstance(v, (int, float))
    cell.font = F(bold=bold or kind == "total", color=BLUE if kind == "input" and isnum else BLACK)
    if fmt: cell.number_format = fmt
    cell.alignment = A(halign or ("right" if isnum else "left"), wrap=wrap)
    if span > 1: ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+span-1)
    return cell


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INSTRUCTIONS
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Instructions")
for col, w in {"A": 4, "B": 28, "C": 94}.items(): ws.column_dimensions[col].width = w
ws["B2"] = "Noon Academy — Daily Cash Input"
ws["B2"].font = Font(name=FONT, bold=True, size=16, color=C_HEADER)
ws["B3"] = "Updated every business day. Separate from the monthly P&L workbook."
ws["B3"].font = F(italic=True, color="595959")

def ibar(r, t):
    c = ws.cell(row=r, column=2, value=t); c.font = F(bold=True, color=C_WHITE)
    for cc in (2,3): ws.cell(row=r, column=cc).fill = fill(C_HEADER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.row_dimensions[r].height = 20

r = 5; ibar(r, "  THE DAILY ROUTINE  (about two minutes)"); r += 2
for t, b in [
    ("1.  Add the day",
     "In Section 2, add one row: today's date, total receipts, total payments. The closing "
     "balance calculates and carries into tomorrow. Business days only — skip weekends."),
    ("2.  Move the as-at date",
     "In Section 1, set 'As-at date' to the day you just added. Everything on the dashboard "
     "reports to that date."),
    ("3.  Regenerate",
     "Run:      python etl_cash.py --file noon_cash_daily.xlsx"),
]:
    ws.cell(row=r, column=2, value=t).font = F(bold=True, size=11)
    c = ws.cell(row=r, column=3, value=b); c.font = F(); c.alignment = A(wrap=True, v="top")
    ws.row_dimensions[r].height = 44; r += 1

r += 1; ibar(r, "  THE WEEKLY ROUTINE"); r += 2
for t, b in [
    ("Roll the forecast",
     "Section 3 holds thirteen weeks of expected receipts and payments. Each Monday, delete the "
     "week just finished from the top and add a new week at the bottom, then revise the numbers "
     "in between. This is the part that tells you when cash gets tight — it is only as good as "
     "the last time it was revised."),
    ("Refresh receivables",
     "Update Sections 4 and 5 with the aging and the largest contract balances as they stand."),
]:
    ws.cell(row=r, column=2, value=t).font = F(bold=True, size=11)
    c = ws.cell(row=r, column=3, value=b); c.font = F(); c.alignment = A(wrap=True, v="top")
    ws.row_dimensions[r].height = 62; r += 1

r += 1; ibar(r, "  CONVENTIONS"); r += 1
for k, v in [
    ("Units",      "USD millions, to three decimals — enter 0.115 for $115k."),
    ("Currency",   "Always USD. The dashboard converts to SAR on display at 3.75."),
    ("Signs",      "Receipts and payments are both POSITIVE. The sign is applied for you."),
    ("Opening",    "Entered once, in Section 1. Every day after that chains from the day before."),
    ("Gaps",       "A blank row ends a table. Add each day at the bottom of the ledger."),
    ("Headers",    "Do not rename a header row — tables are found by their header text."),
]:
    put(ws, r, 2, k, kind="plain", bold=True); put(ws, r, 3, v, kind="plain", wrap=True)
    ws.row_dimensions[r].height = 26; r += 1
ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CASH INPUT
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Cash Input")
ws.column_dimensions["A"].width = 30
for c in "BCDEFGHIJ": ws.column_dimensions[c].width = 14

r = 1
bar(ws, r, "  SECTION 1 — SETUP"); r += 1
note(ws, r, "Set the as-at date each day. Opening balance is entered once and never changes."); r += 1
headers(ws, r, ["Setting", "Value", "Notes"], merges={2: 5}); r += 1
setting_row = {}
for k, v, n, fmt in [
    ("As-at date",       ASAT,         "The last day with actual figures", FMT_DT),
    ("Opening balance",  OPENING,      "Cash at the start of the ledger below", FMT_M),
    ("Currency label",   "USD M",      "Shown under chart titles", None),
    ("Runway floor",     0.75,         "Minimum operating cash — the forecast flags when it is breached", FMT_M),
]:
    setting_row[k] = r
    put(ws, r, 1, k, kind="plain", bold=True)
    put(ws, r, 2, v, kind="input", fmt=fmt)
    put(ws, r, 3, n, kind="plain", span=5)
    r += 1
OPEN_REF = f"$B${setting_row['Opening balance']}"
r += 1

# ── 2 · DAILY LEDGER ─────────────────────────────────────────────────────────
bar(ws, r, "  SECTION 2 — DAILY CASH LEDGER  (USD M)", width=6); r += 1
note(ws, r, "One row per business day. Enter receipts and payments; the rest calculates.", width=6); r += 1
headers(ws, r, ["Ledger Date", "Receipts", "Payments", "Net", "Closing balance"]); r += 1

# Dummy ledger: a mild net outflow, payroll on the last business day of the month
ledger, d, i = [], LEDGER_START, 0
while d <= ASAT:
    if d.weekday() < 5:
        nxt = d + timedelta(days=1)
        while nxt.weekday() >= 5: nxt += timedelta(days=1)
        month_end = nxt.month != d.month
        rec = round(0.095 + 0.030 * ((i * 7) % 5) / 4, 3)
        pay = round(0.088 + 0.025 * ((i * 3) % 4) / 3 + (0.62 if month_end else 0), 3)
        ledger.append((d, rec, pay))
        i += 1
    d += timedelta(days=1)

led_first = r
for dt, rec, pay in ledger:
    put(ws, r, 1, dt, kind="input", fmt=FMT_DT)
    put(ws, r, 2, rec, kind="input", fmt=FMT_M)
    put(ws, r, 3, pay, kind="input", fmt=FMT_M)
    put(ws, r, 4, f"=B{r}-C{r}", kind="calc", fmt=FMT_MS)
    prev = f"E{r-1}" if r > led_first else OPEN_REF
    put(ws, r, 5, f"={prev}+D{r}", kind="calc", fmt=FMT_M)
    r += 1
led_last = r-1
put(ws, r, 1, "Total", kind="total")
put(ws, r, 2, f"=SUM(B{led_first}:B{led_last})", kind="total", fmt=FMT_M)
put(ws, r, 3, f"=SUM(C{led_first}:C{led_last})", kind="total", fmt=FMT_M)
put(ws, r, 4, f"=SUM(D{led_first}:D{led_last})", kind="total", fmt=FMT_MS)
put(ws, r, 5, f"=E{led_last}", kind="total", fmt=FMT_M)
r += 2

# ── 3 · 13-WEEK FORECAST ─────────────────────────────────────────────────────
bar(ws, r, "  SECTION 3 — 13-WEEK CASH FORECAST  (USD M)", width=10); r += 1
note(ws, r, "Expected receipts and payments by week. Roll this forward every Monday.", width=10); r += 1
note(ws, r, "The opening balance of week 1 is the closing balance of the ledger above.", width=10); r += 1
headers(ws, r, ["Week Commencing", "Collections", "Other receipts", "Payroll",
                "Supplier payments", "Tax & government", "Debt service",
                "Other payments", "Net", "Projected closing"]); r += 1

wk = ASAT + timedelta(days=(7 - ASAT.weekday()) % 7 or 7)      # next Monday
fc_first = r
for w in range(13):
    wc = wk + timedelta(weeks=w)
    payroll = 0.62 if (wc + timedelta(days=4)).month != wc.month or w % 4 == 3 else 0.0
    put(ws, r, 1, wc, kind="input", fmt=FMT_DT)
    for c, v in [(2, round(0.52 + 0.04*((w*3) % 5)/4, 3)),   # collections
                 (3, 0.02),                                   # other receipts
                 (4, payroll),                                # payroll
                 (5, round(0.24 + 0.03*((w*2) % 4)/3, 3)),    # suppliers
                 (6, 0.09 if w % 4 == 1 else 0.0),            # tax
                 (7, 0.04),                                   # debt service
                 (8, round(0.03 + 0.01*(w % 3), 3))]:         # other
        put(ws, r, c, v, kind="input", fmt=FMT_M)
    put(ws, r, 9, f"=SUM(B{r}:C{r})-SUM(D{r}:H{r})", kind="calc", fmt=FMT_MS)
    prev = f"J{r-1}" if w else f"E{led_last}"
    put(ws, r, 10, f"={prev}+I{r}", kind="calc", fmt=FMT_M)
    r += 1
fc_last = r-1
put(ws, r, 1, "Total", kind="total")
for c in "BCDEFGH":
    put(ws, r, ord(c)-64, f"=SUM({c}{fc_first}:{c}{fc_last})", kind="total", fmt=FMT_M)
put(ws, r, 9,  f"=SUM(I{fc_first}:I{fc_last})", kind="total", fmt=FMT_MS)
put(ws, r, 10, f"=J{fc_last}", kind="total", fmt=FMT_M)
r += 2

# ── 4 & 5 · RECEIVABLES ──────────────────────────────────────────────────────
def share_table(r, title, note_text, hdr, data, total_label):
    bar(ws, r, title, width=3); r += 1
    note(ws, r, note_text, width=3); r += 1
    headers(ws, r, [hdr, "Amount", "Share"]); r += 1
    first = r
    for label, amt in data:
        put(ws, r, 1, label, kind="input"); put(ws, r, 2, amt, kind="input", fmt=FMT_M); r += 1
    last, tot = r-1, r
    for rr in range(first, last+1):
        put(ws, rr, 3, f'=IF($B${tot}=0,"",B{rr}/$B${tot})', kind="calc", fmt=FMT_PCT)
    put(ws, tot, 1, total_label, kind="total")
    put(ws, tot, 2, f"=SUM(B{first}:B{last})", kind="total", fmt=FMT_M)
    put(ws, tot, 3, 1.0, kind="total", fmt=FMT_PCT)
    return tot+2

r = share_table(r, "  SECTION 4 — RECEIVABLES AGING  (USD M)",
                "Balance at the as-at date, split by age.", "AR Aging Bucket",
                [("Current — due next 30 days", 2.95), ("31–60 days", 1.85),
                 ("60+ days / overdue", 2.05)], "Total AR")
r = share_table(r, "  SECTION 5 — RECEIVABLES BY CONTRACT  (USD M)",
                "Largest outstanding balances.", "Contract",
                [("MCIT", 2.20), ("Takaful", 1.45), ("Taalum", 1.15),
                 ("Ensan", 0.85), ("Tracks", 0.65), ("Other contracts", 0.55)], "Total")

# ── 6 · COMMENTARY ───────────────────────────────────────────────────────────
bar(ws, r, "  SECTION 6 — COMMENTARY"); r += 1
note(ws, r, "Short notes for the reader. Keep to what changed since yesterday."); r += 1
headers(ws, r, ["Note #", "Topic", "Comment"], merges={2: 8}); r += 1
for i, (topic, text) in enumerate([
    ("Position", "Cash stands at the level shown after a heavier-than-usual payment week covering quarterly tax and the September payroll run."),
    ("Collections", "MCIT milestone invoice remains the single largest open item. Collection would move the forecast trough out by roughly three weeks."),
    ("Forecast", "The thirteen-week view dips to its low point in week 8 on the back of the October payroll and the debt service instalment falling in the same week."),
    ("Watch", "Two contracts in the 60+ bucket account for most of the overdue balance and are the focus of collection effort this week."),
], 1):
    put(ws, r, 1, i, kind="plain", halign="center")
    put(ws, r, 2, topic, kind="input")
    put(ws, r, 3, text, kind="input", span=8, wrap=True)
    ws.row_dimensions[r].height = 32
    r += 1

ws.sheet_view.showGridLines = False
wb.save(args.out)
print(f"✓ Created {args.out}")
print(f"  As-at {ASAT:%d %b %Y} · {len(ledger)} ledger days · 13 forecast weeks")
