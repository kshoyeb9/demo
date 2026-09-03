"""
make_template.py — Generates noon_dashboard_input.xlsx, the monthly input workbook
for the Noon P&L Performance dashboard.

Two tabs: 'Instructions' and 'Dashboard Input'. Every chart in the dashboard has
a dedicated table on the input tab. Tables are located by their HEADER TEXT, not
by row number, so inserting or deleting rows is safe.

Usage:
    python make_template.py [--out noon_dashboard_input.xlsx]
"""

import argparse
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ap = argparse.ArgumentParser()
ap.add_argument("--out", default="noon_dashboard_input.xlsx")
args = ap.parse_args()

SHEET = "Dashboard Input"

# ── Style tokens ──────────────────────────────────────────────────────────────
FONT      = "Arial"
C_HEADER  = "11203A"   # Noon dark blue — section bars
C_SUBHDR  = "D9E1F2"   # column header fill
C_INPUT   = "FFF3CD"   # yellow — cells the user edits
C_CALC    = "F2F2F2"   # grey — derived
C_TOTAL   = "BDD7EE"   # blue — totals
C_WHITE   = "FFFFFF"
BLUE_TXT  = "0000FF"
BLACK_TXT = "000000"

FMT_M   = '$#,##0.00;($#,##0.00);-'
FMT_MS  = '+$#,##0.00;-$#,##0.00;-'
FMT_PCT = '0.0%'
FMT_NUM = '#,##0.0'

def F(bold=False, color=BLACK_TXT, size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
def fill(hex_):
    return PatternFill("solid", fgColor=hex_)
def A(h="left", wrap=False, v="center"):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

THIN = Side(style="thin", color="BFBFBF")
BOX  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.remove(wb.active)

# Column layout for the single data sheet:
#   A        names / labels          (wide)
#   B..N     values, month columns   (uniform)
LAST_COL = 14  # N

def section_bar(ws, row, text, width=LAST_COL):
    c = ws.cell(row=row, column=1, value=text)
    c.font = F(bold=True, color=C_WHITE, size=11)
    c.fill = fill(C_HEADER)
    c.alignment = A("left")
    for cc in range(1, width + 1):
        ws.cell(row=row, column=cc).fill = fill(C_HEADER)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    ws.row_dimensions[row].height = 20

def note(ws, row, text, width=LAST_COL):
    c = ws.cell(row=row, column=1, value=text)
    c.font = F(size=9, italic=True, color="595959")
    c.alignment = A("left")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    return c

def headers(ws, row, labels, start=1, merges=None):
    """merges: {label_index: span} to widen a text column."""
    merges = merges or {}
    col = start
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col, value=t)
        c.font      = F(bold=True, size=10)
        c.alignment = A("center", wrap=True)
        span = merges.get(i, 1)
        for cc in range(col, col + span):
            ws.cell(row=row, column=cc).fill   = fill(C_SUBHDR)
            ws.cell(row=row, column=cc).border = BOX
        if span > 1:
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
        col += span
    ws.row_dimensions[row].height = 28

def put(ws, row, col, val, *, kind="input", fmt=None, bold=False,
        wrap=False, halign=None, span=1):
    """kind: input (yellow) | calc (grey) | total (blue) | plain (white)"""
    bg = {"input": C_INPUT, "calc": C_CALC, "total": C_TOTAL}.get(kind, C_WHITE)
    # style the whole span first — merged cells reject styling afterwards
    for cc in range(col, col + span):
        cell = ws.cell(row=row, column=cc)
        cell.fill   = fill(bg)
        cell.border = BOX
    c = ws.cell(row=row, column=col, value=val)
    is_num = isinstance(val, (int, float))
    c.font = F(bold=bold or kind == "total",
               color=BLUE_TXT if kind == "input" and is_num else BLACK_TXT)
    if fmt:
        c.number_format = fmt
    c.alignment = A(halign or ("right" if is_num else "left"), wrap=wrap)
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    return c


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INSTRUCTIONS
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Instructions")
for col, w in {"A": 4, "B": 30, "C": 88, "D": 34}.items():
    ws.column_dimensions[col].width = w

ws["B2"] = "Noon Academy — Dashboard Input Workbook"
ws["B2"].font = Font(name=FONT, bold=True, size=16, color=C_HEADER)
ws["B3"] = "Fill in the 'Dashboard Input' tab each month, then regenerate the dashboard."
ws["B3"].font = F(size=10, italic=True, color="595959")

def ibar(row, text):
    c = ws.cell(row=row, column=2, value=text)
    c.font = F(bold=True, color=C_WHITE, size=11)
    for cc in (2, 3, 4):
        ws.cell(row=row, column=cc).fill = fill(C_HEADER)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 20

r = 5
ibar(r, "  HOW TO USE THIS WORKBOOK"); r += 2
for title, body in [
    ("1.  Update the data",
     "Everything lives on the 'Dashboard Input' tab, in numbered sections from top to bottom. "
     "Every cell shaded YELLOW is an input — type over it. Grey cells are derived and white cells "
     "are labels; leave both alone."),
    ("2.  Keep the headers",
     "Add or remove rows inside a table freely (a new business unit, another vendor). "
     "Do NOT rename or delete a table's header row — the script finds each table by its header "
     "text, not by row number. A blank row marks the end of a table."),
    ("3.  Regenerate",
     "Save and close the file, then run:      python etl_pl.py --file noon_dashboard_input.xlsx"),
    ("4.  Open the dashboard",
     "The script rewrites index.html and prints a summary. Open the file in any browser — "
     "no server, and no Python needed to view it."),
]:
    ws.cell(row=r, column=2, value=title).font = F(bold=True, size=11)
    c = ws.cell(row=r, column=3, value=body); c.font = F(size=10); c.alignment = A(wrap=True, v="top")
    ws.row_dimensions[r].height = 40
    r += 1

r += 1
ibar(r, "  SECTIONS ON THE 'DASHBOARD INPUT' TAB"); r += 1
hdr = ["", "Section", "What it contains", "Dashboard element it drives"]
for i, t in enumerate(hdr):
    if i == 0: continue
    c = ws.cell(row=r, column=1 + i, value=t)
    c.font = F(bold=True); c.fill = fill(C_SUBHDR); c.border = BOX; c.alignment = A("center")
r += 1

for name, contains, drives in [
    ("1 · Setup",            "Period labels, reporting year, and the list of closed months.",
                             "Every header; the period selector"),
    ("2 · Revenue by BU",    "Actual vs budget by business unit, month and YTD.",
                             "Revenue bullet chart; revenue table"),
    ("3 · Costs",            "Actual vs budget by cost category, month and YTD.",
                             "Cost bullet chart; cost table"),
    ("4 · P&L summary",      "The P&L waterfall lines and their margins.",
                             "P&L summary table"),
    ("5 · Monthly series",   "Month-by-month actual and budget for every revenue and cost line.",
                             "Period-range selector (quarters, custom ranges)"),
    ("6 · Working capital",  "Monthly WC position and the YTD movement bridge.",
                             "Working capital table"),
    ("7 · Accounts receivable", "Monthly AR flow, aging buckets, balance by contract.",
                             "AR column chart; AR aging bars; AR by contract bars"),
    ("8 · Accounts payable", "Monthly AP flow, aging buckets, balance by vendor.",
                             "AP column chart; AP aging bars; AP by vendor bars"),
    ("9 · Cash",             "KPI tiles, both cash bridges, monthly closing cash, runway metrics.",
                             "Cash tiles; both waterfalls; cash balance line chart"),
    ("10 · Key updates",     "The narrative commentary paragraphs.",
                             "Key updates section at the bottom"),
]:
    put(ws, r, 2, name, kind="plain", bold=True)
    put(ws, r, 3, contains, kind="plain", wrap=True)
    put(ws, r, 4, drives,   kind="plain", wrap=True)
    ws.row_dimensions[r].height = 30
    r += 1

r += 1
ibar(r, "  CELL COLOUR LEGEND"); r += 1
for hexc, name, mean in [
    (C_INPUT, "Yellow", "An input. Type your number or text here."),
    (C_CALC,  "Grey",   "Derived from your inputs. Do not edit."),
    (C_TOTAL, "Blue",   "A total, recalculated from the rows above it."),
    (C_WHITE, "White",  "A label or heading."),
]:
    c = ws.cell(row=r, column=2, value=name)
    c.fill = fill(hexc); c.font = F(bold=True); c.border = BOX
    put(ws, r, 3, mean, kind="plain", wrap=True)
    r += 1

r += 1
ibar(r, "  CONVENTIONS"); r += 1
for k, vtext in [
    ("Units",           "All money is USD millions. Enter 2.5 for $2.5M — not 2500000."),
    ("Currency",        "Always enter USD here, whatever the source ledger. The dashboard has a "
                        "USD/SAR toggle that converts on display at 1 USD = 3.75 SAR; entering SAR "
                        "in this workbook would double-convert it."),
    ("Percentages",     "Enter as a percentage-formatted number (96.0%), not 0.96 or 96."),
    ("Costs",           "Enter costs as POSITIVE numbers. The dashboard applies the sign."),
    ("Working capital", "Payables and deferred revenue are entered NEGATIVE — they are liabilities."),
    ("Months",          "Use the 'Mmm-YY' format exactly: Jan-26, Feb-26. The script matches on this."),
    ("Forecast",        "Mark a monthly row 'Yes' under Forecast to grey it out in the dashboard."),
    ("Cash bridges",    "Type must be Opening, Inflow, Outflow or Closing. Enter every amount positive; "
                        "the running balance is computed for you."),
    ("Blank rows",      "A blank row ends a table. Never leave a gap in the middle of one."),
    ("Reconciliation",  "Each monthly series must sum across the months to that line's YTD figure. "
                        "The script warns you if it does not."),
]:
    put(ws, r, 2, k, kind="plain", bold=True)
    put(ws, r, 3, vtext, kind="plain", wrap=True)
    ws.row_dimensions[r].height = 28
    r += 1

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — DASHBOARD INPUT  (everything else)
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet(SHEET)
ws.column_dimensions["A"].width = 44
for i in range(2, LAST_COL + 1):
    ws.column_dimensions[get_column_letter(i)].width = 14

MONTH_COLS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
AVB = ["Actual (Month)", "Budget (Month)", "Actual YTD", "Budget YTD"]

r = 1

# ── 1 · SETUP ─────────────────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 1 — SETUP"); r += 1
note(ws, r, "These labels appear in every dashboard header. Update them first each month."); r += 1
headers(ws, r, ["Setting", "Value", "Notes"], merges={2: 4}); r += 1
for k, val, nt in [
    ("Period label (month)",     "June 2026",    "The latest closed month, e.g. 'July 2026'"),
    ("YTD label",                "Jan–Jun 2026", "Range covered year-to-date"),
    ("Reporting year",           "2026",         "Calendar year of the current period"),
    ("Fiscal year label",        "FY2025/26",    "Shown next to quarter labels"),
    ("Fiscal year start month",  7,              "1 = January … 7 = July"),
    ("Currency label",           "USD M",        "Displayed under chart titles"),
]:
    put(ws, r, 1, k, kind="plain", bold=True)
    put(ws, r, 2, val, kind="input")
    put(ws, r, 3, nt, kind="plain", span=4)
    r += 1
r += 1

note(ws, r, "One row per closed month, oldest first. This sets the period selector and the order of every monthly chart."); r += 1
headers(ws, r, ["Setup Month", "Short label"]); r += 1
for full, short in [("Jan-26","Jan"),("Feb-26","Feb"),("Mar-26","Mar"),
                    ("Apr-26","Apr"),("May-26","May"),("Jun-26","Jun")]:
    put(ws, r, 1, full,  kind="input")
    put(ws, r, 2, short, kind="input")
    r += 1
r += 2

# ── 2 & 3 · REVENUE AND COSTS ─────────────────────────────────────────────────
REVENUE = [
    ("Tracks",                0.99, 1.07,  5.65,  6.17),
    ("B2B",                   0.63, 0.63,  3.64,  3.72),
    ("Govt Schools — Legacy",  0.52, 0.56,  2.97,  3.20),
    ("Govt Schools — New",     0.26, 0.29,  1.41,  1.56),
    ("Out of School",         0.10, 0.10,  0.68,  0.60),
]
COSTS = [
    ("Direct costs",             0.89, 0.92, 5.06, 5.28),
    ("Marketing",                0.10, 0.13, 0.67, 0.82),
    ("BU salaries",              0.68, 0.70, 3.87, 3.94),
    ("Noon HQ",                  0.21, 0.20, 1.18, 1.20),
    ("Other operating expenses", 0.26, 0.24, 1.41, 1.41),
]

def avb_block(r, bar, hdr0, rows_data, note_text, total_label):
    """Returns (next_row, first_data_row, total_row). Callers must use the returned
    row numbers — recomputing them by hand is how references silently drift."""
    section_bar(ws, r, bar, width=5); r += 1
    note(ws, r, note_text, width=5); r += 1
    headers(ws, r, [hdr0] + AVB); r += 1
    first = r
    for name, am, bm, ay, by in rows_data:
        put(ws, r, 1, name, kind="plain")
        for col, val in zip(range(2, 6), (am, bm, ay, by)):
            put(ws, r, col, val, kind="input", fmt=FMT_M)
        r += 1
    last = r - 1
    total_row = r
    put(ws, r, 1, total_label, kind="total")
    for col in range(2, 6):
        L = get_column_letter(col)
        put(ws, r, col, f"=SUM({L}{first}:{L}{last})", kind="total", fmt=FMT_M)
    return r + 2, first, total_row

r, rev_first, rev_tot_row = avb_block(
    r, "  SECTION 2 — REVENUE BY BUSINESS UNIT  (USD M)", "Business Unit", REVENUE,
    "One row per business unit. Add rows as new units launch — the total recalculates.",
    "Total Revenue")
r, cost_first, cost_tot_row = avb_block(
    r, "  SECTION 3 — COSTS BY CATEGORY  (USD M)", "Cost Category", COSTS,
    "Enter all costs as POSITIVE numbers.", "Total costs")

COST_ROW = {name: cost_first + i for i, (name, *_) in enumerate(COSTS)}

# ── 4 · P&L SUMMARY ───────────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 4 — P&L SUMMARY  (USD M)", width=7); r += 1
note(ws, r, "Grey cells pull from sections 2 and 3. The script recomputes these regardless — "
            "the formulas are here so the sheet reads correctly in Excel.", width=7); r += 1
headers(ws, r, ["P&L Line"] + AVB + ["Type", "Is a cost?"]); r += 1

def src(row_num):
    return tuple(f"={get_column_letter(c)}{row_num}" for c in range(2, 6))

PL_LINES = [
    ("Revenue",                  *src(rev_tot_row),                          "value", "No"),
    ("Direct costs",             *src(COST_ROW["Direct costs"]),             "value", "Yes"),
    ("Gross profit",             None, None, None, None,                     "value", "No"),
    ("Marketing expense",        *src(COST_ROW["Marketing"]),                "value", "Yes"),
    ("Contribution profit",      None, None, None, None,                     "value", "No"),
    ("BU salaries",              *src(COST_ROW["BU salaries"]),              "value", "Yes"),
    ("Noon HQ",                  *src(COST_ROW["Noon HQ"]),                  "value", "Yes"),
    ("Other operating expenses", *src(COST_ROW["Other operating expenses"]), "value", "Yes"),
    ("EBITDA",                   None, None, None, None,                     "value", "No"),
]
pl_row = {}
for name, am, bm, ay, by, kind, is_cost in PL_LINES:
    pl_row[name] = r
    subtotal = name in ("Gross profit", "Contribution profit", "EBITDA")
    put(ws, r, 1, name, kind="plain", bold=subtotal)
    if am is not None:
        for col, formula in zip(range(2, 6), (am, bm, ay, by)):
            put(ws, r, col, formula, kind="calc", fmt=FMT_M)
    put(ws, r, 6, kind,    kind="plain", halign="center")
    put(ws, r, 7, is_cost, kind="plain", halign="center")
    r += 1

gp, cp, eb = pl_row["Gross profit"], pl_row["Contribution profit"], pl_row["EBITDA"]
rv, dc, mk = pl_row["Revenue"], pl_row["Direct costs"], pl_row["Marketing expense"]
sal, hq, oth = pl_row["BU salaries"], pl_row["Noon HQ"], pl_row["Other operating expenses"]
for col in range(2, 6):
    L = get_column_letter(col)
    put(ws, gp, col, f"={L}{rv}-{L}{dc}",                   kind="total", fmt=FMT_M)
    put(ws, cp, col, f"={L}{gp}-{L}{mk}",                   kind="total", fmt=FMT_M)
    put(ws, eb, col, f"={L}{cp}-{L}{sal}-{L}{hq}-{L}{oth}", kind="total", fmt=FMT_M)
r += 1

headers(ws, r, ["Margin"] + AVB); r += 1
for label, numer in (("Gross profit margin", gp), ("Contribution margin", cp), ("EBITDA margin", eb)):
    put(ws, r, 1, label, kind="plain")
    for col in range(2, 6):
        L = get_column_letter(col)
        put(ws, r, col, f'=IF({L}{rv}=0,"",{L}{numer}/{L}{rv})', kind="calc", fmt=FMT_PCT)
    r += 1
r += 1

# ── 5 · MONTHLY SERIES ────────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 5 — MONTHLY ACTUAL & BUDGET SERIES  (USD M)"); r += 1
note(ws, r, "Drives the period-range selector. Each row must sum across the months to its YTD figure above."); r += 1
note(ws, r, "Month columns must match the short labels in Section 1, in the same order."); r += 1

MONTHLY = {
    "Revenue — Actual": [
        ("Tracks",                [0.876, 0.904, 0.960, 0.932, 0.988, 0.990]),
        ("B2B",                   [0.566, 0.584, 0.620, 0.602, 0.638, 0.630]),
        ("Govt Schools — Legacy",  [0.461, 0.475, 0.505, 0.490, 0.519, 0.520]),
        ("Govt Schools — New",     [0.216, 0.223, 0.237, 0.230, 0.244, 0.260]),
        ("Out of School",         [0.109, 0.113, 0.120, 0.116, 0.123, 0.100]),
    ],
    "Revenue — Budget": [
        ("Tracks",                [0.959, 0.989, 1.051, 1.020, 1.081, 1.070]),
        ("B2B",                   [0.581, 0.600, 0.637, 0.618, 0.655, 0.630]),
        ("Govt Schools — Legacy",  [0.496, 0.512, 0.544, 0.528, 0.560, 0.560]),
        ("Govt Schools — New",     [0.239, 0.246, 0.262, 0.254, 0.269, 0.290]),
        ("Out of School",         [0.094, 0.097, 0.103, 0.100, 0.106, 0.100]),
    ],
    "Costs — Actual": [
        ("Direct costs",             [0.784, 0.809, 0.859, 0.834, 0.884, 0.890]),
        ("Marketing",                [0.107, 0.111, 0.117, 0.114, 0.121, 0.100]),
        ("BU salaries",              [0.600, 0.619, 0.657, 0.638, 0.676, 0.680]),
        ("Noon HQ",                  [0.182, 0.188, 0.200, 0.194, 0.206, 0.210]),
        ("Other operating expenses", [0.216, 0.223, 0.237, 0.230, 0.244, 0.260]),
    ],
    "Costs — Budget": [
        ("Direct costs",             [0.820, 0.846, 0.898, 0.872, 0.924, 0.920]),
        ("Marketing",                [0.130, 0.134, 0.142, 0.138, 0.146, 0.130]),
        ("BU salaries",              [0.609, 0.629, 0.667, 0.648, 0.687, 0.700]),
        ("Noon HQ",                  [0.188, 0.194, 0.206, 0.200, 0.212, 0.200]),
        ("Other operating expenses", [0.220, 0.227, 0.241, 0.234, 0.248, 0.240]),
    ],
}
for block_name, block_rows in MONTHLY.items():
    headers(ws, r, [block_name] + MONTH_COLS); r += 1
    first = r
    for name, vals in block_rows:
        put(ws, r, 1, name, kind="plain")
        for i, val in enumerate(vals):
            put(ws, r, 2 + i, val, kind="input", fmt=FMT_M)
        r += 1
    last = r - 1
    put(ws, r, 1, "Total", kind="total")
    for i in range(len(MONTH_COLS)):
        L = get_column_letter(2 + i)
        put(ws, r, 2 + i, f"=SUM({L}{first}:{L}{last})", kind="total", fmt=FMT_M)
    r += 2

# ── 6 · WORKING CAPITAL ───────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 6 — WORKING CAPITAL  (USD M)", width=8); r += 1
note(ws, r, "Closing balances. Enter payables and deferred revenue as NEGATIVE numbers.", width=8); r += 1
headers(ws, r, ["WC Month", "Receivables", "Payables", "Deferred revenue",
                "Other WC", "Net WC", "Movement", "Forecast?"]); r += 1
wc_first = r
for mo, ar_, apv, dfr, oth_, fc in [
    ("Jan-26", 5.9, -2.9, -2.4, 0.5, "No"), ("Feb-26", 6.1, -2.8, -2.3, 0.5, "No"),
    ("Mar-26", 6.3, -2.7, -2.2, 0.6, "No"), ("Apr-26", 6.4, -2.9, -2.2, 0.5, "No"),
    ("May-26", 6.6, -2.8, -2.1, 0.6, "No"), ("Jun-26", 6.7, -2.7, -2.1, 0.6, "No"),
    ("Jul-26", 6.8, -2.6, -2.0, 0.6, "Yes"),("Aug-26", 7.1, -2.4, -1.8, 0.3, "Yes"),
]:
    put(ws, r, 1, mo,   kind="input")
    put(ws, r, 2, ar_,  kind="input", fmt=FMT_M)
    put(ws, r, 3, apv,  kind="input", fmt=FMT_M)
    put(ws, r, 4, dfr,  kind="input", fmt=FMT_M)
    put(ws, r, 5, oth_, kind="input", fmt=FMT_M)
    put(ws, r, 6, f"=SUM(B{r}:E{r})", kind="calc", fmt=FMT_M)
    put(ws, r, 7, 0 if r == wc_first else f"=F{r}-F{r-1}", kind="calc", fmt=FMT_MS)
    put(ws, r, 8, fc,   kind="input", halign="center")
    r += 1
r += 1

note(ws, r, "Opening versus closing balance for the year to date. Cash impact is the mirror of the movement.", width=5); r += 1
headers(ws, r, ["WC Item", "At 1 Jan", "At period end", "Movement", "Cash impact"]); r += 1
wy_first = r
for item, op, cl in [("Accounts receivable", 5.9, 6.7), ("Accounts payable", -2.9, -2.7),
                     ("Deferred revenue", -2.4, -2.1), ("Other working capital items", 0.5, 0.6)]:
    put(ws, r, 1, item, kind="plain")
    put(ws, r, 2, op,   kind="input", fmt=FMT_M)
    put(ws, r, 3, cl,   kind="input", fmt=FMT_M)
    put(ws, r, 4, f"=C{r}-B{r}",    kind="calc", fmt=FMT_MS)
    put(ws, r, 5, f"=-(C{r}-B{r})", kind="calc", fmt=FMT_MS)
    r += 1
wy_last = r - 1
put(ws, r, 1, "Net working capital", kind="total")
for col in range(2, 6):
    L = get_column_letter(col)
    put(ws, r, col, f"=SUM({L}{wy_first}:{L}{wy_last})", kind="total",
        fmt=FMT_M if col in (2, 3) else FMT_MS)
r += 2

# ── 7 · ACCOUNTS RECEIVABLE ───────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 7 — ACCOUNTS RECEIVABLE  (USD M)", width=6); r += 1
note(ws, r, "Drives the AR column chart. Closing balance and collection rate are calculated.", width=6); r += 1
headers(ws, r, ["AR Month", "Opening AR", "Invoiced", "Collected", "Closing AR", "Collection rate"]); r += 1
ar_first = r
for mo, op, inv, coll in [
    ("Jan-26", 5.70, 2.30, 2.10), ("Feb-26", 5.90, 2.25, 2.05),
    ("Mar-26", 6.10, 2.55, 2.35), ("Apr-26", 6.30, 2.40, 2.30),
    ("May-26", 6.40, 2.35, 2.15), ("Jun-26", 6.60, 2.50, 2.40),
    ("Jul-26", 6.70, 2.55, 2.45), ("Aug-26", 6.80, 2.60, 2.30),
]:
    put(ws, r, 1, mo,   kind="input")
    put(ws, r, 2, op,   kind="input", fmt=FMT_M)
    put(ws, r, 3, inv,  kind="input", fmt=FMT_M)
    put(ws, r, 4, coll, kind="input", fmt=FMT_M)
    put(ws, r, 5, f"=B{r}+C{r}-D{r}",       kind="calc", fmt=FMT_M)
    put(ws, r, 6, f'=IF(C{r}=0,"",D{r}/C{r})', kind="calc", fmt=FMT_PCT)
    r += 1
put(ws, r, 1, "YTD total", kind="total")
put(ws, r, 2, "", kind="total")
put(ws, r, 3, f"=SUM(C{ar_first}:C{ar_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 4, f"=SUM(D{ar_first}:D{ar_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 5, f"=E{ar_first+5}", kind="total", fmt=FMT_M)
put(ws, r, 6, f'=IF(C{r}=0,"",D{r}/C{r})', kind="total", fmt=FMT_PCT)
r += 2

def share_table(r, note_text, header, rows_data, total_label, name_kind="plain"):
    note(ws, r, note_text, width=3); r += 1
    headers(ws, r, [header, "Amount", "Share"]); r += 1
    first = r
    for label, amt in rows_data:
        put(ws, r, 1, label, kind=name_kind)
        put(ws, r, 2, amt,   kind="input", fmt=FMT_M)
        r += 1
    last, total = r - 1, r
    for rr in range(first, last + 1):
        put(ws, rr, 3, f'=IF($B${total}=0,"",B{rr}/$B${total})', kind="calc", fmt=FMT_PCT)
    put(ws, total, 1, total_label, kind="total")
    put(ws, total, 2, f"=SUM(B{first}:B{last})", kind="total", fmt=FMT_M)
    put(ws, total, 3, 1.0, kind="total", fmt=FMT_PCT)
    return total + 2

r = share_table(r, "Balance at the period end, split by age.", "AR Aging Bucket",
                [("Current — due next 30 days", 2.8), ("31–60 days", 2.0), ("60+ days / overdue", 1.9)],
                "Total AR")
r = share_table(r, "Largest receivable balances by contract or customer. Add rows as needed.", "Contract",
                [("Ministry of Education — Framework", 2.10), ("Riyadh Schools Group", 1.50),
                 ("Jeddah Private Academies", 1.20), ("Eastern Province Consortium", 0.80),
                 ("Tracks — corporate accounts", 0.60), ("Other contracts", 0.50)],
                "Total", name_kind="input")

# ── 8 · ACCOUNTS PAYABLE ──────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 8 — ACCOUNTS PAYABLE  (USD M)", width=6); r += 1
note(ws, r, "Drives the AP column chart. Enter DPO in days.", width=6); r += 1
headers(ws, r, ["AP Month", "Opening AP", "Purchases", "Payments", "Closing AP", "DPO (days)"]); r += 1
ap_first = r
for mo, op, pur, pay, dpo in [
    ("Jan-26", 2.95, 2.00, 2.05, 43.5), ("Feb-26", 2.90, 1.95, 2.05, 43.1),
    ("Mar-26", 2.80, 2.00, 2.10, 40.5), ("Apr-26", 2.70, 2.10, 1.90, 41.4),
    ("May-26", 2.90, 1.95, 2.05, 43.1), ("Jun-26", 2.80, 2.05, 2.15, 39.5),
    ("Jul-26", 2.70, 2.00, 2.10, 39.0), ("Aug-26", 2.60, 1.90, 2.10, 37.9),
]:
    put(ws, r, 1, mo,  kind="input")
    put(ws, r, 2, op,  kind="input", fmt=FMT_M)
    put(ws, r, 3, pur, kind="input", fmt=FMT_M)
    put(ws, r, 4, pay, kind="input", fmt=FMT_M)
    put(ws, r, 5, f"=B{r}+C{r}-D{r}", kind="calc", fmt=FMT_M)
    put(ws, r, 6, dpo, kind="input", fmt=FMT_NUM)
    r += 1
put(ws, r, 1, "YTD total", kind="total")
put(ws, r, 2, "", kind="total")
put(ws, r, 3, f"=SUM(C{ap_first}:C{ap_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 4, f"=SUM(D{ap_first}:D{ap_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 5, f"=E{ap_first+5}", kind="total", fmt=FMT_M)
put(ws, r, 6, f"=AVERAGE(F{ap_first}:F{ap_first+5})", kind="total", fmt=FMT_NUM)
r += 2

r = share_table(r, "Balance at the period end, split by age.", "AP Aging Bucket",
                [("Current — due next 30 days", 1.5), ("31–60 days", 0.8), ("60+ days / overdue", 0.4)],
                "Total AP")
r = share_table(r, "Largest payable balances by vendor. Add rows as needed.", "Vendor",
                [("AWS / cloud infrastructure", 0.62), ("Content production partners", 0.48),
                 ("Facilities & office leases", 0.37), ("Marketing agencies", 0.29),
                 ("Professional services", 0.21), ("Other vendors", 0.18)],
                "Total", name_kind="input")

# ── 9 · CASH ──────────────────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 9 — CASH  (USD M)", width=9); r += 1
note(ws, r, "Four KPI tiles, each shown twice — once for the month, once for YTD. "
            "The note is the small print under the number.", width=9); r += 1
headers(ws, r, ["Cash Tile", "Month value", "Month note", "YTD value", "YTD note"],
        merges={2: 3, 4: 3}); r += 1
for name, mv, mn, yv, yn in [
    ("Cash balance",        4.08, "29.1 months runway",        4.08, "opened the year at $4.80M"),
    ("Collections",         2.40, "Budget $2.50M invoiced",   13.35, "on $14.35M invoiced YTD"),
    ("Net working capital", 3.20, "+$0.40M vs prior month",    1.40, "movement YTD · closing $2.50M"),
    ("Accounts receivable", 7.10, "88% collection rate",       7.10, "from $5.70M at 1 Jan"),
]:
    put(ws, r, 1, name, kind="plain", bold=True)
    put(ws, r, 2, mv, kind="input", fmt=FMT_M)
    put(ws, r, 3, mn, kind="input", span=3, wrap=True)
    put(ws, r, 6, yv, kind="input", fmt=FMT_M)
    put(ws, r, 7, yn, kind="input", span=3, wrap=True)
    r += 1
r += 1

def bridge_block(r, note_text, header, steps):
    note(ws, r, note_text, width=4); r += 1
    headers(ws, r, [header, "Type", "Amount", "Running balance"]); r += 1
    first = r
    for step, kind_, amt in steps:
        put(ws, r, 1, step,  kind="input")
        put(ws, r, 2, kind_, kind="input", halign="center")
        put(ws, r, 3, amt,   kind="input", fmt=FMT_M)
        put(ws, r, 4, f"=C{r}" if r == first else
            f'=IF(B{r}="Closing",D{r-1},IF(B{r}="Inflow",D{r-1}+C{r},D{r-1}-C{r}))',
            kind="calc", fmt=FMT_M)
        r += 1
    return r + 1

r = bridge_block(r, 'Type must be Opening, Inflow, Outflow or Closing. Enter every amount as POSITIVE.',
                 "Bridge Step (Month)", [
    ("Opening cash", "Opening", 4.40), ("Collections", "Inflow", 2.40),
    ("Other inflows", "Inflow", 0.10), ("Operating costs", "Outflow", 2.08),
    ("Capex", "Outflow", 0.16), ("Debt service", "Outflow", 0.16),
    ("Other outflows", "Outflow", 0.42), ("Closing cash", "Closing", 4.08)])
r = bridge_block(r, 'Same rules as the monthly bridge. Opening is the 1 January balance.',
                 "Bridge Step (YTD)", [
    ("Opening cash — 1 Jan 2026", "Opening", 4.80), ("Collections", "Inflow", 13.35),
    ("Other inflows", "Inflow", 0.59), ("Operating costs", "Outflow", 11.99),
    ("Capex", "Outflow", 0.89), ("Debt service", "Outflow", 0.89),
    ("Other outflows", "Outflow", 0.89), ("Closing cash — 30 Jun 2026", "Closing", 4.08)])

note(ws, r, "Drives the cash balance line chart. It can run wider than the months in Section 1.", width=4); r += 1
note(ws, r, "Forecast = Yes draws a dashed line. Illustrative = Yes draws a hollow marker "
            "and shows the 'indicative' flag.", width=4); r += 1
headers(ws, r, ["Cash Month", "Closing cash", "Forecast?", "Illustrative?"]); r += 1
for mo, val, fc, ill in [
    ("Jul-25", 5.35, "No", "Yes"), ("Aug-25", 5.22, "No", "Yes"), ("Sep-25", 5.10, "No", "Yes"),
    ("Oct-25", 5.02, "No", "Yes"), ("Nov-25", 4.91, "No", "Yes"), ("Dec-25", 4.80, "No", "No"),
    ("Jan-26", 4.62, "No", "No"),  ("Feb-26", 4.55, "No", "No"),  ("Mar-26", 4.49, "No", "No"),
    ("Apr-26", 4.44, "No", "No"),  ("May-26", 4.40, "No", "No"),  ("Jun-26", 4.08, "No", "No"),
    ("Jul-26", 3.95, "Yes", "No"), ("Aug-26", 3.86, "Yes", "No"), ("Sep-26", 3.80, "Yes", "No"),
    ("Oct-26", 3.72, "Yes", "No"), ("Nov-26", 3.66, "Yes", "No"), ("Dec-26", 3.58, "Yes", "No"),
]:
    put(ws, r, 1, mo,  kind="input")
    put(ws, r, 2, val, kind="input", fmt=FMT_M)
    put(ws, r, 3, fc,  kind="input", halign="center")
    put(ws, r, 4, ill, kind="input", halign="center")
    r += 1
r += 1

note(ws, r, "Burn figures are negative when cash is being consumed.", width=5); r += 1
headers(ws, r, ["Runway Metric", "Value", "Notes"], merges={2: 3}); r += 1
for metric, val, nt in [
    ("Monthly cash burn",             -0.32, "Net cash movement in the latest month"),
    ("Trailing 3-month average burn",  -0.14, "Average of the last three months"),
    ("YTD net cash movement",          -0.72, "Closing cash less opening cash"),
    ("Cash runway (months)",           29.14, "Cash ÷ 3-month average burn"),
]:
    put(ws, r, 1, metric, kind="plain")
    put(ws, r, 2, val, kind="input", fmt=FMT_NUM if "months" in metric else FMT_M)
    put(ws, r, 3, nt, kind="plain", span=3)
    r += 1
r += 1

# ── 10 · KEY UPDATES ──────────────────────────────────────────────────────────
section_bar(ws, r, "  SECTION 10 — KEY NARRATIVE UPDATES", width=LAST_COL); r += 1
note(ws, r, "One row per commentary point, shown in order at the bottom of the dashboard."); r += 1
headers(ws, r, ["Update #", "Topic", "Commentary"], merges={2: 11}); r += 1
for i, (topic, text) in enumerate([
    ("Revenue",         "June revenue of $2.50M landed at 94.3% of budget, missing plan in five of the first six months of the year. YTD revenue of $14.35M is $0.90M (5.9%) behind budget."),
    ("Margin",          "EBITDA of $0.36M gave a 14.4% margin against a 17.4% budget. YTD EBITDA of $2.16M is $0.44M behind plan; the gap is driven by revenue, not cost overrun."),
    ("Costs",           "Cost control is holding — total costs ran at 97.7% of budget for the month and 96.4% YTD. The underspend has absorbed a meaningful share of the revenue shortfall."),
    ("Collections",     "The June collection rate of 96.0% is the strongest month of the year so far. AR stands at $6.70M, of which $1.90M (28.4%) is now 60+ days overdue."),
    ("Working capital", "Net working capital has absorbed $1.40M of cash since 1 January, almost entirely through the AR build. Payables have been drawn down $0.20M over the same period."),
    ("Cash",            "Cash closed at $4.08M, down $0.72M since 1 January. At the trailing three-month burn of $0.14M per month that is 29.1 months of runway, before the $1.50M current portion of Facility A."),
], 1):
    put(ws, r, 1, i, kind="plain", halign="center")
    put(ws, r, 2, topic, kind="input")
    put(ws, r, 3, text,  kind="input", span=11, wrap=True)
    ws.row_dimensions[r].height = 34
    r += 1

ws.freeze_panes = "B2"
ws.sheet_view.showGridLines = False

wb.save(args.out)
print(f"✓ Created {args.out}")
print(f"  Tabs: {', '.join(wb.sheetnames)}  ·  {r} rows on '{SHEET}'")
