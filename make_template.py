"""
make_template.py — Generates noon_dashboard_input.xlsx, the monthly input workbook
for the Noon P&L Performance dashboard.

Every chart in the dashboard has a dedicated table here. Tables are located by
their header text (not by row number), so inserting or deleting rows is safe.

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

# ── Style tokens ──────────────────────────────────────────────────────────────
FONT      = "Arial"
C_HEADER  = "11203A"   # Noon dark blue — section bars
C_SUBHDR  = "D9E1F2"   # column header fill
C_INPUT   = "FFF3CD"   # yellow — cells the user edits
C_CALC    = "F2F2F2"   # grey — derived/reference
C_TOTAL   = "BDD7EE"   # blue — totals
C_WHITE   = "FFFFFF"
BLUE_TXT  = "0000FF"   # hardcoded input convention
BLACK_TXT = "000000"

FMT_M    = '$#,##0.00;($#,##0.00);-'    # USD millions
FMT_MS   = '+$#,##0.00;-$#,##0.00;-'    # signed
FMT_PCT  = '0.0%'
FMT_NUM  = '#,##0.0'

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

def section_bar(ws, row, text, width=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = F(bold=True, color=C_WHITE, size=11)
    c.fill = fill(C_HEADER)
    c.alignment = A("left")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    ws.row_dimensions[row].height = 20

def note(ws, row, text, col=1):
    c = ws.cell(row=row, column=col, value=text)
    c.font = F(size=9, italic=True, color="595959")
    return c

def headers(ws, row, labels, start=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=t)
        c.font      = F(bold=True, size=10)
        c.fill      = fill(C_SUBHDR)
        c.alignment = A("center", wrap=True)
        c.border    = BOX
    ws.row_dimensions[row].height = 28

def put(ws, row, col, val, *, kind="input", fmt=None, bold=False, wrap=False, halign=None):
    """kind: input (yellow/blue) | calc (grey) | total (blue) | text | plain"""
    c = ws.cell(row=row, column=col, value=val)
    c.border = BOX
    c.font   = F(bold=bold, color=BLUE_TXT if kind == "input" and isinstance(val, (int, float)) else BLACK_TXT)
    if kind == "input":  c.fill = fill(C_INPUT)
    elif kind == "calc": c.fill = fill(C_CALC)
    elif kind == "total":
        c.fill = fill(C_TOTAL); c.font = F(bold=True)
    else:                c.fill = fill(C_WHITE)
    if fmt: c.number_format = fmt
    c.alignment = A(halign or ("right" if isinstance(val, (int, float)) else "left"), wrap=wrap)
    return c

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INSTRUCTIONS
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Instructions")
widths(ws, {"A": 4, "B": 26, "C": 78, "D": 30})

ws["B2"] = "Noon Academy — Dashboard Input Workbook"
ws["B2"].font = Font(name=FONT, bold=True, size=16, color=C_HEADER)
ws["B3"] = "Fill this workbook each month, then regenerate the dashboard."
ws["B3"].font = F(size=10, italic=True, color="595959")

r = 5
section_bar(ws, r, "  HOW TO USE THIS WORKBOOK", width=4); r += 2

steps = [
    ("1.  Update the data",
     "Work through the tabs left to right. Every cell shaded YELLOW is an input — type over it. "
     "Grey cells are derived and white cells are labels; leave both alone."),
    ("2.  Keep the shape",
     "Add or remove rows inside a table freely (e.g. a new business unit, another vendor). "
     "Do NOT delete a table's header row — the script finds each table by its header text, not by row number."),
    ("3.  Regenerate",
     "Save and close the file, then run:      python etl_pl.py --file noon_dashboard_input.xlsx"),
    ("4.  Open the dashboard",
     "The script rewrites index.html. Open it in any browser — no server, no Python needed to view it."),
]
for title, body in steps:
    ws.cell(row=r, column=2, value=title).font = F(bold=True, size=11)
    c = ws.cell(row=r, column=3, value=body); c.font = F(size=10); c.alignment = A(wrap=True, v="top")
    ws.row_dimensions[r].height = 32
    r += 1

r += 1
section_bar(ws, r, "  WHAT EACH TAB FEEDS", width=4); r += 1
headers(ws, r, ["", "Tab", "What it contains", "Dashboard element it drives"]); r += 1

tabs = [
    ("Setup",           "Period labels, reporting year, list of actual months.",
                        "Every header, the period selector"),
    ("P&L",             "Revenue by BU, costs by category, P&L summary — month and YTD, actual vs budget.",
                        "Revenue & cost bullet charts, all three P&L tables"),
    ("P&L Monthly",     "Month-by-month actual and budget for every revenue, cost and P&L line.",
                        "Period-range selector (quarters, custom ranges)"),
    ("Working Capital", "Monthly WC position and the YTD movement bridge.",
                        "Working capital table"),
    ("AR",              "Monthly AR flow, aging buckets, balance by contract.",
                        "AR column chart, AR aging bars, AR by contract bars"),
    ("AP",              "Monthly AP flow, aging buckets, balance by vendor.",
                        "AP column chart, AP aging bars, AP by vendor bars"),
    ("Cash",            "KPI tiles, cash bridges (month & YTD), monthly closing cash, runway metrics.",
                        "Cash tiles, both waterfalls, cash balance line chart"),
    ("Key Updates",     "The narrative commentary paragraphs.",
                        "Key updates section at the bottom"),
]
for name, contains, drives in tabs:
    put(ws, r, 2, name, kind="plain", bold=True)
    put(ws, r, 3, contains, kind="plain", wrap=True)
    put(ws, r, 4, drives,   kind="plain", wrap=True)
    ws.row_dimensions[r].height = 30
    r += 1

r += 1
section_bar(ws, r, "  CELL COLOUR LEGEND", width=4); r += 1
headers(ws, r, ["", "Colour", "Meaning", ""]); r += 1
legend = [
    (C_INPUT, "Yellow", "An input. Type your number or text here."),
    (C_CALC,  "Grey",   "Derived from your inputs, or a reference value. Do not edit."),
    (C_TOTAL, "Blue",   "A total. Recalculated by the script from the rows above."),
    (C_WHITE, "White",  "A label or heading."),
]
for hexc, name, mean in legend:
    ws.cell(row=r, column=2, value=name).fill = fill(hexc)
    ws.cell(row=r, column=2).font = F(bold=True)
    ws.cell(row=r, column=2).border = BOX
    put(ws, r, 3, mean, kind="plain", wrap=True)
    r += 1

r += 1
section_bar(ws, r, "  CONVENTIONS", width=4); r += 1
convs = [
    ("Units",        "All monetary figures are USD millions. Enter 2.5 for $2.5M — not 2500000."),
    ("Percentages",  "Enter as a percentage-formatted number (96.0%), not 0.96 or 96."),
    ("Costs",        "Enter costs as POSITIVE numbers. The dashboard applies the sign."),
    ("Working capital", "AP and deferred revenue are entered NEGATIVE (they are liabilities)."),
    ("Months",       "Use the 'Mmm-YY' format exactly: Jan-26, Feb-26. The script matches on this."),
    ("Forecast",     "In monthly tables, mark a row Y under 'Forecast?' to grey it in the dashboard."),
    ("Blank rows",   "A blank row ends a table. Don't leave gaps in the middle of one."),
]
for k, vtext in convs:
    put(ws, r, 2, k, kind="plain", bold=True)
    put(ws, r, 3, vtext, kind="plain", wrap=True)
    ws.row_dimensions[r].height = 26
    r += 1

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — SETUP
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Setup")
widths(ws, {"A": 34, "B": 40, "C": 60})

section_bar(ws, 1, "  REPORTING PERIOD", width=3)
note(ws, 2, "These labels appear in every dashboard header. Update them first each month.")
headers(ws, 4, ["Setting", "Value", "Notes"])

setup_rows = [
    ("Period label (month)", "June 2026",      "The latest closed month, e.g. 'July 2026'"),
    ("YTD label",            "Jan–Jun 2026",   "Range covered year-to-date"),
    ("Reporting year",       "2026",           "Calendar year of the current period"),
    ("Fiscal year label",    "FY2025/26",      "Shown next to quarter labels"),
    ("Fiscal year start month", 7,             "1 = January … 7 = July"),
    ("Currency label",       "USD M",          "Displayed under chart titles"),
]
r = 5
for k, val, nt in setup_rows:
    put(ws, r, 1, k, kind="plain", bold=True)
    put(ws, r, 2, val, kind="input")
    put(ws, r, 3, nt, kind="plain")
    r += 1

r += 1
section_bar(ws, r, "  ACTUAL MONTHS", width=3); r += 1
note(ws, r, "One row per closed month, oldest first. This defines the period selector and the order of every monthly chart."); r += 1
headers(ws, r, ["Month", "Short label", "Notes"]); r += 1

months_2026 = [("Jan-26","Jan"),("Feb-26","Feb"),("Mar-26","Mar"),
               ("Apr-26","Apr"),("May-26","May"),("Jun-26","Jun")]
for full, short in months_2026:
    put(ws, r, 1, full,  kind="input")
    put(ws, r, 2, short, kind="input")
    put(ws, r, 3, "", kind="plain")
    r += 1

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — P&L  (month + YTD, actual vs budget)
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("P&L")
widths(ws, {"A": 46, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13})

AVB_HDR = ["", "Actual (Month)", "Budget (Month)", "Actual YTD", "Budget YTD"]

def avb_block(ws, start_row, bar, hdr0, rows_data, note_text=None, total_label=None):
    """Returns (next_free_row, first_data_row, total_row). Callers use the returned
    row numbers directly — never recompute them, or references drift silently."""
    r = start_row
    section_bar(ws, r, bar, width=5); r += 1
    if note_text:
        note(ws, r, note_text); r += 1
    headers(ws, r, [hdr0] + AVB_HDR[1:]); r += 1
    first = r
    for name, am, bm, ay, by in rows_data:
        put(ws, r, 1, name, kind="plain")
        put(ws, r, 2, am, kind="input", fmt=FMT_M)
        put(ws, r, 3, bm, kind="input", fmt=FMT_M)
        put(ws, r, 4, ay, kind="input", fmt=FMT_M)
        put(ws, r, 5, by, kind="input", fmt=FMT_M)
        r += 1
    last = r - 1
    total_row = None
    if total_label:
        total_row = r
        put(ws, r, 1, total_label, kind="total", bold=True)
        for col in range(2, 6):
            L = get_column_letter(col)
            put(ws, r, col, f"=SUM({L}{first}:{L}{last})", kind="total", fmt=FMT_M)
        r += 1
    return r + 1, first, total_row

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
r, rev_first, rev_tot_row = avb_block(
    ws, 1, "  REVENUE BY BUSINESS UNIT  (USD M)", "Business Unit", REVENUE,
    "One row per business unit. Add rows as new units launch — the total recalculates.",
    "Total Revenue")
r, cost_first, cost_tot_row = avb_block(
    ws, r, "  COSTS BY CATEGORY  (USD M)", "Cost Category", COSTS,
    "Enter all costs as POSITIVE numbers.", "Total costs")

# Named row numbers for each cost line, so P&L references cannot drift
COST_ROW = {name: cost_first + i for i, (name, *_) in enumerate(COSTS)}

# P&L summary — derived, but left editable so it can be overridden
section_bar(ws, r, "  P&L SUMMARY  (USD M)", width=7); r += 1
note(ws, r, "Grey cells pull from the blocks above. Override only if your P&L differs from Revenue − Costs."); r += 1
headers(ws, r, ["P&L Line", "Actual (Month)", "Budget (Month)", "Actual YTD", "Budget YTD",
                "Type", "Is a cost?"]); r += 1

def src(row_num):
    """Column formulas pulling one source row across all four value columns."""
    return tuple(f"={get_column_letter(c)}{row_num}" for c in range(2, 6))

PL_LINES = [
    ("Revenue",                  *src(rev_tot_row),                         "value", "No"),
    ("Direct costs",             *src(COST_ROW["Direct costs"]),            "value", "Yes"),
    ("Gross profit",             None, None, None, None,                    "value", "No"),
    ("Marketing expense",        *src(COST_ROW["Marketing"]),               "value", "Yes"),
    ("Contribution profit",      None, None, None, None,                    "value", "No"),
    ("BU salaries",              *src(COST_ROW["BU salaries"]),             "value", "Yes"),
    ("Noon HQ",                  *src(COST_ROW["Noon HQ"]),                 "value", "Yes"),
    ("Other operating expenses", *src(COST_ROW["Other operating expenses"]),"value", "Yes"),
    ("EBITDA",                   None, None, None, None,                    "value", "No"),
]
pl_first = r
pl_rownum = {}
for i, (name, am, bm, ay, by, kind, is_cost) in enumerate(PL_LINES):
    pl_rownum[name] = r
    put(ws, r, 1, name, kind="plain", bold=name in ("Gross profit","Contribution profit","EBITDA"))
    if am is None:   # computed below
        pass
    else:
        for col, formula in zip(range(2, 6), (am, bm, ay, by)):
            put(ws, r, col, formula, kind="calc", fmt=FMT_M)
    put(ws, r, 6, kind, kind="plain")
    put(ws, r, 7, is_cost, kind="plain")
    r += 1

# Fill the three computed subtotal lines
gp, cp, eb = pl_rownum["Gross profit"], pl_rownum["Contribution profit"], pl_rownum["EBITDA"]
rv, dc, mk = pl_rownum["Revenue"], pl_rownum["Direct costs"], pl_rownum["Marketing expense"]
sal, hq, oth = pl_rownum["BU salaries"], pl_rownum["Noon HQ"], pl_rownum["Other operating expenses"]
for col in range(2, 6):
    L = get_column_letter(col)
    put(ws, gp, col, f"={L}{rv}-{L}{dc}",                          kind="total", fmt=FMT_M)
    put(ws, cp, col, f"={L}{gp}-{L}{mk}",                          kind="total", fmt=FMT_M)
    put(ws, eb, col, f"={L}{cp}-{L}{sal}-{L}{hq}-{L}{oth}",        kind="total", fmt=FMT_M)

r += 1
section_bar(ws, r, "  MARGINS  (calculated)", width=5); r += 1
headers(ws, r, ["Margin", "Actual (Month)", "Budget (Month)", "Actual YTD", "Budget YTD"]); r += 1
for label, num in (("Gross profit margin", gp), ("Contribution margin", cp), ("EBITDA margin", eb)):
    put(ws, r, 1, label, kind="plain")
    for col in range(2, 6):
        L = get_column_letter(col)
        put(ws, r, col, f"=IF({L}{rv}=0,\"\",{L}{num}/{L}{rv})", kind="calc", fmt=FMT_PCT)
    r += 1

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — P&L MONTHLY  (drives the period-range selector)
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("P&L Monthly")
widths(ws, {"A": 46})
for i in range(2, 9):
    ws.column_dimensions[get_column_letter(i)].width = 12

MONTH_COLS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]

# Realistic monthly series that sum to the YTD figures on the P&L tab
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

r = 1
section_bar(ws, r, "  MONTHLY ACTUAL & BUDGET SERIES  (USD M)", width=7); r += 1
note(ws, r, "Drives the period-range selector. Each row must sum across the months to its YTD figure on the P&L tab."); r += 1
note(ws, r, "Month columns must match the short labels on the Setup tab, in the same order."); r += 2

for block_name, block_rows in MONTHLY.items():
    section_bar(ws, r, f"  {block_name.upper()}", width=7); r += 1
    headers(ws, r, [block_name] + MONTH_COLS); r += 1
    first = r
    for name, vals in block_rows:
        put(ws, r, 1, name, kind="plain")
        for i, val in enumerate(vals):
            put(ws, r, 2 + i, val, kind="input", fmt=FMT_M)
        r += 1
    last = r - 1
    put(ws, r, 1, "Total", kind="total", bold=True)
    for i in range(len(MONTH_COLS)):
        L = get_column_letter(2 + i)
        put(ws, r, 2 + i, f"=SUM({L}{first}:{L}{last})", kind="total", fmt=FMT_M)
    r += 3

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — WORKING CAPITAL
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Working Capital")
widths(ws, {"A": 30, "B": 14, "C": 14, "D": 15, "E": 14, "F": 14, "G": 14, "H": 12})

r = 1
section_bar(ws, r, "  WORKING CAPITAL — MONTHLY POSITION  (USD M)", width=8); r += 1
note(ws, r, "Closing balances. Enter AP and deferred revenue as NEGATIVE numbers."); r += 1
headers(ws, r, ["Month", "Receivables", "Payables", "Deferred revenue",
                "Other WC", "Net WC", "Movement", "Forecast?"]); r += 1

WC_MONTHLY = [
    ("Jan-26", 5.9, -2.9, -2.4, 0.5, "No"),
    ("Feb-26", 6.1, -2.8, -2.3, 0.5, "No"),
    ("Mar-26", 6.3, -2.7, -2.2, 0.6, "No"),
    ("Apr-26", 6.4, -2.9, -2.2, 0.5, "No"),
    ("May-26", 6.6, -2.8, -2.1, 0.6, "No"),
    ("Jun-26", 6.7, -2.7, -2.1, 0.6, "No"),
    ("Jul-26", 6.8, -2.6, -2.0, 0.6, "Yes"),
    ("Aug-26", 7.1, -2.4, -1.8, 0.3, "Yes"),
]
wc_first = r
for mo, ar, apv, dfr, oth_, fc in WC_MONTHLY:
    put(ws, r, 1, mo,  kind="input")
    put(ws, r, 2, ar,  kind="input", fmt=FMT_M)
    put(ws, r, 3, apv, kind="input", fmt=FMT_M)
    put(ws, r, 4, dfr, kind="input", fmt=FMT_M)
    put(ws, r, 5, oth_,kind="input", fmt=FMT_M)
    put(ws, r, 6, f"=SUM(B{r}:E{r})", kind="calc", fmt=FMT_M)
    put(ws, r, 7, f"=F{r}-F{r-1}" if r > wc_first else 0, kind="calc", fmt=FMT_MS)
    put(ws, r, 8, fc, kind="input", halign="center")
    r += 1

r += 1
section_bar(ws, r, "  WORKING CAPITAL — YTD MOVEMENT  (USD M)", width=5); r += 1
note(ws, r, "Opening vs closing balance for the year to date. Cash impact is the mirror of the movement."); r += 1
headers(ws, r, ["WC Item", "At 1 Jan", "At period end", "Movement", "Cash impact"]); r += 1

WC_YTD = [
    ("Accounts receivable",          5.9,  6.7),
    ("Accounts payable",            -2.9, -2.7),
    ("Deferred revenue",            -2.4, -2.1),
    ("Other working capital items",  0.5,  0.6),
]
wy_first = r
for item, op, cl in WC_YTD:
    put(ws, r, 1, item, kind="plain")
    put(ws, r, 2, op,   kind="input", fmt=FMT_M)
    put(ws, r, 3, cl,   kind="input", fmt=FMT_M)
    put(ws, r, 4, f"=C{r}-B{r}",  kind="calc", fmt=FMT_MS)
    put(ws, r, 5, f"=-(C{r}-B{r})", kind="calc", fmt=FMT_MS)
    r += 1
wy_last = r - 1
put(ws, r, 1, "Net working capital", kind="total", bold=True)
for col, L in ((2,"B"),(3,"C"),(4,"D"),(5,"E")):
    put(ws, r, col, f"=SUM({L}{wy_first}:{L}{wy_last})", kind="total",
        fmt=FMT_M if col in (2,3) else FMT_MS)

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — AR
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("AR")
widths(ws, {"A": 32, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16})

r = 1
section_bar(ws, r, "  ACCOUNTS RECEIVABLE — MONTHLY FLOW  (USD M)", width=6); r += 1
note(ws, r, "Drives the AR column chart. Closing and the collection rate are calculated."); r += 1
headers(ws, r, ["Month", "Opening AR", "Invoiced", "Collected", "Closing AR", "Collection rate"]); r += 1

AR_MONTHLY = [
    ("Jan-26", 5.70, 2.30, 2.10),
    ("Feb-26", 5.90, 2.25, 2.05),
    ("Mar-26", 6.10, 2.55, 2.35),
    ("Apr-26", 6.30, 2.40, 2.30),
    ("May-26", 6.40, 2.35, 2.15),
    ("Jun-26", 6.60, 2.50, 2.40),
    ("Jul-26", 6.70, 2.55, 2.45),
    ("Aug-26", 6.80, 2.60, 2.30),
]
ar_first = r
for mo, op, inv, coll in AR_MONTHLY:
    put(ws, r, 1, mo,   kind="input")
    put(ws, r, 2, op,   kind="input", fmt=FMT_M)
    put(ws, r, 3, inv,  kind="input", fmt=FMT_M)
    put(ws, r, 4, coll, kind="input", fmt=FMT_M)
    put(ws, r, 5, f"=B{r}+C{r}-D{r}", kind="calc", fmt=FMT_M)
    put(ws, r, 6, f"=IF(C{r}=0,\"\",D{r}/C{r})", kind="calc", fmt=FMT_PCT)
    r += 1
ar_last = r - 1
put(ws, r, 1, "YTD total", kind="total", bold=True)
put(ws, r, 2, "", kind="total")
put(ws, r, 3, f"=SUM(C{ar_first}:C{ar_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 4, f"=SUM(D{ar_first}:D{ar_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 5, f"=E{ar_first+5}", kind="total", fmt=FMT_M)
put(ws, r, 6, f"=IF(C{r}=0,\"\",D{r}/C{r})", kind="total", fmt=FMT_PCT)
r += 2

section_bar(ws, r, "  AR AGING  (USD M)", width=3); r += 1
note(ws, r, "Balance at the period end, split by age. Share is calculated."); r += 1
headers(ws, r, ["AR Aging Bucket", "Amount", "Share"]); r += 1
AR_AGING = [("Current — due next 30 days", 2.8), ("31–60 days", 2.0), ("60+ days / overdue", 1.9)]
ag_first = r
for bucket, amt in AR_AGING:
    put(ws, r, 1, bucket, kind="plain")
    put(ws, r, 2, amt,    kind="input", fmt=FMT_M)
    r += 1
ag_last = r - 1
for rr in range(ag_first, ag_last + 1):
    put(ws, rr, 3, f"=IF($B${r}=0,\"\",B{rr}/$B${r})", kind="calc", fmt=FMT_PCT)
put(ws, r, 1, "Total AR", kind="total", bold=True)
put(ws, r, 2, f"=SUM(B{ag_first}:B{ag_last})", kind="total", fmt=FMT_M)
put(ws, r, 3, 1.0, kind="total", fmt=FMT_PCT)
r += 2

section_bar(ws, r, "  AR BY CONTRACT  (USD M)", width=3); r += 1
note(ws, r, "Largest receivable balances by contract or customer. Add rows as needed."); r += 1
headers(ws, r, ["Contract", "Amount", "Share"]); r += 1
AR_CONTRACT = [("Ministry of Education — Framework", 2.10), ("Riyadh Schools Group", 1.50),
               ("Jeddah Private Academies", 1.20), ("Eastern Province Consortium", 0.80),
               ("Tracks — corporate accounts", 0.60), ("Other contracts", 0.50)]
ct_first = r
for contract, amt in AR_CONTRACT:
    put(ws, r, 1, contract, kind="input")
    put(ws, r, 2, amt,      kind="input", fmt=FMT_M)
    r += 1
ct_last = r - 1
for rr in range(ct_first, ct_last + 1):
    put(ws, rr, 3, f"=IF($B${r}=0,\"\",B{rr}/$B${r})", kind="calc", fmt=FMT_PCT)
put(ws, r, 1, "Total", kind="total", bold=True)
put(ws, r, 2, f"=SUM(B{ct_first}:B{ct_last})", kind="total", fmt=FMT_M)
put(ws, r, 3, 1.0, kind="total", fmt=FMT_PCT)

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — AP
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("AP")
widths(ws, {"A": 32, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16})

r = 1
section_bar(ws, r, "  ACCOUNTS PAYABLE — MONTHLY FLOW  (USD M)", width=6); r += 1
note(ws, r, "Drives the AP column chart. Enter DPO in days."); r += 1
headers(ws, r, ["Month", "Opening AP", "Purchases", "Payments", "Closing AP", "DPO (days)"]); r += 1

AP_MONTHLY = [
    ("Jan-26", 2.95, 2.00, 2.05, 43.5),
    ("Feb-26", 2.90, 1.95, 2.05, 43.1),
    ("Mar-26", 2.80, 2.00, 2.10, 40.5),
    ("Apr-26", 2.70, 2.10, 1.90, 41.4),
    ("May-26", 2.90, 1.95, 2.05, 43.1),
    ("Jun-26", 2.80, 2.05, 2.15, 39.5),
    ("Jul-26", 2.70, 2.00, 2.10, 39.0),
    ("Aug-26", 2.60, 1.90, 2.10, 37.9),
]
ap_first = r
for mo, op, pur, pay, dpo in AP_MONTHLY:
    put(ws, r, 1, mo,  kind="input")
    put(ws, r, 2, op,  kind="input", fmt=FMT_M)
    put(ws, r, 3, pur, kind="input", fmt=FMT_M)
    put(ws, r, 4, pay, kind="input", fmt=FMT_M)
    put(ws, r, 5, f"=B{r}+C{r}-D{r}", kind="calc", fmt=FMT_M)
    put(ws, r, 6, dpo, kind="input", fmt=FMT_NUM)
    r += 1
put(ws, r, 1, "YTD total", kind="total", bold=True)
put(ws, r, 2, "", kind="total")
put(ws, r, 3, f"=SUM(C{ap_first}:C{ap_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 4, f"=SUM(D{ap_first}:D{ap_first+5})", kind="total", fmt=FMT_M)
put(ws, r, 5, f"=E{ap_first+5}", kind="total", fmt=FMT_M)
put(ws, r, 6, f"=AVERAGE(F{ap_first}:F{ap_first+5})", kind="total", fmt=FMT_NUM)
r += 2

section_bar(ws, r, "  AP AGING  (USD M)", width=3); r += 1
note(ws, r, "Balance at the period end, split by age."); r += 1
headers(ws, r, ["AP Aging Bucket", "Amount", "Share"]); r += 1
AP_AGING = [("Current — due next 30 days", 1.5), ("31–60 days", 0.8), ("60+ days / overdue", 0.4)]
apg_first = r
for bucket, amt in AP_AGING:
    put(ws, r, 1, bucket, kind="plain")
    put(ws, r, 2, amt,    kind="input", fmt=FMT_M)
    r += 1
apg_last = r - 1
for rr in range(apg_first, apg_last + 1):
    put(ws, rr, 3, f"=IF($B${r}=0,\"\",B{rr}/$B${r})", kind="calc", fmt=FMT_PCT)
put(ws, r, 1, "Total AP", kind="total", bold=True)
put(ws, r, 2, f"=SUM(B{apg_first}:B{apg_last})", kind="total", fmt=FMT_M)
put(ws, r, 3, 1.0, kind="total", fmt=FMT_PCT)
r += 2

section_bar(ws, r, "  AP BY VENDOR  (USD M)", width=3); r += 1
note(ws, r, "Largest payable balances by vendor. Add rows as needed."); r += 1
headers(ws, r, ["Vendor", "Amount", "Share"]); r += 1
AP_VENDOR = [("AWS / cloud infrastructure", 0.62), ("Content production partners", 0.48),
             ("Facilities & office leases", 0.37), ("Marketing agencies", 0.29),
             ("Professional services", 0.21), ("Other vendors", 0.18)]
vn_first = r
for vendor, amt in AP_VENDOR:
    put(ws, r, 1, vendor, kind="input")
    put(ws, r, 2, amt,    kind="input", fmt=FMT_M)
    r += 1
vn_last = r - 1
for rr in range(vn_first, vn_last + 1):
    put(ws, rr, 3, f"=IF($B${r}=0,\"\",B{rr}/$B${r})", kind="calc", fmt=FMT_PCT)
put(ws, r, 1, "Total", kind="total", bold=True)
put(ws, r, 2, f"=SUM(B{vn_first}:B{vn_last})", kind="total", fmt=FMT_M)
put(ws, r, 3, 1.0, kind="total", fmt=FMT_PCT)

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — CASH
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Cash")
widths(ws, {"A": 34, "B": 15, "C": 34, "D": 15, "E": 34, "F": 14})

r = 1
section_bar(ws, r, "  CASH KPI TILES", width=5); r += 1
note(ws, r, "Four tiles, each shown twice — once for the month, once for YTD. The note is the small text under the number."); r += 1
headers(ws, r, ["Cash Tile", "Month value", "Month note", "YTD value", "YTD note"]); r += 1

TILES = [
    ("Cash balance",  4.08, "29.1 months runway",       4.08,  "opened the year at $4.80M"),
    ("Collections",   2.40, "Budget $2.50M invoiced",  13.35,  "on $14.35M invoiced YTD"),
    ("Net working capital", 3.20, "+$0.40M vs prior month", 1.40, "movement YTD · closing $2.50M"),
    ("Accounts receivable", 7.10, "88% collection rate",    7.10, "from $5.70M at 1 Jan"),
]
for name, mv, mn, yv, yn in TILES:
    put(ws, r, 1, name, kind="plain", bold=True)
    put(ws, r, 2, mv, kind="input", fmt=FMT_M)
    put(ws, r, 3, mn, kind="input")
    put(ws, r, 4, yv, kind="input", fmt=FMT_M)
    put(ws, r, 5, yn, kind="input")
    r += 1
r += 1

def bridge_block(ws, r, title, note_text, header, steps):
    section_bar(ws, r, title, width=4); r += 1
    note(ws, r, note_text); r += 1
    headers(ws, r, [header, "Type", "Amount", "Running balance"]); r += 1
    first = r
    for step, kind_, amt in steps:
        put(ws, r, 1, step, kind="input")
        put(ws, r, 2, kind_, kind="input", halign="center")
        put(ws, r, 3, amt,  kind="input", fmt=FMT_M)
        if r == first:
            put(ws, r, 4, f"=C{r}", kind="calc", fmt=FMT_M)
        else:
            put(ws, r, 4,
                f'=IF(B{r}="Closing",D{r-1},IF(B{r}="Inflow",D{r-1}+C{r},D{r-1}-C{r}))',
                kind="calc", fmt=FMT_M)
        r += 1
    return r + 1

BRIDGE_M = [
    ("Opening cash",    "Opening", 4.40),
    ("Collections",     "Inflow",  2.40),
    ("Other inflows",   "Inflow",  0.10),
    ("Operating costs", "Outflow", 2.08),
    ("Capex",           "Outflow", 0.16),
    ("Debt service",    "Outflow", 0.16),
    ("Other outflows",  "Outflow", 0.42),
    ("Closing cash",    "Closing", 4.08),
]
BRIDGE_Y = [
    ("Opening cash — 1 Jan 2026",  "Opening",  4.80),
    ("Collections",                "Inflow",  13.35),
    ("Other inflows",              "Inflow",   0.59),
    ("Operating costs",            "Outflow", 11.99),
    ("Capex",                      "Outflow",  0.89),
    ("Debt service",               "Outflow",  0.89),
    ("Other outflows",             "Outflow",  0.89),
    ("Closing cash — 30 Jun 2026", "Closing",  4.08),
]
r = bridge_block(ws, r, "  CASH BRIDGE — MONTH  (USD M)",
                 'Type must be one of: Opening, Inflow, Outflow, Closing. Enter every amount as POSITIVE.',
                 "Bridge Step (Month)", BRIDGE_M)
r = bridge_block(ws, r, "  CASH BRIDGE — YEAR TO DATE  (USD M)",
                 'Same rules as the monthly bridge. Opening is the 1 January balance.',
                 "Bridge Step (YTD)", BRIDGE_Y)

section_bar(ws, r, "  MONTHLY CLOSING CASH  (USD M)", width=4); r += 1
note(ws, r, "Drives the cash balance line chart. Cover as many months as you want to show — it can run wider than the P&L months."); r += 1
note(ws, r, "Forecast = Y draws a dashed line. Illustrative = Y draws a hollow marker and shows the 'indicative' flag."); r += 1
headers(ws, r, ["Cash Month", "Closing cash", "Forecast?", "Illustrative?"]); r += 1

CASH_MONTHLY = [
    ("Jul-25", 5.35, "No",  "Yes"), ("Aug-25", 5.22, "No",  "Yes"),
    ("Sep-25", 5.10, "No",  "Yes"), ("Oct-25", 5.02, "No",  "Yes"),
    ("Nov-25", 4.91, "No",  "Yes"), ("Dec-25", 4.80, "No",  "No"),
    ("Jan-26", 4.62, "No",  "No"),  ("Feb-26", 4.55, "No",  "No"),
    ("Mar-26", 4.49, "No",  "No"),  ("Apr-26", 4.44, "No",  "No"),
    ("May-26", 4.40, "No",  "No"),  ("Jun-26", 4.08, "No",  "No"),
    ("Jul-26", 3.95, "Yes", "No"),  ("Aug-26", 3.86, "Yes", "No"),
    ("Sep-26", 3.80, "Yes", "No"),  ("Oct-26", 3.72, "Yes", "No"),
    ("Nov-26", 3.66, "Yes", "No"),  ("Dec-26", 3.58, "Yes", "No"),
]
for mo, val, fc, ill in CASH_MONTHLY:
    put(ws, r, 1, mo,  kind="input")
    put(ws, r, 2, val, kind="input", fmt=FMT_M)
    put(ws, r, 3, fc,  kind="input", halign="center")
    put(ws, r, 4, ill, kind="input", halign="center")
    r += 1
r += 1

section_bar(ws, r, "  RUNWAY METRICS", width=3); r += 1
note(ws, r, "Burn figures are negative when cash is being consumed."); r += 1
headers(ws, r, ["Runway Metric", "Value", "Notes"]); r += 1
RUNWAY = [
    ("Monthly cash burn",            -0.32, "Net cash movement in the latest month (USD M)"),
    ("Trailing 3-month average burn", -0.14, "Average of the last three months (USD M)"),
    ("YTD net cash movement",        -0.72, "Closing cash less opening cash (USD M)"),
    ("Cash runway (months)",         29.14, "Cash ÷ 3-month average burn"),
]
for metric, val, nt in RUNWAY:
    put(ws, r, 1, metric, kind="plain")
    put(ws, r, 2, val, kind="input", fmt=FMT_M if "months" not in metric else FMT_NUM)
    put(ws, r, 3, nt, kind="plain")
    r += 1

ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════════
# TAB 9 — KEY UPDATES
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Key Updates")
widths(ws, {"A": 6, "B": 24, "C": 105})

r = 1
section_bar(ws, r, "  KEY NARRATIVE UPDATES", width=3); r += 1
note(ws, r, "One row per commentary point, shown in order at the bottom of the dashboard. Add or remove rows freely."); r += 1
headers(ws, r, ["#", "Topic", "Commentary"]); r += 1

UPDATES = [
    ("Revenue",         "June revenue of $2.50M landed at 94.3% of budget, missing plan in five of the first six months of the year. YTD revenue of $14.35M is $0.90M (5.9%) behind budget."),
    ("Margin",          "EBITDA of $0.36M gave a 14.4% margin against a 17.4% budget. YTD EBITDA of $2.16M is $0.44M behind plan; the gap is driven by revenue, not cost overrun."),
    ("Costs",           "Cost control is holding — total costs ran at 97.7% of budget for the month and 96.4% YTD. The underspend has absorbed a meaningful share of the revenue shortfall."),
    ("Collections",     "The June collection rate of 96.0% is the strongest month of the year so far. AR stands at $6.70M, of which $1.90M (28.4%) is now 60+ days overdue."),
    ("Working capital", "Net working capital has absorbed $1.40M of cash since 1 January, almost entirely through the AR build. Payables have been drawn down $0.20M over the same period."),
    ("Cash",            "Cash closed at $4.08M, down $0.72M since 1 January. At the trailing three-month burn of $0.14M per month that is 29.1 months of runway, before the $1.50M current portion of Facility A."),
]
for i, (topic, text) in enumerate(UPDATES, 1):
    put(ws, r, 1, i, kind="plain", halign="center")
    put(ws, r, 2, topic, kind="input")
    put(ws, r, 3, text,  kind="input", wrap=True)
    ws.row_dimensions[r].height = 46
    r += 1

ws.sheet_view.showGridLines = False

wb.save(args.out)
print(f"✓ Created {args.out}")
print(f"  Tabs: {', '.join(wb.sheetnames)}")
