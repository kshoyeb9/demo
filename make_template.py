"""
make_template.py — Generates pl_only_final.xlsx pre-filled with dummy data.
Run once to create the template, then update the yellow cells each month.

Usage:
    python make_template.py
"""

from pathlib import Path
import openpyxl
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              numbers as xl_numbers)
from openpyxl.utils import get_column_letter

OUT = "pl_only_final.xlsx"

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "P&L Performance"

# ── Colour palette ────────────────────────────────────────────────────────────
C_DARK   = "11203A"   # Noon dark blue
C_GREEN  = "17E4A1"   # Noon green
C_AMBER  = "FFF3CD"   # input cell highlight
C_HEADER = "1F3864"   # section header
C_SUBHDR = "D9E1F2"   # column header row
C_TOTAL  = "BDD7EE"   # total rows
C_WHITE  = "FFFFFF"
C_LIGHT  = "F2F2F2"

FMT_2D  = '#,##0.00'
FMT_PCT = '0.0%'
FMT_INT = '#,##0'

def fill(hex_): return PatternFill("solid", fgColor=hex_)
def font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size)
def align(h="left", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)
def thin_border():
    s = Side(style="thin")
    return Border(bottom=s)

def section_header(row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font   = Font(bold=True, color=C_WHITE, size=10)
    c.fill   = fill(C_HEADER)
    c.alignment = align("left")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=28)

def col_headers(row, labels, start_col=1):
    for i, txt in enumerate(labels):
        c = ws.cell(row=row, column=start_col+i, value=txt)
        c.font      = font(bold=True, size=9)
        c.fill      = fill(C_SUBHDR)
        c.alignment = align("center")

def data_row(row, values, start_col=1, total=False, input_=False):
    bg = C_TOTAL if total else (C_AMBER if input_ else C_WHITE)
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col+i, value=v)
        c.fill = fill(bg)
        if isinstance(v, float):
            if abs(v) < 5:
                c.number_format = FMT_2D
            else:
                c.number_format = FMT_2D
        c.alignment = align("right" if isinstance(v,(int,float)) else "left")
        if total:
            c.font = font(bold=True)

def pct_row(row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col+i, value=v)
        c.fill = fill(C_LIGHT)
        if isinstance(v, float):
            c.number_format = FMT_PCT
        c.alignment = align("right" if isinstance(v,(int,float)) else "left")
        c.font = font(size=9)

# ── Column widths ─────────────────────────────────────────────────────────────
ws.column_dimensions["A"].width = 36
for col in range(2, 30):
    ws.column_dimensions[get_column_letter(col)].width = 11

# ═══════════════════════════════════════════════════════════════════════════════
# TITLE ROW 1
# ═══════════════════════════════════════════════════════════════════════════════
c = ws.cell(row=1, column=1, value="Noon Academy — P&L Performance Dashboard")
c.font = Font(bold=True, size=14, color=C_DARK)
ws.merge_cells("A1:Z1")

c = ws.cell(row=2, column=1, value="Update yellow cells each month, then run: python etl_pl.py --file pl_only_final.xlsx")
c.font = Font(size=9, color="595959", italic=True)
ws.merge_cells("A2:Z2")

ws.row_dimensions[3].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 5 — PERIOD HEADERS  (etl reads col 10 and col 20)
# ═══════════════════════════════════════════════════════════════════════════════
ws.cell(row=4, column=1,  value="▶ Update period label below (cell K5 and U5) each month")
ws.cell(row=4, column=1).font = Font(size=9, italic=True, color="595959")

# K5 = col 11 in openpyxl (1-indexed), but etl reads col index 10 (0-indexed) → openpyxl col 11
ws.cell(row=5, column=11, value="Key Metrics — June 2026  (Month)").fill = fill(C_AMBER)
ws.cell(row=5, column=21, value="Key Metrics — Year-to-Date  (Jan–Jun 2026)").fill = fill(C_AMBER)
ws.cell(row=5, column=11).font = font(bold=True, size=10)
ws.cell(row=5, column=21).font = font(bold=True, size=10)
ws.merge_cells(start_row=5, start_column=11, end_row=5, end_column=19)
ws.merge_cells(start_row=5, start_column=21, end_row=5, end_column=28)

ws.row_dimensions[6].height = 6

# ═══════════════════════════════════════════════════════════════════════════════
# ROWS 7-13 — column headers for P&L sections
# ═══════════════════════════════════════════════════════════════════════════════
# Two blocks: Month (cols B-I = cols 2-9) and YTD (cols L-S = cols 12-19)
ws.cell(row=7, column=2,  value="← MONTH (latest period) →").font  = Font(bold=True, size=9, color=C_HEADER)
ws.cell(row=7, column=12, value="← YEAR-TO-DATE →").font           = Font(bold=True, size=9, color=C_HEADER)

HDR = ["Actual", "Budget", "Variance", "% Bdgt"]
col_headers(8, [""] + HDR + ["",""] + HDR, start_col=1)
# Col A blank, B-E = Month, F-G blank, H-K blank → adjust:
# etl reads: col 1=act_mo, col 2=bgt_mo, col 5=act_ytd, col 6=bgt_ytd
# So: col B(2)=Act_Mo, col C(3)=Bgt_Mo, col D(4)=Var_Mo, col E(5)=Pct_Mo
#          col F(6)=Act_YTD, col G(7)=Bgt_YTD, col H(8)=Var_YTD, col I(9)=Pct_YTD
# Wait - etl uses 0-indexed from the tuple: col0=A, col1=B, col2=C, col5=F, col6=G
# 0-idx: 0=A 1=B 2=C 3=D 4=E 5=F 6=G 7=H 8=I
# So headers:
#  A=name, B=Act_Mo, C=Bgt_Mo, D=Var_Mo, E=Pct_Mo, F=Act_YTD, G=Bgt_YTD, H=Var_YTD, I=Pct_YTD

col_headers(8, ["Description",
                "Actual (Mo)","Budget (Mo)","Var (Mo)","% Bgt (Mo)",
                "Actual YTD", "Budget YTD", "Var YTD",  "% Bgt YTD"], start_col=1)

ws.row_dimensions[9].height  = 6

# ═══════════════════════════════════════════════════════════════════════════════
# REVENUE BY BU  — rows 15–20 (etl starts scanning at row 15)
# ═══════════════════════════════════════════════════════════════════════════════
# Rows 10-13: section title
section_header(10, "SECTION 1 — REVENUE BY BUSINESS UNIT  (USD Millions)")
ws.cell(row=11, column=1, value="Update yellow cells: Actual (Mo) in col B and Actual YTD in col F for each BU")
ws.cell(row=11, column=1).font = Font(size=9, italic=True)

col_headers(13, ["Business Unit",
                 "Actual (Mo)","Budget (Mo)","Var (Mo)","% Bgt (Mo)",
                 "Actual YTD", "Budget YTD", "Var YTD",  "% Bgt YTD"], start_col=1)
ws.row_dimensions[14].height = 4

REV_BUS = [
    # name,         act_mo, bgt_mo, act_ytd, bgt_ytd
    ("Tracks",              0.99, 1.07,  5.65,  6.17),
    ("B2B",                 0.63, 0.63,  3.64,  3.72),
    ("Govt Schools — Legacy",0.52, 0.56,  2.97,  3.20),
    ("Govt Schools — New",  0.26, 0.29,  1.41,  1.56),
    ("Out of School",       0.10, 0.10,  0.68,  0.60),
]
for i, (name, am, bm, ay, by) in enumerate(REV_BUS):
    r = 15 + i
    var_m = round(am-bm, 4); pct_m = round(am/bm, 4) if bm else None
    var_y = round(ay-by, 4); pct_y = round(ay/by, 4) if by else None
    data_row(r, [name, am, bm, var_m, pct_m, ay, by, var_y, pct_y], input_=True)
    # formula cells for Var and % (cols D,E,H,I)
    ws.cell(row=r, column=4).value  = f"=B{r}-C{r}"
    ws.cell(row=r, column=5).value  = f"=IF(C{r},B{r}/C{r},\"\")"
    ws.cell(row=r, column=5).number_format = FMT_PCT
    ws.cell(row=r, column=9).value  = f"=F{r}-G{r}"
    ws.cell(row=r, column=10).value = f"=IF(G{r},F{r}/G{r},\"\")"
    ws.cell(row=r, column=10).number_format = FMT_PCT

# Total Revenue — row 20
r = 20
data_row(r, ["Total Revenue",
             f"=SUM(B15:B19)", f"=SUM(C15:C19)",
             f"=B{r}-C{r}",   f"=IF(C{r},B{r}/C{r},\"\")",
             f"=SUM(F15:F19)", f"=SUM(G15:G19)",
             f"=F{r}-G{r}",   f"=IF(G{r},F{r}/G{r},\"\")"], total=True)
for col in [5, 10]: ws.cell(row=r, column=col).number_format = FMT_PCT

ws.row_dimensions[21].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# COSTS — rows 24–29
# ═══════════════════════════════════════════════════════════════════════════════
section_header(22, "SECTION 2 — COSTS BY CATEGORY  (USD Millions)")
col_headers(23, ["Cost Category",
                 "Actual (Mo)","Budget (Mo)","Var (Mo)","% Bgt (Mo)",
                 "Actual YTD", "Budget YTD", "Var YTD",  "% Bgt YTD"], start_col=1)

COSTS_DATA = [
    ("Direct costs",              0.89, 0.92,  5.06,  5.28),
    ("Marketing",                 0.10, 0.13,  0.67,  0.82),
    ("BU salaries",               0.68, 0.70,  3.87,  3.94),
    ("Noon HQ",                   0.21, 0.20,  1.18,  1.20),
    ("Other operating expenses",  0.26, 0.24,  1.41,  1.41),
]
for i, (name, am, bm, ay, by) in enumerate(COSTS_DATA):
    r = 24 + i
    data_row(r, [name, am, bm, round(am-bm,4), None, ay, by, round(ay-by,4), None], input_=True)
    ws.cell(row=r, column=4).value  = f"=B{r}-C{r}"
    ws.cell(row=r, column=5).value  = f"=IF(C{r},B{r}/C{r},\"\")"
    ws.cell(row=r, column=5).number_format = FMT_PCT
    ws.cell(row=r, column=9).value  = f"=F{r}-G{r}"
    ws.cell(row=r, column=10).value = f"=IF(G{r},F{r}/G{r},\"\")"
    ws.cell(row=r, column=10).number_format = FMT_PCT

r = 29
data_row(r, ["Total costs",
             "=SUM(B24:B28)", "=SUM(C24:C28)",
             f"=B{r}-C{r}", f"=IF(C{r},B{r}/C{r},\"\")",
             "=SUM(F24:F28)", "=SUM(G24:G28)",
             f"=F{r}-G{r}", f"=IF(G{r},F{r}/G{r},\"\")"], total=True)
for col in [5, 10]: ws.cell(row=r, column=col).number_format = FMT_PCT

ws.row_dimensions[30].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# P&L SUMMARY — rows 34–41
# etl reads: row 34+ with parse_avb, then margin % from rows 40-41
# ═══════════════════════════════════════════════════════════════════════════════
section_header(31, "SECTION 3 — P&L SUMMARY  (USD Millions)")
col_headers(32, ["P&L Line",
                 "Actual (Mo)","Budget (Mo)","Var (Mo)","% Bgt (Mo)",
                 "Actual YTD", "Budget YTD", "Var YTD",  "% Bgt YTD"], start_col=1)
ws.row_dimensions[33].height = 4

PL_ROWS = [
    # name, act_mo, bgt_mo, act_ytd, bgt_ytd, is_total
    ("Revenue",                    2.50, 2.65, 14.35, 15.25, True),
    ("Direct / operator costs",    0.89, 0.92,  5.06,  5.28, False),
    ("Gross profit / contribution", 1.61, 1.73,  9.29,  9.97, True),
    ("Operating expenses (staff, facilities, mktg, central)", 0.89, 0.93, 5.06, 5.37, False),
    ("EBITDA",                     0.36, 0.46,  2.16,  2.60, True),
]
for i, (name, am, bm, ay, by, tot) in enumerate(PL_ROWS):
    r = 34 + i
    data_row(r, [name, am, bm, round(am-bm,4), None, ay, by, round(ay-by,4), None],
             total=tot, input_=(not tot))
    ws.cell(row=r, column=4).value  = f"=B{r}-C{r}"
    ws.cell(row=r, column=5).value  = f"=IF(C{r},B{r}/C{r},\"\")"
    ws.cell(row=r, column=5).number_format = FMT_PCT
    ws.cell(row=r, column=9).value  = f"=F{r}-G{r}"
    ws.cell(row=r, column=10).value = f"=IF(G{r},F{r}/G{r},\"\")"
    ws.cell(row=r, column=10).number_format = FMT_PCT

ws.row_dimensions[39].height = 4

# Row 40: Gross margin %  (etl reads col 1 and col 5)
ws.cell(row=40, column=1, value="Gross profit margin %").fill = fill(C_LIGHT)
ws.cell(row=40, column=1).font = font(size=9)
ws.cell(row=40, column=2, value="=B36/B34").number_format = FMT_PCT; ws.cell(row=40,column=2).fill=fill(C_LIGHT)
ws.cell(row=40, column=6, value="=F36/F34").number_format = FMT_PCT; ws.cell(row=40,column=6).fill=fill(C_LIGHT)

# Row 41: EBITDA margin %
ws.cell(row=41, column=1, value="EBITDA margin %").fill = fill(C_LIGHT)
ws.cell(row=41, column=1).font = font(size=9)
ws.cell(row=41, column=2, value="=B38/B34").number_format = FMT_PCT; ws.cell(row=41,column=2).fill=fill(C_LIGHT)
ws.cell(row=41, column=6, value="=F38/F34").number_format = FMT_PCT; ws.cell(row=41,column=6).fill=fill(C_LIGHT)

ws.row_dimensions[42].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# WORKING CAPITAL MONTHLY — rows 62–73
# etl reads starting at row 62, cols 0-6
# ═══════════════════════════════════════════════════════════════════════════════
section_header(58, "SECTION 4 — WORKING CAPITAL MONTHLY  (USD Millions)")
ws.cell(row=59, column=1, value="Months after the current period are treated as forecast (greyed in dashboard)")
ws.cell(row=59, column=1).font = Font(size=9, italic=True)
col_headers(60, ["Month","AR","AP","Deferred Rev","Other","Net WC","Movement"], start_col=1)
ws.row_dimensions[61].height = 4

WC_MO = [
    ("Jan-26", 5.9,-2.9,-2.4, 0.5, 1.1, 0.0),
    ("Feb-26", 6.1,-2.8,-2.3, 0.5, 1.5, 0.4),
    ("Mar-26", 6.3,-2.7,-2.2, 0.6, 2.0, 0.5),
    ("Apr-26", 6.4,-2.9,-2.2, 0.5, 1.8,-0.2),
    ("May-26", 6.6,-2.8,-2.1, 0.6, 2.3, 0.5),
    ("Jun-26", 6.7,-2.7,-2.1, 0.6, 2.5, 0.2),
    ("Jul-26", 6.8,-2.6,-2.0, 0.6, 2.8, 0.3),
    ("Aug-26", 7.1,-2.4,-1.8, 0.3, 3.2, 0.4),
]
for i, row_data in enumerate(WC_MO):
    data_row(62+i, list(row_data), input_=True)

ws.row_dimensions[70].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# WORKING CAPITAL YTD — rows 74–78
# etl reads starting at row 74, cols 0-4
# ═══════════════════════════════════════════════════════════════════════════════
section_header(71, "SECTION 5 — WORKING CAPITAL YTD MOVEMENT  (USD Millions)")
col_headers(72, ["WC Item","At 1 Jan","At Period End","Movement","Cash Impact"], start_col=1)
ws.row_dimensions[73].height = 4

WC_YTD = [
    ("Accounts receivable",       5.9, 6.7, 0.8,-0.8),
    ("Accounts payable",         -2.9,-2.7, 0.2,-0.2),
    ("Deferred revenue",         -2.4,-2.1, 0.3,-0.3),
    ("Other working capital items", 0.5, 0.6, 0.1,-0.1),
    ("Net working capital",       1.1, 2.5, 1.4,-1.4),
]
for i, row_data in enumerate(WC_YTD):
    tot = "Net working capital" in row_data[0]
    data_row(74+i, list(row_data), total=tot, input_=(not tot))

ws.row_dimensions[79].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# AR MONTHLY — rows 84–91, YTD row 92
# etl reads starting at row 84, cols 0-5
# ═══════════════════════════════════════════════════════════════════════════════
section_header(80, "SECTION 6 — ACCOUNTS RECEIVABLE MONTHLY  (USD Millions)")
col_headers(82, ["Month","Opening AR","Invoiced","Collected","Closing AR","Collection Rate"], start_col=1)
ws.row_dimensions[83].height = 4

AR_MO = [
    ("Jan-26", 5.70, 2.30, 2.10, 5.90, 0.913),
    ("Feb-26", 5.90, 2.25, 2.05, 6.10, 0.911),
    ("Mar-26", 6.10, 2.55, 2.35, 6.30, 0.922),
    ("Apr-26", 6.30, 2.40, 2.30, 6.40, 0.958),
    ("May-26", 6.40, 2.35, 2.15, 6.60, 0.915),
    ("Jun-26", 6.60, 2.50, 2.40, 6.70, 0.960),
    ("Jul-26", 6.70, 2.55, 2.45, 6.80, 0.961),
    ("Aug-26", 6.80, 2.60, 2.30, 7.10, 0.885),
]
for i, (mo, op, inv, col_, cl, rate) in enumerate(AR_MO):
    r = 84+i
    data_row(r, [mo, op, inv, col_, cl, rate], input_=True)
    ws.cell(row=r, column=6).number_format = FMT_PCT
    ws.cell(row=r, column=5).value = f"=B{r}+C{r}-D{r}"

# AR YTD row 92  (etl scans r_n to r_n+3 for "ytd" label)
data_row(92, ["YTD Total", None, "=SUM(C84:C89)", "=SUM(D84:D89)", None,
              "=IF(C92,D92/C92,\"\")"], total=True)
ws.cell(row=92, column=6).number_format = FMT_PCT

ws.row_dimensions[93].height = 8

# AR Aging — rows 96–98, Total row 99
section_header(94, "SECTION 7 — AR AGING  (USD Millions)")
col_headers(95, ["Aging Bucket","Amount","Share %"], start_col=1)

AR_AGING = [
    ("Current — due next 30 days", 2.8, 0.4179),
    ("31–60 days",                 2.0, 0.2985),
    ("60+ days / overdue",         1.9, 0.2836),
]
for i, (bk, amt, sh) in enumerate(AR_AGING):
    data_row(96+i, [bk, amt, sh], input_=True)
    ws.cell(row=96+i, column=3).number_format = FMT_PCT

data_row(99, ["Total AR", "=SUM(B96:B98)", "=B99/B99"], total=True)
ws.cell(row=99, column=3).number_format = FMT_PCT
ws.row_dimensions[100].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# AP MONTHLY — rows 105–112, YTD row 113
# etl reads starting at row 105, cols 0-5
# ═══════════════════════════════════════════════════════════════════════════════
section_header(101, "SECTION 8 — ACCOUNTS PAYABLE MONTHLY  (USD Millions)")
col_headers(103, ["Month","Opening AP","Purchases","Payments","Closing AP","DPO (days)"], start_col=1)
ws.row_dimensions[104].height = 4

AP_MO = [
    ("Jan-26", 2.95, 2.00, 2.05, 2.90, 43.5),
    ("Feb-26", 2.90, 1.95, 2.05, 2.80, 43.1),
    ("Mar-26", 2.80, 2.00, 2.10, 2.70, 40.5),
    ("Apr-26", 2.70, 2.10, 1.90, 2.90, 41.4),
    ("May-26", 2.90, 1.95, 2.05, 2.80, 43.1),
    ("Jun-26", 2.80, 2.05, 2.15, 2.70, 39.5),
    ("Jul-26", 2.70, 2.00, 2.10, 2.60, 39.0),
    ("Aug-26", 2.60, 1.90, 2.10, 2.40, 37.9),
]
for i, (mo, op, pur, pay, cl, dpo) in enumerate(AP_MO):
    r = 105+i
    data_row(r, [mo, op, pur, pay, cl, dpo], input_=True)
    ws.cell(row=r, column=5).value = f"=B{r}+C{r}-D{r}"

data_row(113, ["YTD Total", None, "=SUM(C105:C110)", "=SUM(D105:D110)", None,
               "=AVERAGE(F105:F110)"], total=True)
ws.cell(row=113, column=6).number_format = FMT_2D

ws.row_dimensions[114].height = 8

# AP Aging — rows 118–120, Total row 121
section_header(115, "SECTION 9 — AP AGING  (USD Millions)")
col_headers(116, ["Aging Bucket","Amount","Share %"], start_col=1)
ws.row_dimensions[117].height = 4

AP_AGING = [
    ("Current — due next 30 days", 1.5, 0.5556),
    ("31–60 days",                 0.8, 0.2963),
    ("60+ days / overdue",         0.4, 0.1481),
]
for i, (bk, amt, sh) in enumerate(AP_AGING):
    data_row(118+i, [bk, amt, sh], input_=True)
    ws.cell(row=118+i, column=3).number_format = FMT_PCT

data_row(121, ["Total AP", "=SUM(B118:B120)", "=B121/B121"], total=True)
ws.cell(row=121, column=3).number_format = FMT_PCT
ws.row_dimensions[122].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# CASH TILES — etl reads specific cell positions
# R135 col10=cash_val_m, col15=coll_val_m, col20=cash_val_y, col25=coll_val_y
# R136 col10=cash_note_m, col15=coll_note_m, col20=cash_note_y, col25=coll_note_y
# R138 col10=nwc_val_m,  col15=ar_val_m,   col20=nwc_val_y,  col25=ar_val_y
# R139 col10=nwc_note_m, col15=ar_note_m,  col20=nwc_note_y, col25=ar_note_y
# (col indices are 0-based → openpyxl column = index+1)
# ═══════════════════════════════════════════════════════════════════════════════
section_header(123, "SECTION 10 — CASH DASHBOARD TILES  (USD Millions)")
ws.cell(row=124, column=1, value="Each tile: value row then note row (text). Columns K,P = month tiles; U,Z = YTD tiles.")
ws.cell(row=124, column=1).font = Font(size=9, italic=True)

ws.cell(row=126, column=11, value="← CASH BALANCE (Mo) →").font = font(bold=True)
ws.cell(row=126, column=16, value="← COLLECTIONS (Mo) →").font = font(bold=True)
ws.cell(row=126, column=21, value="← CASH BALANCE (YTD) →").font = font(bold=True)
ws.cell(row=126, column=26, value="← COLLECTIONS (YTD) →").font = font(bold=True)

ws.cell(row=127, column=11, value="Label").fill = fill(C_SUBHDR); ws.cell(row=127,column=11).font=font(bold=True)
ws.cell(row=127, column=16, value="Label").fill = fill(C_SUBHDR); ws.cell(row=127,column=16).font=font(bold=True)
ws.cell(row=127, column=21, value="Label").fill = fill(C_SUBHDR); ws.cell(row=127,column=21).font=font(bold=True)
ws.cell(row=127, column=26, value="Label").fill = fill(C_SUBHDR); ws.cell(row=127,column=26).font=font(bold=True)

ws.row_dimensions[128].height = 4

# Cash tile values — row 135 (col index 10 = openpyxl col 11)
#                                             col idx: 10        15          20         25
ws.cell(row=129, column=11, value="NWC (Mo)").fill = fill(C_SUBHDR); ws.cell(row=129,column=11).font=font(bold=True)
ws.cell(row=129, column=16, value="AR (Mo)").fill  = fill(C_SUBHDR); ws.cell(row=129,column=16).font=font(bold=True)
ws.cell(row=129, column=21, value="NWC (YTD)").fill = fill(C_SUBHDR); ws.cell(row=129,column=21).font=font(bold=True)
ws.cell(row=129, column=26, value="AR (YTD)").fill  = fill(C_SUBHDR); ws.cell(row=129,column=26).font=font(bold=True)

ws.row_dimensions[130].height = 8

# Row 135: cash tile VALUE row
def cash_cell(row, col_0idx, val, note=None):
    c = ws.cell(row=row, column=col_0idx+1, value=val)
    c.fill = fill(C_AMBER); c.number_format = FMT_2D; c.font=font(bold=True)
    if note is not None:
        n = ws.cell(row=row+1, column=col_0idx+1, value=note)
        n.fill = fill(C_AMBER); n.font=Font(size=9, italic=True)

# 2-row group labels
ws.cell(row=133, column=1, value="► Cash tile section — edit values and notes in yellow cells").font=Font(size=9,italic=True)
ws.cell(row=134, column=1, value="Row 135 = value / Row 136 = note / Row 138 = value / Row 139 = note").font=Font(size=9,color="595959",italic=True)

cash_cell(135, 10, 4.08, "29.1 months runway")          # cash_val_m, cash_note_m
cash_cell(135, 15, 2.40, "Budget $2.50M invoiced")       # coll_val_m, coll_note_m
cash_cell(135, 20, 4.08, "opened the year at $4.80M")    # cash_val_y, cash_note_y
cash_cell(135, 25, 13.35,"on $14.35M invoiced YTD")      # coll_val_y, coll_note_y

cash_cell(138, 10, 3.20, "+$0.40M vs prior month")       # nwc_val_m, nwc_note_m
cash_cell(138, 15, 7.10, "88% collection rate")          # ar_val_m,  ar_note_m
cash_cell(138, 20, 1.40, "movement YTD · closing $2.50M")# nwc_val_y, nwc_note_y
cash_cell(138, 25, 7.10, "from $5.70M at 1 Jan")         # ar_val_y,  ar_note_y

ws.row_dimensions[140].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# CASH BRIDGE MONTH — rows 143–150
# etl reads: col0=name, col2=running, col4=total_value, col5=inflow, col6=outflow
# ═══════════════════════════════════════════════════════════════════════════════
section_header(141, "SECTION 11 — CASH BRIDGE (MONTH)  (USD Millions)")
col_headers(142, ["Step","▶ Running after","(ignore)","(ignore)","▶ Total (open/close)","▶ Inflow","▶ Outflow"], start_col=1)

# Bridge: col A=name, col C(idx2)=running, col E(idx4)=total, col F(idx5)=inflow, col G(idx6)=outflow
BRIDGE_MO = [
    # name,              total, inflow, outflow, running
    ("Opening cash",     4.40,  0,     0,       4.40),
    ("Collections",      0,     2.40,  0,       6.80),
    ("Other inflows",    0,     0.10,  0,       6.90),
    ("Operating costs",  0,     0,     2.08,    4.82),
    ("Capex",            0,     0,     0.16,    4.66),
    ("Debt service",     0,     0,     0.16,    4.50),
    ("Other outflows",   0,     0,     0.42,    4.08),
    ("Closing cash",     4.08,  0,     0,       4.08),
]
for i, (nm, tot, inf, out, run) in enumerate(BRIDGE_MO):
    r = 143+i
    is_tot = tot > 0
    data_row(r, [nm, run, None, None, tot if is_tot else None,
                 inf if inf > 0 else None, out if out > 0 else None],
             total=is_tot, input_=True)

ws.row_dimensions[151].height = 8

# CASH BRIDGE YTD — rows 154–161
section_header(152, "SECTION 12 — CASH BRIDGE (YEAR-TO-DATE)  (USD Millions)")
col_headers(153, ["Step","▶ Running after","(ignore)","(ignore)","▶ Total (open/close)","▶ Inflow","▶ Outflow"], start_col=1)

BRIDGE_YTD = [
    ("Opening cash — 1 Jan 2026",    4.80,  0,     0,      4.80),
    ("Collections",                  0,    13.35,  0,     18.15),
    ("Other inflows",                0,     0.59,  0,     18.74),
    ("Operating costs",              0,     0,    11.99,   6.75),
    ("Capex",                        0,     0,     0.89,   5.86),
    ("Debt service",                 0,     0,     0.89,   4.97),
    ("Other outflows",               0,     0,     0.89,   4.08),
    ("Closing cash — 30 Jun 2026",   4.08,  0,     0,      4.08),
]
for i, (nm, tot, inf, out, run) in enumerate(BRIDGE_YTD):
    r = 154+i
    is_tot = tot > 0
    data_row(r, [nm, run, None, None, tot if is_tot else None,
                 inf if inf > 0 else None, out if out > 0 else None],
             total=is_tot, input_=True)

ws.row_dimensions[162].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# RUNWAY — rows 166–169  (etl reads col 1 for each value)
# ═══════════════════════════════════════════════════════════════════════════════
section_header(163, "SECTION 13 — CASH RUNWAY METRICS")
col_headers(165, ["Metric","Value"], start_col=1)

RUNWAY = [
    ("Monthly cash burn (latest month, USD M)", -0.32),
    ("Trailing 3-month average burn (USD M)",   -0.14),
    ("YTD net cash movement (USD M)",           -0.72),
    ("Cash runway (months)",                    29.14),
]
for i, (label, val) in enumerate(RUNWAY):
    data_row(166+i, [label, val], input_=True)

ws.row_dimensions[170].height = 8

# ═══════════════════════════════════════════════════════════════════════════════
# KEY UPDATES — rows 174+  (etl reads col0=num, col1=topic, col3=text)
# ═══════════════════════════════════════════════════════════════════════════════
section_header(171, "SECTION 14 — KEY NARRATIVE UPDATES")
col_headers(173, ["#","Topic","(blank)","Update text (plain paragraph)"], start_col=1)
ws.column_dimensions["D"].width = 80
ws.row_dimensions[172].height = 4

UPDATES = [
    (1, "Revenue",        "June revenue of $2.50M landed at 94.3% of budget, missing plan in five of the first six months of the year. YTD revenue of $14.35M is $0.90M (5.9%) behind budget."),
    (2, "Margin",         "EBITDA of $0.36M gave a 14.4% margin against a 17.4% budget. YTD EBITDA of $2.16M is $0.44M behind plan; the gap is driven by revenue, not cost overrun."),
    (3, "Costs",          "Cost control is holding — total costs ran at 97.7% of budget for the month and 96.4% YTD. The underspend has absorbed a meaningful share of the revenue shortfall."),
    (4, "Collections",    "The June collection rate of 96.0% is the strongest month of the year so far. AR stands at $6.70M, of which $1.90M (28.4%) is now 60+ days overdue."),
    (5, "Working capital","Net working capital has absorbed $1.40M of cash since 1 January, almost entirely through the AR build. Payables have been drawn down $0.20M over the same period."),
    (6, "Cash",           "Cash closed at $4.08M, down $0.72M since 1 January. At the trailing three-month burn of $0.14M per month that is 29.1 months of runway, before the $1.50M current portion of Facility A."),
]
for i, (num, topic, text) in enumerate(UPDATES):
    r = 174+i
    ws.cell(row=r, column=1, value=num)
    ws.cell(row=r, column=2, value=topic).fill = fill(C_AMBER)
    ws.cell(row=r, column=4, value=text).fill  = fill(C_AMBER)
    ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 40

# ── Freeze panes & print setup ────────────────────────────────────────────────
ws.freeze_panes = "B9"
ws.sheet_properties.tabColor = C_DARK

wb.save(OUT)
print(f"✓ Created {OUT}")
print("  Open in Excel, update the yellow cells, then run:")
print(f"  python etl_pl.py --file {OUT}")
