"""etl.py — Noon Academy Financials ETL
Reads the financial model and CF tracker Excel files, then writes fy_data.json.

Usage (from the noon_dashboard folder):
    python etl.py --model "path/to/Noon_financial_model_2_0_v42.xlsx" \
                  --cf    "path/to/Weekly_Report_CF_2026_Final_.xlsx"

Defaults assume the files are in the same folder as this script.
"""

import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl not installed. Run:  pip install openpyxl")

# ── BU tab configuration ──────────────────────────────────────────────────────
# sheet_name -> revenue label(s) to try in order
BU_TABS = {
    "Tracks":            ("03_BU_Tracks",            ["Net revenue",  "Net Revenue"]),
    "Government Schools":("04_BU_Government Schools", ["Net revenue",  "Net Revenue"]),
    "B2B":               ("05_BU_B2B",               ["Net Revenue (NGOs and MCIT)", "Net Revenue"]),
    "KSA B2C":           ("06_BU_KSA_B2C",           ["Total Revenue", "Net revenue"]),
    "Global Non-Saudi":  ("07_Global_Non_Saudi BU",  ["Total Revenue", "Net revenue"]),
    "Out of School":     ("06_BU_Out_of_School",     ["Net revenues",  "Net revenue"]),
}

GP_LABELS   = ["Gross profit", "Gross Margin", "Gross margin"]
CM_LABELS   = ["Contribution margin", "Contribution Margin"]
EBITDA_LABELS = ["EBITDA"]

# Column layout: row 8 headers; FY24/25 = cols 5–16 (E–P), FY25/26 = cols 18–29 (R–AC)
# Col indices are 1-based (openpyxl default). Col 17 and 30 are annual subtotals — skip.
FY2425_COLS = list(range(5, 17))   # cols E..P  (12 months Jul-24..Jun-25)
FY2526_COLS = list(range(18, 30))  # cols R..AC (12 months Jul-25..Jun-26)

LABELS_2425 = ["Jul-24","Aug-24","Sep-24","Oct-24","Nov-24","Dec-24",
               "Jan-25","Feb-25","Mar-25","Apr-25","May-25","Jun-25"]
LABELS_2526 = ["Jul-25","Aug-25","Sep-25","Oct-25","Nov-25","Dec-25",
               "Jan-26","Feb-26","Mar-26","Apr-26","May-26","Jun-26"]

# ── Counterparty → BU mapping (first-pass; confirm with Henry) ───────────────
# Keys are substrings found in the CF tracker row labels (case-insensitive).
# Edit mapping.csv (loaded below) to override after Henry's review.
DEFAULT_MAPPING = {
    "tracks":   "Tracks",
    "olaya":    "Tracks",
    "tanmya":   "Government Schools",
    "atyab":    "Government Schools",
    "maarif":   "Government Schools",
    "tamkeen":  "Government Schools",
    "uffoq":    "Global Non-Saudi",
    "takaful":  "KSA B2C",
    "b2b":      "B2B",
    "mcit":     "B2B",
    "ngos":     "B2B",
}


def load_mapping(csv_path: Path) -> dict:
    mapping = dict(DEFAULT_MAPPING)
    if not csv_path.exists():
        return mapping
    with open(csv_path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or line.startswith("counterparty"):
                continue
            parts = line.split(",", 1)
            if len(parts) == 2:
                mapping[parts[0].strip().lower()] = parts[1].strip()
    return mapping


def find_row(ws, labels: list[str], label_col: int = 3, max_row: int = 80) -> int | None:
    """Return 1-based row index where col label_col matches any of labels."""
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=label_col, max_col=label_col):
        cell_val = row[0].value
        if cell_val is None:
            continue
        cell_str = str(cell_val).strip()
        for lbl in labels:
            if cell_str.lower() == lbl.lower():
                return row[0].row
    return None


def row_values(ws, row: int, cols: list[int]) -> list[float]:
    return [float(ws.cell(row=row, column=c).value or 0) for c in cols]


def extract_bu(ws, revenue_labels: list[str], cols: list[int]) -> dict:
    rev_row = find_row(ws, revenue_labels)
    gp_row  = find_row(ws, GP_LABELS)
    cm_row  = find_row(ws, CM_LABELS)
    eb_row  = find_row(ws, EBITDA_LABELS)

    missing = []
    if rev_row is None: missing.append("revenue")
    if gp_row  is None: missing.append("gross_profit")
    if cm_row  is None: missing.append("contribution_margin")
    if eb_row  is None: missing.append("ebitda")
    if missing:
        print(f"  WARNING: could not find rows for: {missing}")

    def safe_vals(r):
        return row_values(ws, r, cols) if r else [0.0] * len(cols)

    return {
        "revenue":             safe_vals(rev_row),
        "gross_profit":        safe_vals(gp_row),
        "contribution_margin": safe_vals(cm_row),
        "ebitda":              safe_vals(eb_row),
    }


def build_fy_months(bu_data: dict, labels: list[str]) -> list[dict]:
    bus = list(bu_data.keys())
    months = []
    for i, lbl in enumerate(labels):
        by_bu = {}
        for bu in bus:
            by_bu[bu] = {field: round(bu_data[bu][field][i]) for field in bu_data[bu]}
        months.append({"n": i + 1, "label": lbl, "by_bu": by_bu})
    return months


def extract_cash(cf_path: Path) -> list[dict]:
    """Extract monthly cash from 'Monthly CF' tab."""
    print(f"\nLoading CF tracker: {cf_path.name}")
    wb = openpyxl.load_workbook(cf_path, data_only=True)
    if "Monthly CF" not in wb.sheetnames:
        print("  WARNING: 'Monthly CF' sheet not found in CF tracker.")
        return []
    ws = wb["Monthly CF"]

    # Row 5 has month headers (calendar, e.g. Jan-2026)
    # Row 9  = Beginning cash, Row 94 = Ending cash
    header_row = 5
    beg_row    = 9
    end_row    = 94

    cash_months = []
    for col in range(2, ws.max_column + 1):
        hdr = ws.cell(row=header_row, column=col).value
        if hdr is None:
            continue
        label = _parse_month_label(hdr)
        if label is None:
            continue
        eom = ws.cell(row=end_row, column=col).value or 0
        cash_months.append({"label": label, "eom_cash": round(float(eom))})

    # Map to FY25/26 months
    fy_order = {lbl: i + 1 for i, lbl in enumerate(LABELS_2526)}
    result = []
    for m in cash_months:
        if m["label"] in fy_order:
            result.append({
                "n":        fy_order[m["label"]],
                "label":    m["label"],
                "eom_cash": m["eom_cash"],
                "invoiced": 0,
                "collected": 0,
            })
    result.sort(key=lambda x: x["n"])
    # Pad missing months with zeros
    existing = {r["n"] for r in result}
    for i, lbl in enumerate(LABELS_2526):
        if (i + 1) not in existing:
            result.append({"n": i + 1, "label": lbl, "eom_cash": 0, "invoiced": 0, "collected": 0})
    result.sort(key=lambda x: x["n"])
    return result


def _parse_month_label(val) -> str | None:
    """Convert various month formats to e.g. 'Jan-26'."""
    if isinstance(val, (datetime, date)):
        return val.strftime("%b-%y")
    s = str(val).strip()
    # try formats like "Jan-2026", "January 2026", "Jan 26"
    for fmt in ("%b-%Y", "%B %Y", "%b %y", "%b-%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%b-%y")
        except ValueError:
            pass
    return None


def main():
    parser = argparse.ArgumentParser(description="Noon Academy ETL — model + CF → fy_data.json")
    parser.add_argument("--model",   default="Noon_financial_model_2_0_v42.xlsx")
    parser.add_argument("--cf",      default="Weekly_Report_CF_2026_Final_.xlsx")
    parser.add_argument("--mapping", default="mapping.csv")
    parser.add_argument("--out",     default="fy_data.json")
    args = parser.parse_args()

    model_path   = Path(args.model)
    cf_path      = Path(args.cf)
    mapping_path = Path(args.mapping)
    out_path     = Path(args.out)

    if not model_path.exists():
        sys.exit(f"ERROR: Model file not found: {model_path}\n"
                 "Pass the correct path with --model")

    print(f"Loading model: {model_path.name}  (this may take ~30 s for a large file)")
    wb = openpyxl.load_workbook(model_path, data_only=True)

    bu_2425: dict[str, dict] = {}
    bu_2526: dict[str, dict] = {}

    for bu_name, (sheet, rev_labels) in BU_TABS.items():
        if sheet not in wb.sheetnames:
            print(f"  SKIP: sheet '{sheet}' not found in workbook")
            continue
        ws = wb[sheet]
        print(f"  Extracting {bu_name} from '{sheet}' ...")
        bu_2425[bu_name] = extract_bu(ws, rev_labels, FY2425_COLS)
        bu_2526[bu_name] = extract_bu(ws, rev_labels, FY2526_COLS)

    fy2425_months = build_fy_months(bu_2425, LABELS_2425)
    fy2526_months = build_fy_months(bu_2526, LABELS_2526)

    # Cash
    cash = []
    if cf_path.exists():
        cash = extract_cash(cf_path)
    else:
        print(f"\nWARNING: CF tracker not found at {cf_path} — cash data will be empty.")
        cash = [{"n": i+1, "label": lbl, "eom_cash": 0, "invoiced": 0, "collected": 0}
                for i, lbl in enumerate(LABELS_2526)]

    output = {
        "meta": {
            "generated_at": date.today().isoformat(),
            "source": f"Extracted from {model_path.name}",
            "billing_note": "Invoiced/collected: update from billing schedules via mapping.csv",
        },
        "buses":  list(BU_TABS.keys()),
        "fy2425": fy2425_months,
        "fy2526": fy2526_months,
        "cash":   cash,
    }

    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)
    print(f"\nWrote {out_path}  ({out_path.stat().st_size:,} bytes)")
    print("Open the dashboard:  streamlit run app.py")


if __name__ == "__main__":
    main()
