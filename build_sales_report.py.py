from pathlib import Path
import re
import sys
import tkinter as tk
from tkinter import filedialog

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
SEQUENCE_FILE = BASE_DIR / "branch_sequence.csv"
DEFAULT_OUTPUT = BASE_DIR / "Custom_Sales_Report.xlsx"

# Matches e.g. "MPesa: Ksh1,300.00" or "MPesa: -Ksh1,300.00"
PAYMENT_PATTERN = re.compile(
    r"([A-Za-z0-9\s\-]+?):\s*(-?)Ksh([\d,]+(?:\.\d+)?)",
    re.IGNORECASE,
)

# Exact row labels matching the target report format
REPORT_ROWS = [
    "MPESA SALES",
    "CREDIT CARD SALES",
    "BANK TRANFER SALES",
    "CHQ RECEIPTS SALES",
    "ON ACCOUNT SALES",
    "CREDIT SALES",
    "CASH SALES",
    "TOTAL SALES",
]


def categorize_payment(raw_method: str) -> str:
    cleaned = re.sub(r"\s+", " ", raw_method.strip().lower())
    if cleaned in {"mpesa", "mpesa by push"}:
        return "MPESA SALES"
    if cleaned in {"cc-coop", "cc-absa", "cc-equity"}:
        return "CREDIT CARD SALES"
    if cleaned == "bank":
        return "BANK TRANFER SALES"
    if "chq" in cleaned or "cheque" in cleaned:
        return "CHQ RECEIPTS SALES"
    if cleaned == "store account":
        return "ON ACCOUNT SALES"
    if "credit" in cleaned and "card" not in cleaned:
        return "CREDIT SALES"
    if cleaned == "cash":
        return "CASH SALES"
    return f"OTHER ({raw_method.strip().upper()})"


def normalize_branch_key(name: str) -> str:
    """Normalizes branch names so 'POS JUJA' and 'JUJA' match cleanly."""
    s = re.sub(r"\s+", " ", str(name).strip().upper())
    if s.startswith("POS "):
        s = s[4:].strip()
    return s


def pick_input_file() -> Path:
    if len(sys.argv) > 1:
        candidate = Path(sys.argv[1]).expanduser().resolve()
        if candidate.exists():
            return candidate

    matches = sorted(
        [p for p in BASE_DIR.glob("Detailed Sales Report*.xlsx") if not p.name.startswith("~$")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if matches:
        return matches[0]

    root = tk.Tk()
    root.withdraw()
    selected = filedialog.askopenfilename(
        title="Select Detailed Sales Report Excel File",
        initialdir=str(BASE_DIR),
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    root.destroy()
    if not selected:
        raise FileNotFoundError("No detailed sales report file was selected.")
    return Path(selected)


def load_or_init_sequence(observed_branches: list[str]) -> list[str]:
    if SEQUENCE_FILE.exists():
        seq_df = pd.read_csv(SEQUENCE_FILE)
        seq_df["Branch"] = seq_df["Branch"].astype(str).str.strip()
        seq_df["Sequence"] = pd.to_numeric(seq_df["Sequence"], errors="coerce")
        seq_df = seq_df.dropna(subset=["Branch"]).sort_values(
            by=["Sequence"], kind="mergesort", na_position="last"
        )
        ordered = [b for b in seq_df["Branch"].tolist() if b]
    else:
        ordered = []

    known_keys = {normalize_branch_key(b) for b in ordered}
    for b in observed_branches:
        k = normalize_branch_key(b)
        if k and k not in known_keys:
            ordered.append(b)
            known_keys.add(k)

    out_seq = pd.DataFrame({
        "Sequence": range(1, len(ordered) + 1),
        "Branch": ordered,
    })
    out_seq.to_csv(SEQUENCE_FILE, index=False)
    return ordered


def process_sales(input_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    xls = pd.ExcelFile(input_path)
    df = pd.read_excel(input_path, sheet_name=xls.sheet_names[0])

    observed_branches: list[str] = []
    for loc in df["Location"].dropna():
        b_name = re.sub(r"\s+", " ", str(loc).strip())
        if b_name and b_name not in observed_branches:
            observed_branches.append(b_name)

    branch_order = load_or_init_sequence(observed_branches)

    # Map normalized branch keys to their display name in branch_sequence.csv
    key_to_display: dict[str, str] = {}
    for b in branch_order:
        k = normalize_branch_key(b)
        if k not in key_to_display:
            key_to_display[k] = b

    # Deduplicate to sale level so multi-item receipts aren't double-counted
    sales_df = df.drop_duplicates(subset=["Sale Id"]).copy() if "Sale Id" in df.columns else df.copy()

    long_rows = []
    audit_rows = []

    for _, row in sales_df.iterrows():
        raw_branch = re.sub(r"\s+", " ", str(row.get("Location", "") or "").strip())
        b_key = normalize_branch_key(raw_branch)
        branch_col = key_to_display.get(b_key, raw_branch)

        pay_text = str(row.get("Payment Type", "") or "")
        parsed_total = 0.0
        cat_amounts: dict[str, float] = {}

        for method_name, sign, amount_str in PAYMENT_PATTERN.findall(pay_text):
            val = float(amount_str.replace(",", ""))
            if sign == "-":
                val = -val
            cat = categorize_payment(method_name)
            cat_amounts[cat] = cat_amounts.get(cat, 0.0) + val
            parsed_total += val
            long_rows.append({"Branch": branch_col, "Category": cat, "Amount": val})

        audit_row = {
            "Sale Id": row.get("Sale Id", ""),
            "Date": row.get("Date", ""),
            "Raw Location": raw_branch,
            "Report Branch": branch_col,
            "Payment Type": pay_text,
            "Reported Total": row.get("Total.1", row.get("Total", 0.0)),
            "Parsed Total": round(parsed_total, 2),
        }
        audit_row.update({k: round(v, 2) for k, v in cat_amounts.items()})
        audit_rows.append(audit_row)

    long_df = pd.DataFrame(long_rows)
    if long_df.empty:
        raw_pivot = pd.DataFrame(index=REPORT_ROWS[:-1], columns=branch_order).fillna(0.0)
    else:
        raw_pivot = long_df.pivot_table(
            index="Category",
            columns="Branch",
            values="Amount",
            aggfunc="sum",
            fill_value=0.0,
        )

    extra_cats = sorted([c for c in raw_pivot.index if c not in REPORT_ROWS])
    row_labels = REPORT_ROWS[:-1] + extra_cats + ["TOTAL SALES"]

    report_df = pd.DataFrame(0.0, index=row_labels, columns=branch_order)
    for r_lbl in row_labels:
        if r_lbl != "TOTAL SALES" and r_lbl in raw_pivot.index:
            for b_col in branch_order:
                if b_col in raw_pivot.columns:
                    report_df.loc[r_lbl, b_col] = float(raw_pivot.loc[r_lbl, b_col])

    # TOTAL SALES row per branch
    report_df.loc["TOTAL SALES"] = report_df.iloc[:-1].sum(axis=0)

    return report_df, pd.DataFrame(audit_rows), branch_order


def style_report_sheet(ws, report_df: pd.DataFrame, title: str) -> None:
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_fill = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")

    title_font = Font(name="Calibri", size=12, bold=True, color="000000")
    header_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    regular_font = Font(name="Calibri", size=10, bold=False, color="000000")

    thin = Side(border_style="thin", color="BFBFBF")
    top_thin = Side(border_style="thin", color="000000")
    double_bottom = Side(border_style="double", color="000000")

    grid_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_border = Border(left=thin, right=thin, top=top_thin, bottom=double_bottom)

    ws["A1"] = title
    ws["A1"].font = title_font

    header_row = 3
    ws.cell(row=header_row, column=1, value="BRANCH NAME").font = header_font
    ws.cell(row=header_row, column=1).fill = header_fill
    ws.cell(row=header_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=header_row, column=1).border = grid_border

    for c_idx, branch in enumerate(report_df.columns, start=2):
        cell = ws.cell(row=header_row, column=c_idx, value=branch)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = grid_border

    ws.row_dimensions[header_row].height = 22

    for r_offset, row_name in enumerate(report_df.index, start=1):
        r_num = header_row + r_offset
        is_total = row_name == "TOTAL SALES"

        lbl = ws.cell(row=r_num, column=1, value=row_name)
        lbl.font = total_font if is_total else regular_font
        lbl.alignment = Alignment(horizontal="left", vertical="center")
        lbl.border = total_border if is_total else grid_border
        if is_total:
            lbl.fill = total_fill

        for c_offset, branch in enumerate(report_df.columns, start=1):
            c_num = 1 + c_offset
            val = round(float(report_df.loc[row_name, branch]), 2)
            cell = ws.cell(row=r_num, column=c_num, value=val)
            cell.number_format = "#,##0.00;(#,##0.00);0" if is_total else "#,##0.00;(#,##0.00);0.00"
            cell.font = total_font if is_total else regular_font
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = total_border if is_total else grid_border
            if is_total:
                cell.fill = total_fill

    ws.freeze_panes = "B4"
    for c_num in range(1, len(report_df.columns) + 2):
        col_letter = get_column_letter(c_num)
        max_len = 14
        for r_num in range(header_row, header_row + len(report_df.index) + 1):
            v = ws.cell(row=r_num, column=c_num).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 24)


def main() -> None:
    input_path = pick_input_file()
    report_df, audit_df, branch_order = process_sales(input_path)

    with pd.ExcelWriter(DEFAULT_OUTPUT, engine="openpyxl") as writer:
        # 1. Main matrix matching your image layout (all branches in sequence)
        report_df.to_excel(writer, sheet_name="Branch_Sales_Report")
        style_report_sheet(
            writer.sheets["Branch_Sales_Report"],
            report_df,
            f"Sales Report by Branch ({input_path.name})",
        )

        # 2. Separate sheet per branch for standalone viewing/printing
        for branch in branch_order:
            sheet_title = re.sub(r"[\\/*?:\[\]]", "", branch)[:31] or "Branch"
            single_df = report_df[[branch]]
            single_df.to_excel(writer, sheet_name=sheet_title)
            style_report_sheet(
                writer.sheets[sheet_title],
                single_df,
                f"{branch} — Sales Report",
            )

        # 3. Audit trail sheet
        audit_df.to_excel(writer, sheet_name="Sale_Audit", index=False)

    print(f"Generated: {DEFAULT_OUTPUT.name}")
    print(f"Branches included ({len(branch_order)}): {', '.join(branch_order[:5])}...")


if __name__ == "__main__":
    main()