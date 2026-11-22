from io import BytesIO
from pathlib import Path
import re

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
SEQUENCE_FILE = (
    BASE_DIR / "branch_sequence.txt"
    if (BASE_DIR / "branch_sequence.txt").exists()
    else BASE_DIR / "branches.txt"
)

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

BRANCH_ALIASES = {
    "KASARAN 2": "KASARANI 2",
    "KASARANI 2": "KASARANI 2",
    "ALL SOPS": "ALLSOPS",
    "ALLSOPS": "ALLSOPS",
    "BEE CENTRE": "BEE CENTER",
    "BEE CENTER": "BEE CENTER",
    "ELD PIONEER": "ELDORET PIONEER",
    "ELDORET PIONEER": "ELDORET PIONEER",
    "ELD ROMA": "ELDORET ROMA",
    "ELDORET ROMA": "ELDORET ROMA",
}


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
    """Normalizes branch names and common spelling variants for matching."""
    s = re.sub(r"\s+", " ", str(name).strip().upper())
    if s.startswith("POS "):
        s = s[4:].strip()
    return BRANCH_ALIASES.get(s, s)


def rectify_payments_to_invoice(
    entries: list[tuple[str, float]], invoice_total: float | None
) -> tuple[list[tuple[str, float]], bool, bool, float]:
    """
    1. If parsed payment sum exceeds invoice_total and exact duplicate (category, amount)
       entries exist, removes duplicates first so the remaining sum does not fall below invoice_total.
    2. If parsed payment sum still exceeds invoice_total (e.g., payment entered higher than
       invoice total as in image_9ab4e2.png), scales payments proportionally to match invoice_total
       and flags the invoice/branch as overpaid.
    Returns (final_entries, was_rectified_dup, is_overpaid_flagged, excess_amount).
    """
    if invoice_total is None or pd.isna(invoice_total):
        return entries, False, False, 0.0

    inv_total = float(invoice_total)
    raw_sum = sum(amt for _, amt in entries)
    was_rectified_dup = False

    working = list(entries)
    if raw_sum > inv_total + 0.05:
        seen: set[tuple[str, float]] = set()
        deduped: list[tuple[str, float]] = []
        for idx, (cat, amt) in enumerate(working):
            key = (cat, round(amt, 2))
            remaining_sum = sum(a for _, a in working[idx + 1 :])
            if key in seen and (sum(a for _, a in deduped) + remaining_sum) >= inv_total - 0.05:
                was_rectified_dup = True
                continue
            seen.add(key)
            deduped.append((cat, amt))
        working = deduped

    post_dup_sum = sum(amt for _, amt in working)
    if post_dup_sum > inv_total + 0.05 and abs(post_dup_sum) > 0:
        excess_amount = round(post_dup_sum - inv_total, 2)
        ratio = inv_total / post_dup_sum
        scaled = [(cat, amt * ratio) for cat, amt in working]
        # Adjust rounding residual on last item so total equals inv_total exactly
        diff_rem = round(inv_total - sum(a for _, a in scaled), 2)
        if scaled:
            last_cat, last_amt = scaled[-1]
            scaled[-1] = (last_cat, last_amt + diff_rem)
        return scaled, was_rectified_dup, True, excess_amount

    return working, was_rectified_dup, False, 0.0


def load_sequence_df() -> pd.DataFrame:
    if SEQUENCE_FILE.exists():
        seq_df = pd.read_csv(SEQUENCE_FILE)
        seq_df["Sequence"] = pd.to_numeric(seq_df["Sequence"], errors="coerce")
        seq_df["Branch"] = seq_df["Branch"].astype(str).str.strip()
        seq_df = seq_df.dropna(subset=["Branch"]).sort_values(
            by=["Sequence"], kind="mergesort", na_position="last"
        )
        seq_df["Sequence"] = range(1, len(seq_df) + 1)
        return seq_df[["Sequence", "Branch"]].reset_index(drop=True)
    return pd.DataFrame(columns=["Sequence", "Branch"])


def merge_new_branches(seq_df: pd.DataFrame, observed_branches: list[str]) -> pd.DataFrame:
    ordered = seq_df["Branch"].tolist()
    known_keys = {normalize_branch_key(b) for b in ordered}

    for b in observed_branches:
        k = normalize_branch_key(b)
        if k and k not in known_keys:
            ordered.append(b)
            known_keys.add(k)

    return pd.DataFrame({
        "Sequence": range(1, len(ordered) + 1),
        "Branch": ordered,
    })


def process_sales(
    uploaded_file,
    seq_df: pd.DataFrame,
    fix_duplicates: bool = True,
    cap_to_invoice: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    xls = pd.ExcelFile(uploaded_file)
    df = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0])

    observed_branches: list[str] = []
    for loc in df.get("Location", pd.Series(dtype=str)).dropna():
        b_name = re.sub(r"\s+", " ", str(loc).strip())
        if b_name and b_name not in observed_branches:
            observed_branches.append(b_name)

    updated_seq_df = merge_new_branches(seq_df, observed_branches)
    branch_order = updated_seq_df["Branch"].tolist()

    key_to_display: dict[str, str] = {}
    for b in branch_order:
        k = normalize_branch_key(b)
        if k not in key_to_display:
            key_to_display[k] = b

    sales_df = df.drop_duplicates(subset=["Sale Id"]).copy() if "Sale Id" in df.columns else df.copy()

    long_rows = []
    audit_rows = []
    flagged_rows = []

    for _, row in sales_df.iterrows():
        raw_branch = re.sub(r"\s+", " ", str(row.get("Location", "") or "").strip())
        b_key = normalize_branch_key(raw_branch)
        branch_col = key_to_display.get(b_key, raw_branch)

        pay_text = str(row.get("Payment Type", "") or "")
        reported_total = row.get("Total.1", row.get("Total", None))
        inv_val = float(reported_total) if (reported_total is not None and pd.notna(reported_total)) else None

        raw_entries: list[tuple[str, float]] = []
        for method_name, sign, amount_str in PAYMENT_PATTERN.findall(pay_text):
            val = float(amount_str.replace(",", ""))
            if sign == "-":
                val = -val
            raw_entries.append((categorize_payment(method_name), val))

        raw_payment_sum = round(sum(a for _, a in raw_entries), 2)
        exceeds_invoice = inv_val is not None and raw_payment_sum > inv_val + 0.05
        raw_excess = round(raw_payment_sum - inv_val, 2) if exceeds_invoice else 0.0

        if exceeds_invoice:
            flagged_rows.append({
                "Sale Id": row.get("Sale Id", ""),
                "Branch (Location)": branch_col,
                "Raw Location": raw_branch,
                "Date": row.get("Date", ""),
                "Invoice Total": round(inv_val, 2),
                "Payment Total": raw_payment_sum,
                "Excess Amount": raw_excess,
                "Payment Type": pay_text,
            })

        if fix_duplicates or cap_to_invoice:
            cleaned_entries, was_rectified_dup, was_capped, _ = rectify_payments_to_invoice(
                raw_entries, inv_val
            )
            if not fix_duplicates and was_rectified_dup:
                cleaned_entries = raw_entries
                was_rectified_dup = False
            if not cap_to_invoice and was_capped:
                cleaned_entries = raw_entries
                was_capped = False
        else:
            cleaned_entries, was_rectified_dup, was_capped = raw_entries, False, False

        parsed_total = 0.0
        cat_amounts: dict[str, float] = {}
        for cat, val in cleaned_entries:
            cat_amounts[cat] = cat_amounts.get(cat, 0.0) + val
            parsed_total += val
            long_rows.append({"Branch": branch_col, "Category": cat, "Amount": val})

        audit_row = {
            "Sale Id": row.get("Sale Id", ""),
            "Date": row.get("Date", ""),
            "Raw Location": raw_branch,
            "Branch (Location)": branch_col,
            "Payment Type": pay_text,
            "Invoice Total": round(inv_val, 2) if inv_val is not None else 0.0,
            "Reported Payment Sum": raw_payment_sum,
            "Excess Amount": raw_excess,
            "Reported Sales Total": round(parsed_total, 2),
            "Flagged Overpayment": exceeds_invoice,
            "Rectified Dup": was_rectified_dup,
            "Capped To Invoice": was_capped,
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

    report_df.loc["TOTAL SALES"] = report_df.iloc[:-1].sum(axis=0)
    report_df["TOTAL"] = report_df.sum(axis=1)

    return report_df, pd.DataFrame(audit_rows), updated_seq_df, pd.DataFrame(flagged_rows)


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
            cell.number_format = "#,##0.00;(#,##0.00);0.00"
            cell.font = total_font if (is_total or branch == "TOTAL") else regular_font
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = total_border if is_total else grid_border
            if is_total or branch == "TOTAL":
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


def to_excel_bytes(
    report_df: pd.DataFrame,
    audit_df: pd.DataFrame,
    flagged_df: pd.DataFrame,
    source_name: str,
) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        report_df.to_excel(writer, sheet_name="Branch_Sales_Report")
        style_report_sheet(
            writer.sheets["Branch_Sales_Report"],
            report_df,
            f"Sales Report by Branch ({source_name})",
        )

        if not flagged_df.empty:
            flagged_df.to_excel(writer, sheet_name="Flagged_Overpayments", index=False)

        branch_cols = [c for c in report_df.columns if c != "TOTAL"]
        for branch in branch_cols:
            sheet_title = re.sub(r"[\\/*?:\[\]]", "", str(branch))[:31] or "Branch"
            single_df = report_df[[branch]]
            single_df.to_excel(writer, sheet_name=sheet_title)
            style_report_sheet(
                writer.sheets[sheet_title],
                single_df,
                f"{branch} — Sales Report",
            )

        audit_df.to_excel(writer, sheet_name="Sale_Audit", index=False)
    buf.seek(0)
    return buf.getvalue()


def main() -> None:
    st.set_page_config(page_title="📊 Customised Report", layout="wide")
    st.title("📊 Customised Report")

    base_seq_df = load_sequence_df()

    with st.expander("Branch Sequence Settings (Adjust order for same-day branches)", expanded=False):
        st.caption(
            "Adjust `Sequence` numbers to control left-to-right branch order. "
            "New branches in an uploaded report are automatically added to the right."
        )
        edited_seq_df = st.data_editor(
            base_seq_df,
            num_rows="dynamic",
            use_container_width=True,
            key="branch_seq_editor",
        )
        edited_seq_df["Sequence"] = pd.to_numeric(edited_seq_df["Sequence"], errors="coerce")
        edited_seq_df["Branch"] = edited_seq_df["Branch"].astype(str).str.strip()
        active_seq_df = (
            edited_seq_df.dropna(subset=["Branch"])
            .loc[lambda d: d["Branch"] != ""]
            .sort_values(by=["Sequence"], kind="mergesort", na_position="last")
            .reset_index(drop=True)
        )
        st.download_button(
            "Download Updated branch_sequence.txt",
            data=active_seq_df.to_csv(index=False).encode("utf-8"),
            file_name="branch_sequence.txt",
            mime="text/csv",
        )

    fix_duplicates = st.checkbox(
        "Auto-rectify duplicate payment entries when payment sum exceeds invoice total (e.g., double-clicked MPesa)",
        value=True,
    )
    cap_to_invoice = st.checkbox(
        "Report invoice total amount when payment exceeds invoice total and flag Sale Id + Branch",
        value=True,
    )

    uploaded = st.file_uploader("Upload Detailed Sales Report (.xlsx)", type=["xlsx", "xls"])
    if uploaded is None:
        st.info("Upload your Detailed Sales Report Excel file to generate the customised report.")
        return

    report_df, audit_df, final_seq_df, flagged_df = process_sales(
        uploaded,
        active_seq_df,
        fix_duplicates=fix_duplicates,
        cap_to_invoice=cap_to_invoice,
    )

    if not flagged_df.empty:
        st.warning(
            f"Flagged {len(flagged_df)} invoice(s) where payment amount exceeded invoice total."
        )
        with st.expander("Flagged Invoices (Payment Exceeds Invoice Total)", expanded=True):
            st.dataframe(flagged_df, use_container_width=True, hide_index=True)

    st.subheader("Customised Report")
    st.dataframe(
        report_df.style.format("{:,.2f}"),
        use_container_width=True,
    )

    excel_data = to_excel_bytes(report_df, audit_df, flagged_df, uploaded.name)
    st.download_button(
        label="Download Customised Report (.xlsx)",
        data=excel_data,
        file_name="Custom_Sales_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Branch Sequence Used"):
        st.dataframe(final_seq_df, use_container_width=True, hide_index=True)

    with st.expander("Sale-Level Audit Trail"):
        st.dataframe(audit_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
