"""
AR Aging Automation - core analysis script.

Loads an invoice extract (and optional payment receipts), validates data integrity
and cross-field controls, calculates net open AR aging buckets, flags high-risk
customers, and writes a formatted multi-sheet Excel report, including a separate
Paid Status Exceptions review tab. Source statuses are not overwritten.

Run from the project root:
    python src/ar_aging_automation.py
    python src/ar_aging_automation.py --payments data/payment_receipts.csv
"""

import argparse
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ------- Config -------
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = PROJECT_ROOT / "data" / "invoices_sample.csv"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "aging_summary.xlsx"

AGING_BUCKETS = ["Current", "1-30", "31-60", "61-90", "90+"]

RISK_MIN_OVERDUE_COUNT = 3
RISK_MIN_OVERDUE_AMOUNT = 10_000

CURRENCY_FORMAT = "$#,##0.00"
INT_FORMAT = "#,##0"
DATE_FORMAT = "YYYY-MM-DD"

# Excel styling
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
SUBTOTAL_FONT = Font(bold=True)
SUBTOTAL_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
RISK_FILL = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")


def load_and_validate(
    file_path: str | Path,
    payments_path: Optional[str | Path] = None,
    as_of_date: Optional[str] = None,
) -> Tuple[pd.DataFrame, list[dict]]:
    """Load invoices (and optional payments), run control checks, and return (df, issues)."""
    df = pd.read_csv(file_path)
    issues: list[dict] = []

    # 1. Duplicate invoice_id check
    dup_ids = df[df["invoice_id"].duplicated(keep=False) & df["invoice_id"].notna()]
    for _, row in dup_ids.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "invoice_id",
            "problem": "Duplicate invoice_id detected",
        })

    # 2. Validate dates: bad strings become NaT
    for date_col in ["invoice_date", "due_date"]:
        parsed = pd.to_datetime(df[date_col], format="%Y-%m-%d", errors="coerce")
        bad_rows = df[parsed.isna() & df[date_col].notna()]
        for _, row in bad_rows.iterrows():
            issues.append({
                "invoice_id": row["invoice_id"],
                "field": date_col,
                "problem": f"Unparseable date: {row[date_col]!r}",
            })
        df[date_col] = parsed

    # 3. Date sequence control: due_date should not precede invoice_date
    bad_order = df[
        df["invoice_date"].notna()
        & df["due_date"].notna()
        & (df["due_date"] < df["invoice_date"])
    ]
    for _, row in bad_order.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "due_date",
            "problem": f"Due date ({row['due_date'].strftime('%Y-%m-%d')}) precedes invoice date ({row['invoice_date'].strftime('%Y-%m-%d')})",
        })

    # 4. Validate PO numbers
    missing_po = df[df["po_number"].isna() | (df["po_number"].astype(str).str.strip() == "")]
    for _, row in missing_po.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "po_number",
            "problem": "Missing PO number",
        })

    # 5. Coerce numeric types & validate amounts
    df["amount_usd"] = pd.to_numeric(df["amount_usd"], errors="coerce")
    bad_amounts = df[df["amount_usd"].isna() | (df["amount_usd"] <= 0)]
    for _, row in bad_amounts.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "amount_usd",
            "problem": f"Invalid or non-positive invoice amount: {row['amount_usd']!r}",
        })

    supplied_days = pd.to_numeric(df["days_overdue"], errors="coerce").fillna(0).astype(int)

    # 6. Optional calendar recalculation when --as-of-date is supplied
    if as_of_date:
        cutoff = pd.to_datetime(as_of_date, format="%Y-%m-%d")
        calc_days = (cutoff - df["due_date"]).dt.days.clip(lower=0).fillna(supplied_days).astype(int)
        valid_due = df["due_date"].notna()
        mismatched_days = df[valid_due & (calc_days != supplied_days)]
        for idx, row in mismatched_days.iterrows():
            issues.append({
                "invoice_id": row["invoice_id"],
                "field": "days_overdue",
                "problem": f"Supplied days_overdue ({supplied_days.loc[idx]}) != recalculated ({calc_days.loc[idx]}) as of {as_of_date}",
            })
        df["days_overdue"] = calc_days
    else:
        df["days_overdue"] = supplied_days

    # 7. Cross-field control: status vs. days_overdue consistency
    norm_status = df["status"].astype(str).str.strip().str.lower()
    open_but_past_due = df[(norm_status == "open") & (df["days_overdue"] > 0)]
    for _, row in open_but_past_due.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "status",
            "problem": f"Status is 'Open' despite days_overdue = {row['days_overdue']}",
        })

    paid_but_past_due = df[(norm_status == "paid") & (df["days_overdue"] > 0)]
    for _, row in paid_but_past_due.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "status",
            "problem": f"Status is 'Paid' despite days_overdue = {row['days_overdue']}",
        })

    overdue_but_current = df[(norm_status == "overdue") & (df["days_overdue"] <= 0)]
    for _, row in overdue_but_current.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "status",
            "problem": "Status is 'Overdue' despite days_overdue <= 0",
        })

    # 8. Optional payment receipt application & ghost-Paid audit control
    if payments_path:
        pay = pd.read_csv(payments_path)
        pay["amount_paid"] = pd.to_numeric(pay["amount_paid"], errors="coerce").fillna(0.0)

        # Flag orphan payments referencing unknown invoices
        orphans = pay[~pay["invoice_id"].isin(df["invoice_id"])]
        for _, row in orphans.iterrows():
            issues.append({
                "invoice_id": row["invoice_id"],
                "field": "payment_receipt",
                "problem": f"Unapplied payment {row['payment_id']} (${row['amount_paid']:,.2f}) has no matching invoice",
            })

        pay_by_inv = (
            pay[pay["invoice_id"].isin(df["invoice_id"])]
            .groupby("invoice_id")
            .agg(amount_paid=("amount_paid", "sum"),
                 payment_receipt_count=("amount_paid", "size"))
            .reset_index()
        )
        df = df.merge(pay_by_inv, on="invoice_id", how="left")
        df["payment_receipt_count"] = df["payment_receipt_count"].fillna(0).astype(int)
        df["amount_paid"] = df["amount_paid"].fillna(0.0).round(2)
        df["net_open_usd"] = (df["amount_usd"] - df["amount_paid"]).round(2)

        # Preserve the audit flag and include insufficient payment support.
        for _, row in paid_status_exceptions(df).iterrows():
            issues.append({
                "invoice_id": row["invoice_id"],
                "field": "status",
                "problem": (
                    f"Status is 'Paid': {row['exception_reason']}. "
                    f"Unreconciled difference: ${row['unreconciled_difference']:,.2f}; "
                    "requires review, not automatically collectible AR."
                ),
            })
    else:
        df["payment_receipt_count"] = 0
        df["amount_paid"] = 0.0
        df["net_open_usd"] = df["amount_usd"].round(2)

    return df, issues



def paid_status_exceptions(df: pd.DataFrame) -> pd.DataFrame:
    """Review Paid statuses against supplied receipts, not proof of collectibility.

    Call only when a receipts file was supplied. Original amounts are assumed
    to be billed amounts, not balances already reduced by payments.
    """
    is_paid = df["status"].astype(str).str.strip().str.lower().eq("paid")
    missing = df["payment_receipt_count"].eq(0)
    residual = df["net_open_usd"].round(2).gt(0)
    result = df.loc[is_paid & (missing | residual)].copy()
    result["unreconciled_difference"] = result["net_open_usd"]
    result["exception_reason"] = "Recorded payments are less than original invoice amount"
    result.loc[result["amount_paid"].eq(0), "exception_reason"] = "Matching receipts total $0.00"
    result.loc[result["payment_receipt_count"].eq(0), "exception_reason"] = "No matching payment receipt on file"
    return result[["invoice_id", "customer_name", "invoice_date", "due_date", "status",
                   "amount_usd", "amount_paid", "unreconciled_difference", "exception_reason"]]


def write_paid_status_exceptions(wb, exceptions, has_payments: bool) -> None:
    """Write a focused review tab, separate from collections and risk lists."""
    ws = wb.create_sheet("Paid Status Exceptions")
    ws.append(["Paid status—payment not supported by receipts"])
    ws.merge_cells("A1:I1")
    _style_header_row(ws, 1)
    note = (
        "Review only: source Paid status is preserved. Differences require investigation, "
        "are not automatically collectible, and are excluded from collections/risk lists."
        if has_payments else
        "Payment support not evaluated—no receipts file supplied."
    )
    ws.append([note])
    ws.merge_cells("A2:I2")
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 32
    if has_payments and not exceptions.empty:
        ws.append(["Invoice ID", "Customer name", "Invoice date", "Due date", "Source status",
                   "Original invoice amount", "Recorded payment amount", "Unreconciled difference",
                   "Exception reason"])
        _style_header_row(ws, 9, row_idx=3)
        for cell in ws[3]:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[3].height = 32
        for values in exceptions.itertuples(index=False, name=None):
            ws.append([None if pd.isna(v) else v for v in values])
        for row in ws.iter_rows(min_row=4):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            for col in (3, 4):
                row[col - 1].number_format = DATE_FORMAT
            for col in (6, 7, 8):
                row[col - 1].number_format = CURRENCY_FORMAT
            ws.row_dimensions[row[0].row].height = 36
        ws.auto_filter.ref = f"A3:I{ws.max_row}"
    elif has_payments:
        ws.append(["No unsupported Paid-status invoices identified."])
        ws.merge_cells("A3:I3")
    ws.freeze_panes = "A4"
    for col, width in {"A":16,"B":32,"C":14,"D":14,"E":14,"F":23,"G":24,"H":24,"I":52}.items():
        ws.column_dimensions[col].width = width


def calculate_aging_buckets(df: pd.DataFrame) -> pd.DataFrame:
    """Add an 'aging_bucket' column based on days_overdue."""
    def bucket(days: int) -> str:
        if days <= 0:
            return "Current"
        if days <= 30:
            return "1-30"
        if days <= 60:
            return "31-60"
        if days <= 90:
            return "61-90"
        return "90+"

    df = df.copy()
    df["aging_bucket"] = df["days_overdue"].apply(bucket)
    return df


def flag_risk_accounts(
    open_df: pd.DataFrame,
    min_overdue_count: int = RISK_MIN_OVERDUE_COUNT,
    min_overdue_amount: float = RISK_MIN_OVERDUE_AMOUNT,
) -> pd.DataFrame:
    """Return one row per customer flagged as high-risk based on net open overdue balance."""
    overdue = open_df[open_df["days_overdue"] > 0]
    if overdue.empty:
        return pd.DataFrame(columns=["customer_name", "overdue_count", "overdue_amount", "max_days_overdue"])

    by_customer = overdue.groupby("customer_name").agg(
        overdue_count=("invoice_id", "count"),
        overdue_amount=("net_open_usd", "sum"),
        max_days_overdue=("days_overdue", "max"),
    ).reset_index()

    flagged = by_customer[
        (by_customer["overdue_count"] >= min_overdue_count)
        | (by_customer["overdue_amount"] >= min_overdue_amount)
    ].copy()
    flagged["overdue_amount"] = flagged["overdue_amount"].round(2)
    return flagged.sort_values("overdue_amount", ascending=False)


def _style_header_row(ws, num_columns: int, row_idx: int = 1) -> None:
    for col_idx in range(1, num_columns + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        if cell.value is not None:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")
    if row_idx == 1:
        ws.freeze_panes = "A2"


def _autosize_columns(ws) -> None:
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length = 10
        for cell in column_cells:
            val = cell.value
            if val is None:
                continue
            if isinstance(val, (int, float)) and cell.number_format == CURRENCY_FORMAT:
                display_len = len(f"${val:,.2f}")
            elif isinstance(val, (int, float)) and cell.number_format == INT_FORMAT:
                display_len = len(f"{int(val):,}")
            elif hasattr(val, "strftime"):
                display_len = 10
            else:
                display_len = len(str(val))
            max_length = max(max_length, display_len)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 5, 45)


def _format_dataframe_columns(ws, df: pd.DataFrame) -> None:
    currency_cols = {"amount_usd", "overdue_amount", "total_amount", "amount_paid", "net_open_usd"}
    int_cols = {"days_overdue", "overdue_count", "max_days_overdue", "invoice_count"}
    date_cols = {"invoice_date", "due_date", "payment_date"}

    for col_idx, col_name in enumerate(df.columns, start=1):
        fmt = None
        if col_name in currency_cols:
            fmt = CURRENCY_FORMAT
        elif col_name in int_cols:
            fmt = INT_FORMAT
        elif col_name in date_cols:
            fmt = DATE_FORMAT

        if fmt:
            for row_idx in range(2, len(df) + 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    cell.number_format = fmt


def _write_dataframe(ws, df: pd.DataFrame) -> None:
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        clean_row = [None if pd.isna(v) else v for v in row]
        ws.append(clean_row)
    _style_header_row(ws, len(df.columns))
    _format_dataframe_columns(ws, df)
    _autosize_columns(ws)


def generate_summary_report(
    df: pd.DataFrame,
    issues: list[dict],
    output_path: str | Path,
    has_payments: bool = False,
) -> None:
    """Write a multi-sheet Excel workbook to output_path."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    is_paid = df["status"].astype(str).str.strip().str.lower() == "paid"
    paid_df = df[is_paid].copy()
    open_df = df[~is_paid].copy()

    gross_billed = float(df["amount_usd"].sum())
    settled_paid_invoices = float(paid_df["amount_usd"].sum())
    partial_payments_applied = float(open_df["amount_paid"].sum())
    net_open_ar = float(open_df["net_open_usd"].sum())

    overdue_df = open_df[open_df["days_overdue"] > 0]
    total_overdue = float(overdue_df["net_open_usd"].sum())
    pct_overdue = (total_overdue / net_open_ar * 100) if net_open_ar else 0.0

    # Defensive accounting control assertion
    bucket_sum = round(float(open_df["net_open_usd"].sum()), 2)
    assert round(net_open_ar, 2) == bucket_sum, f"Control tie-out mismatch: {net_open_ar} != {bucket_sum}"

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"

    ws_summary.append(["Metric", "Value", "Invoice Count"])
    ws_summary.append(["Gross Billed Invoices", round(gross_billed, 2), len(df)])
    ws_summary.append(["Less: Invoices marked Paid", round(-settled_paid_invoices, 2), len(paid_df)])

    if has_payments:
        partial_count = int((open_df["amount_paid"] > 0).sum())
        ws_summary.append(["Less: Partial Payments Applied", round(-partial_payments_applied, 2), partial_count])
        net_row_idx = 5
    else:
        net_row_idx = 4

    ws_summary.append(["Net Open AR", round(net_open_ar, 2), len(open_df)])
    ws_summary.append(["Total Overdue (>0 Days)", round(total_overdue, 2), len(overdue_df)])
    ws_summary.append(["% of Open AR Overdue", f"{pct_overdue:.1f}%", None])
    ws_summary.append([])  # spacer

    for r in range(2, net_row_idx + 2):
        ws_summary.cell(row=r, column=2).number_format = CURRENCY_FORMAT
        ws_summary.cell(row=r, column=3).number_format = INT_FORMAT
    ws_summary.cell(row=net_row_idx + 2, column=2).alignment = Alignment(horizontal="right")

    for col_idx in (1, 2, 3):
        cell = ws_summary.cell(row=net_row_idx, column=col_idx)
        cell.font = SUBTOTAL_FONT
        cell.fill = SUBTOTAL_FILL

    bucket_header_row = net_row_idx + 4
    ws_summary.append(["Aging Bucket (Net Open AR)", "Invoice Count", "Net Open Amount"])
    bucket_summary = (
        open_df.groupby("aging_bucket")
        .agg(invoice_count=("invoice_id", "count"), total_amount=("net_open_usd", "sum"))
        .reindex(AGING_BUCKETS, fill_value=0)
        .reset_index()
    )
    for idx, row in bucket_summary.iterrows():
        excel_row = bucket_header_row + 1 + idx
        ws_summary.append([row["aging_bucket"], int(row["invoice_count"]), round(float(row["total_amount"]), 2)])
        ws_summary.cell(row=excel_row, column=2).number_format = INT_FORMAT
        ws_summary.cell(row=excel_row, column=3).number_format = CURRENCY_FORMAT

    total_row_idx = bucket_header_row + 1 + len(AGING_BUCKETS)
    ws_summary.append(["Total Net Open AR", int(bucket_summary["invoice_count"].sum()), round(float(bucket_summary["total_amount"].sum()), 2)])
    ws_summary.cell(row=total_row_idx, column=2).number_format = INT_FORMAT
    ws_summary.cell(row=total_row_idx, column=3).number_format = CURRENCY_FORMAT
    for col_idx in (1, 2, 3):
        cell = ws_summary.cell(row=total_row_idx, column=col_idx)
        cell.font = SUBTOTAL_FONT
        cell.fill = SUBTOTAL_FILL

    # Exceptions are visible but are not added to collectible AR.
    exceptions = paid_status_exceptions(df) if has_payments else None
    ws_summary.append([])
    ws_summary.append(["Paid-status payment support", "Value"])
    support_header = ws_summary.max_row
    _style_header_row(ws_summary, 2, row_idx=support_header)
    if has_payments:
        ws_summary.append(["Paid-status exception count", len(exceptions)])
        ws_summary.cell(ws_summary.max_row, 2).number_format = INT_FORMAT
        ws_summary.append(["Paid-status unreconciled difference",
                           round(float(exceptions["unreconciled_difference"].sum()), 2)])
        ws_summary.cell(ws_summary.max_row, 2).number_format = CURRENCY_FORMAT
        note = ("Paid status is not fully supported by supplied receipts. "
                "Differences require review; not automatically collectible AR.")
    else:
        note = "Payment support not evaluated—no receipts file supplied."
    ws_summary.append([note])
    note_row = ws_summary.max_row
    ws_summary.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=3)
    ws_summary.cell(note_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws_summary.row_dimensions[note_row].height = 45

    _style_header_row(ws_summary, 3, row_idx=1)
    _style_header_row(ws_summary, 3, row_idx=bucket_header_row)
    _autosize_columns(ws_summary)

    # ----- Sheet 2: Detail >30d -----
    detail_df = open_df[open_df["days_overdue"] > 30].sort_values("days_overdue", ascending=False)
    if not has_payments and "amount_paid" in detail_df.columns:
        detail_df = detail_df.drop(columns=["amount_paid", "net_open_usd"])

    ws_detail = wb.create_sheet("Detail >30d")
    if not detail_df.empty:
        _write_dataframe(ws_detail, detail_df)
        for row_idx, days in enumerate(detail_df["days_overdue"].tolist(), start=2):
            if days > 60:
                for col_idx in range(1, len(detail_df.columns) + 1):
                    ws_detail.cell(row=row_idx, column=col_idx).fill = RISK_FILL
    else:
        ws_detail.append(["No open invoices >30 days overdue"])

    # ----- Sheet 3: Risk Accounts -----
    risk_df = flag_risk_accounts(open_df)
    ws_risk = wb.create_sheet("Risk Accounts")
    if not risk_df.empty:
        _write_dataframe(ws_risk, risk_df)
    else:
        ws_risk.append(["No risk accounts flagged"])

    # ----- Sheet 4: Validation Issues -----
    ws_issues = wb.create_sheet("Validation Issues")
    if issues:
        issues_df = pd.DataFrame(issues)
        _write_dataframe(ws_issues, issues_df)
    else:
        ws_issues.append(["No validation issues found"])

    write_paid_status_exceptions(wb, exceptions, has_payments)
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate an AR aging report from an invoices CSV."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help=f"Path to invoices CSV (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--payments",
        default=None,
        help="Optional path to payment receipts CSV to net partial payments",
    )
    parser.add_argument(
        "--as-of-date",
        default=None,
        help="Optional cutoff date (YYYY-MM-DD) to recalculate days_overdue from due_date",
    )
    parser.add_argument(
        "-o", "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to output xlsx (default: {DEFAULT_OUTPUT})",
    )
    args = parser.parse_args()

    print(f"Loading data from {args.input_file}...")
    df, issues = load_and_validate(
        args.input_file,
        payments_path=args.payments,
        as_of_date=args.as_of_date,
    )
    print(f"  Loaded {len(df)} invoices")
    print(f"  Found {len(issues)} validation issues")

    print("Calculating aging buckets...")
    df = calculate_aging_buckets(df)

    print(f"Writing report to {args.output}...")
    generate_summary_report(
        df,
        issues,
        args.output,
        has_payments=bool(args.payments),
    )
    print("Done. Open the workbook to review the report.")


if __name__ == "__main__":
    main()
