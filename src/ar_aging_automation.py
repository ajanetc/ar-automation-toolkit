"""
AR Aging Automation - core analysis script. 

Loads an invoice extract, validates it, classifies invoices into aging flags high-risk customers, 
and writes a formatted multi-sheet Excel report.

Run from the project root:
    python src/ar_aging_automation.py

"""

import argparse ## added for hardcoding input/output paths, can be extended for CLI flexibility
from pathlib import Path
from typing import Tuple

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

# Excel styling
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
RISK_FILL = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")


def load_and_validate(file_path: str | Path) -> Tuple[pd.DataFrame, list[dict]]:
    """Load CSV, coerce dtypes, return (clean_df, issues).

    issues is a list of {invoice_id, field, problem} dicts for downstream review.
    Rows with critical errors (bad dates) are still returned in the df with NaT,
    so the validation sheet can show every problem row.
    """
    df = pd.read_csv(file_path)
    issues: list[dict] = []

    # Validate dates: bad strings become NaT (pandas "missing date")
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

    # Validate PO numbers
    missing_po = df[df["po_number"].isna() | (df["po_number"].astype(str).str.strip() == "")]
    for _, row in missing_po.iterrows():
        issues.append({
            "invoice_id": row["invoice_id"],
            "field": "po_number",
            "problem": "Missing PO number",
        })

    # Coerce types
    df["amount_usd"] = pd.to_numeric(df["amount_usd"], errors="coerce")
    df["days_overdue"] = pd.to_numeric(df["days_overdue"], errors="coerce").fillna(0).astype(int)

    return df, issues


def calculate_aging_buckets(df: pd.DataFrame) -> pd.DataFrame:
    """Add an 'aging_bucket' column based on days_overdue.

    Buckets: Current (<=0), 1-30, 31-60, 61-90, 90+.
    """
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
    df: pd.DataFrame,
    min_overdue_count: int = RISK_MIN_OVERDUE_COUNT,
    min_overdue_amount: float = RISK_MIN_OVERDUE_AMOUNT,
) -> pd.DataFrame:
    """Return one row per customer flagged as high-risk.

    A customer is flagged when they meet either threshold:
      - 3+ overdue invoices, OR
      - More than $10,000 in overdue balance.
    """
    overdue = df[df["days_overdue"] > 0]
    by_customer = overdue.groupby("customer_name").agg(
        overdue_count=("invoice_id", "count"),
        overdue_amount=("amount_usd", "sum"),
        max_days_overdue=("days_overdue", "max"),
    ).reset_index()

    flagged = by_customer[
        (by_customer["overdue_count"] >= min_overdue_count)
        | (by_customer["overdue_amount"] >= min_overdue_amount)
    ].copy()
    flagged["overdue_amount"] = flagged["overdue_amount"].round(2)
    return flagged.sort_values("overdue_amount", ascending=False)


def _style_header_row(ws, num_columns: int) -> None:
    """Apply header styling and freeze the top row."""
    for col_idx in range(1, num_columns + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def _autosize_columns(ws) -> None:
    """Roughly auto-size each column to fit its widest value."""
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length = max((len(str(cell.value)) for cell in column_cells if cell.value is not None), default=10)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 40)


def _write_dataframe(ws, df: pd.DataFrame) -> None:
    """Write a DataFrame to a worksheet with header styling."""
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    _style_header_row(ws, len(df.columns))
    _autosize_columns(ws)


def generate_summary_report(df: pd.DataFrame, issues: list[dict], output_path: str | Path) -> None:
    """Write a multi-sheet Excel workbook to output_path.

    Sheets:
      1. Summary           — totals + breakdown by aging bucket
      2. Detail >30d       — invoices >30 days overdue, for collections
      3. Risk Accounts     — customers flagged by flag_risk_accounts
      4. Validation Issues — anything caught by load_and_validate
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # ----- Sheet 1: Summary -----
    ws_summary = wb.active
    ws_summary.title = "Summary"

    total_ar = df["amount_usd"].sum()
    overdue_df = df[df["days_overdue"] > 0]
    total_overdue = overdue_df["amount_usd"].sum()
    pct_overdue = (total_overdue / total_ar * 100) if total_ar else 0

    ws_summary.append(["Metric", "Value"])
    ws_summary.append(["Total AR", round(total_ar, 2)])
    ws_summary.append(["Total Overdue", round(total_overdue, 2)])
    ws_summary.append(["% Overdue", f"{pct_overdue:.1f}%"])
    ws_summary.append(["Invoice Count", len(df)])
    ws_summary.append(["Overdue Invoice Count", len(overdue_df)])
    ws_summary.append([])  # spacer

    # Bucket breakdown
    ws_summary.append(["Aging Bucket", "Invoice Count", "Total Amount"])
    bucket_summary = (
        df.groupby("aging_bucket")
        .agg(invoice_count=("invoice_id", "count"), total_amount=("amount_usd", "sum"))
        .reindex(AGING_BUCKETS, fill_value=0)
        .reset_index()
    )
    for _, row in bucket_summary.iterrows():
        ws_summary.append([row["aging_bucket"], int(row["invoice_count"]), round(row["total_amount"], 2)])

    # Style summary headers
    for header_row in (1, 8):
        for col_idx in (1, 2, 3):
            cell = ws_summary.cell(row=header_row, column=col_idx)
            if cell.value is not None:
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(horizontal="center")
    _autosize_columns(ws_summary)

    # ----- Sheet 2: Detail >30d -----
    detail_df = df[df["days_overdue"] > 30].sort_values("days_overdue", ascending=False)
    ws_detail = wb.create_sheet("Detail >30d")
    if not detail_df.empty:
        _write_dataframe(ws_detail, detail_df)
        # Highlight rows >60 days overdue in light red
        for row_idx, days in enumerate(detail_df["days_overdue"].tolist(), start=2):
            if days > 60:
                for col_idx in range(1, len(detail_df.columns) + 1):
                    ws_detail.cell(row=row_idx, column=col_idx).fill = RISK_FILL
    else:
        ws_detail.append(["No invoices >30 days overdue \ud83c\udf89"])

    # ----- Sheet 3: Risk Accounts -----
    risk_df = flag_risk_accounts(df)
    ws_risk = wb.create_sheet("Risk Accounts")
    if not risk_df.empty:
        _write_dataframe(ws_risk, risk_df)
    else:
        ws_risk.append(["No risk accounts flagged \ud83c\udf89"])

    # ----- Sheet 4: Validation Issues -----
    ws_issues = wb.create_sheet("Validation Issues")
    if issues:
        issues_df = pd.DataFrame(issues)
        _write_dataframe(ws_issues, issues_df)
    else:
        ws_issues.append(["No validation issues found \ud83c\udf89"])

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
        "-o", "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to output xlsx (default: {DEFAULT_OUTPUT})",
    )
    args = parser.parse_args()

    print(f"Loading data from {args.input_file}...")
    df, issues = load_and_validate(args.input_file)
    print(f"  Loaded {len(df)} invoices")
    print(f"  Found {len(issues)} validation issues")

    print("Calculating aging buckets...")
    df = calculate_aging_buckets(df)

    print(f"Writing report to {args.output}...")
    generate_summary_report(df, issues, args.output)
    print("✅ Done. Open the workbook to review the report.")


if __name__ == "__main__":
    main()