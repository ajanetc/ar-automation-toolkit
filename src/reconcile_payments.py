"""
Payment Reconciliation — match payment receipts against open invoices.

Produces a four-sheet Excel report:
  1. Matched           — payment fully covers the invoice
  2. Partial           — payment received, balance still owed
  3. Unmatched Payments— payments with no matching invoice (orphans)
  4. Unpaid Invoices   — open/overdue invoices with no payment

Run from the project root:
    python src/reconcile_payments.py
"""

import argparse ## added for hardcoding input/output paths, can be extended for CLI flexibility
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ------- Config -------
PROJECT_ROOT = Path(__file__).resolve().parent.parent
INVOICES_INPUT = PROJECT_ROOT / "data" / "invoices_sample.csv"
PAYMENTS_INPUT = PROJECT_ROOT / "data" / "payment_receipts.csv"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "reconciliation_report.xlsx"

# Tolerance for "fully matched" — payments within $0.01 of invoice are matched
MATCH_TOLERANCE = 0.01

# Excel styling
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")


def reconcile_payments(
    invoices: pd.DataFrame, payments: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    """Match payments to invoices, return four DataFrames keyed by status.

    Returns a dict with keys: matched, partial, unmatched_payments, unpaid_invoices.
    """
    # Sum payments per invoice (handles multiple partial payments on one invoice)
    payments_by_invoice = (
        payments.groupby("invoice_id")
        .agg(total_paid=("amount_paid", "sum"), payment_count=("payment_id", "count"))
        .reset_index()
    )

    # Outer join: every invoice + every group of payments
    merged = invoices.merge(
        payments_by_invoice, on="invoice_id", how="outer", indicator=True
    )

    # Unmatched payments: invoice_id appears in payments but not in invoices
    unmatched_mask = merged["_merge"] == "right_only"
    unmatched_payments = payments[
        payments["invoice_id"].isin(merged.loc[unmatched_mask, "invoice_id"])
    ].copy()

    # Drop unmatched-payment-only rows from the merged frame
    merged = merged[merged["_merge"] != "right_only"].drop(columns="_merge")
    merged["total_paid"] = merged["total_paid"].fillna(0)
    merged["balance_due"] = (merged["amount_usd"] - merged["total_paid"]).round(2)

    # Matched: balance is within tolerance of zero
    matched_mask = merged["balance_due"].abs() <= MATCH_TOLERANCE
    matched = merged[matched_mask & (merged["total_paid"] > 0)].copy()

    # Partial: payment received but balance still owed
    partial_mask = (merged["total_paid"] > 0) & (merged["balance_due"] > MATCH_TOLERANCE)
    partial = merged[partial_mask].copy()

    # Unpaid: no payment received
    unpaid_mask = merged["total_paid"] == 0
    unpaid_invoices = merged[unpaid_mask].drop(columns=["total_paid", "payment_count", "balance_due"]).copy()

    return {
        "matched": matched.sort_values("customer_name"),
        "partial": partial.sort_values("balance_due", ascending=False),
        "unmatched_payments": unmatched_payments.sort_values("payment_date", ascending=False),
        "unpaid_invoices": unpaid_invoices.sort_values("days_overdue", ascending=False),
    }


def _style_header_row(ws, num_columns: int) -> None:
    for col_idx in range(1, num_columns + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def _autosize_columns(ws) -> None:
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=10,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 40)


def _write_sheet(wb: Workbook, name: str, df: pd.DataFrame, empty_message: str) -> None:
    """Add a worksheet for a DataFrame, with header styling. Show empty_message if df is empty."""
    ws = wb.create_sheet(name) if name != wb.active.title else wb.active
    if df.empty:
        ws.append([empty_message])
        return
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    _style_header_row(ws, len(df.columns))
    _autosize_columns(ws)


def generate_reconciliation_report(
    results: dict[str, pd.DataFrame], output_path: str | Path
) -> None:
    """Write the four reconciliation sheets to an Excel workbook."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    # Replace default sheet with the Summary
    wb.active.title = "Summary"
    ws_summary = wb.active
    ws_summary.append(["Bucket", "Row Count", "Total $"])

    for label, key, amount_col in [
        ("Matched", "matched", "amount_usd"),
        ("Partial", "partial", "balance_due"),
        ("Unmatched Payments", "unmatched_payments", "amount_paid"),
        ("Unpaid Invoices", "unpaid_invoices", "amount_usd"),
    ]:
        df = results[key]
        total = round(df[amount_col].sum(), 2) if not df.empty else 0
        ws_summary.append([label, len(df), total])

    _style_header_row(ws_summary, 3)
    _autosize_columns(ws_summary)

    # Detail sheets
    _write_sheet(wb, "Matched", results["matched"], "No matched payments")
    _write_sheet(wb, "Partial", results["partial"], "No partial payments")
    _write_sheet(wb, "Unmatched Payments", results["unmatched_payments"], "No unmatched payments \ud83c\udf89")
    _write_sheet(wb, "Unpaid Invoices", results["unpaid_invoices"], "No unpaid invoices \ud83c\udf89")

    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Reconcile payment receipts against invoices."
    )
    parser.add_argument(
        "--invoices",
        default=str(INVOICES_INPUT),
        help=f"Path to invoices CSV (default: {INVOICES_INPUT})",
    )
    parser.add_argument(
        "--payments",
        default=str(PAYMENTS_INPUT),
        help=f"Path to payments CSV (default: {PAYMENTS_INPUT})",
    )
    parser.add_argument(
        "-o", "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to output xlsx (default: {DEFAULT_OUTPUT})",
    )
    args = parser.parse_args()

    print(f"Loading invoices from {args.invoices}...")
    invoices = pd.read_csv(args.invoices)
    print(f"  Loaded {len(invoices)} invoices")

    print(f"Loading payments from {args.payments}...")
    payments = pd.read_csv(args.payments)
    print(f"  Loaded {len(payments)} payments")

    print("Reconciling...")
    results = reconcile_payments(invoices, payments)
    for key, df in results.items():
        print(f"  {key}: {len(df)} rows")

    print(f"Writing report to {args.output}...")
    generate_reconciliation_report(results, args.output)

    print("✅ Done. Open the workbook to review the reconciliation.")


if __name__ == "__main__":
    main()