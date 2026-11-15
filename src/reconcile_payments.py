"""
Payment Reconciliation — match payment receipts against open invoices.

Produces a multi-sheet Excel report with full billing and cash control tie-outs:
  1. Summary            — bucket totals + subledger & cash control reconciliation
  2. Matched            — payment fully covers the invoice (within $0.01)
  3. Partial            — payment received, balance still owed
  4. Unmatched Payments — missing or unknown invoice references, with reasons
  5. Unpaid Invoices    — invoices with $0 payment received

Input requirement: one row per invoice. Repeated nonblank invoice IDs stop
reconciliation before payment application; records are never automatically
combined or deleted. Row numbers in validation messages are 1-based data rows,
excluding the CSV header. Payments with missing/blank invoice references are
retained as unmatched, never applied. Missing invoice IDs in the invoice file
and broader validation remain separate work.

Critical amount errors (blank, unreadable, non-finite, or unsupported negative),
missing required columns, and repeated invoice IDs block the command-line run.
A separate reconciliation_diagnostics.xlsx preserves loaded source records and
lists all detected critical issues. The normal report is NOT refreshed. Exit 2
means blocked/diagnostic failure; a normal completed run exits 0. Amount zeros
and unmatched receipts are not blocked. Overpayment and wider tie-out redesign
remain separate work; existing control totals are not claimed as fully verified.

Run from the project root:
    python src/reconcile_payments.py
"""

import argparse
import math
import os
import tempfile
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
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

CURRENCY_FORMAT = "$#,##0.00"
INT_FORMAT = "#,##0"

# Excel styling
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
SUBTOTAL_FONT = Font(bold=True)
SUBTOTAL_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
ALERT_FILL = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")



class DuplicateInvoiceIDError(ValueError):
    """The invoice input is not at the required one-row-per-invoice grain."""


def validate_unique_invoice_ids(invoices: pd.DataFrame) -> None:
    """Reject repeated nonblank IDs without editing or aggregating source rows.

    IDs are compared exactly as supplied (no whitespace/case normalization).
    Missing/blank IDs are not covered by this check and need separate validation.
    Positions identify 1-based data records, excluding the CSV header; they do
    not depend on the DataFrame's index or physical CSV lines with quoted text.
    """
    ids = invoices["invoice_id"]
    nonblank = ids.notna() & ids.astype("string").str.strip().ne("").fillna(False)
    repeated = nonblank & ids.duplicated(keep=False)
    if not repeated.any():
        return

    affected = {}
    for position, (invoice_id, is_repeated) in enumerate(zip(ids, repeated), start=1):
        if is_repeated:
            affected.setdefault(invoice_id, []).append(position)
    details = "\n".join(
        f"  - {invoice_id!r}: data rows {', '.join(str(n) for n in positions)}"
        for invoice_id, positions in affected.items()
    )
    raise DuplicateInvoiceIDError(
        "Repeated invoice IDs found. This tool requires one row per invoice.\n"
        "These may be duplicate records or legitimate invoice/product lines. "
        "Check the source data and prepare invoice-level records before proceeding.\n"
        "No rows were combined or deleted, and no payments have been applied.\n"
        "Affected records (1-based data rows; CSV header excluded):\n"
        + details
    )



REQUIRED_INVOICE_COLUMNS = {"invoice_id", "customer_name", "amount_usd", "days_overdue", "status"}
REQUIRED_PAYMENT_COLUMNS = {"payment_id", "invoice_id", "amount_paid", "payment_date"}
ISSUE_COLUMNS = ["source", "data_row", "record_id", "field", "original_value", "problem"]


class InputValidationError(ValueError):
    """Critical input issues prevent financial reconciliation."""

    def __init__(self, issues: list[dict]):
        self.issues = issues
        super().__init__(f"Reconciliation blocked: {len(issues)} critical input issue(s).")


def _original_text(value) -> str:
    """Keep invalid values visible; never replace an unknown amount with zero."""
    if value is None or value is pd.NA:
        return "<missing>"
    return str(value)


def _amount_problem(value) -> str | None:
    if value is None or value is pd.NA or not str(value).strip():
        return "Missing amount; cannot be treated as zero"
    if isinstance(value, bool):
        return "Unreadable amount; boolean values are not monetary amounts"
    try:
        decimal = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return "Unreadable amount; expected a plain numeric value"
    if not decimal.is_finite():
        return "Non-finite amount (NaN or infinity)"
    # The existing calculation engine uses pandas floating-point arithmetic.
    try:
        parsed = pd.to_numeric(value, errors="coerce")
        representable = pd.notna(parsed) and math.isfinite(float(parsed))
    except (TypeError, ValueError, OverflowError):
        representable = False
    if not representable:
        return "Amount cannot be represented safely by the current numeric engine"
    if decimal < 0:
        return "Negative amount unsupported: credit/reversal treatment must be defined"
    return None


def collect_input_issues(invoices: pd.DataFrame, payments: pd.DataFrame) -> list[dict]:
    """Collect ALL critical issues before grouping or payment application.

    Positive invoice/payment amounts, genuine zeros, and positive-receipt
    overpayments are permitted. Known unmatched receipts remain review items.
    Missing invoice-ID values, payment-ID duplication, date/status controls and
    the wider reconciliation redesign remain separate work.
    """
    issues = []
    for source, frame, required, amount_column, id_column in (
        ("Invoices", invoices, REQUIRED_INVOICE_COLUMNS, "amount_usd", "invoice_id"),
        ("Payments", payments, REQUIRED_PAYMENT_COLUMNS, "amount_paid", "payment_id"),
    ):
        for column in sorted(required - set(frame.columns)):
            issues.append(dict(source=source, data_row=None, record_id="", field=column,
                               original_value="<column missing>", problem="Required column is missing"))
        if amount_column in frame:
            for position, (_, row) in enumerate(frame.iterrows(), start=1):
                value = row[amount_column]
                problem = _amount_problem(value)
                if problem:
                    issues.append(dict(source=source, data_row=position,
                                       record_id=_original_text(row.get(id_column, "")),
                                       field=amount_column, original_value=_original_text(value), problem=problem))
    if "invoice_id" in invoices:
        ids = invoices["invoice_id"]
        nonblank = ids.notna() & ids.astype("string").str.strip().ne("").fillna(False)
        repeated = nonblank & ids.duplicated(keep=False)
        for position, (invoice_id, duplicate) in enumerate(zip(ids, repeated), start=1):
            if duplicate:
                issues.append(dict(source="Invoices", data_row=position, record_id=str(invoice_id),
                                   field="invoice_id", original_value=str(invoice_id),
                                   problem=("Repeated invoice ID: requires one row per invoice. "
                                            "May be duplicate records or legitimate product lines; not automatically combined.")))
    return issues


def _read_source_csv(path: str, source: str) -> tuple[pd.DataFrame, list[dict]]:
    """Preserve source text, blanks, leading zeros and tokens such as 'NaN'."""
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, skip_blank_lines=False), []
    except pd.errors.EmptyDataError:
        return pd.DataFrame(), []
    except (OSError, UnicodeError, pd.errors.ParserError) as exc:
        return pd.DataFrame(), [dict(source=source, data_row=None, record_id="", field="input_file",
                                    original_value=str(path), problem=f"Input could not be read: {exc}")]


def diagnostic_output_path(report_path: str | Path) -> Path:
    report_path = Path(report_path)
    name = ("reconciliation_diagnostics.xlsx" if report_path.stem == "reconciliation_report"
            else f"{report_path.stem}_diagnostics.xlsx")
    return report_path.with_name(name)


def _write_diagnostic_table(wb, name: str, frame: pd.DataFrame) -> None:
    ws = wb.create_sheet(name)
    if not len(frame.columns):
        ws.append(["No source rows/columns available. See Input Issues for details."])
        ws.column_dimensions["A"].width = 75
        return
    ws.append([str(column) for column in frame.columns])
    for values in frame.itertuples(index=False, name=None):
        row = []
        for column, value in zip(frame.columns, values):
            if value is None or value is pd.NA or (not isinstance(value, str) and pd.isna(value)):
                row.append(None)
            elif name == "Input Issues" and column == "data_row":
                row.append(int(value))
            else:
                row.append(str(value))
        ws.append(row)
    if name == "Input Issues" and "data_row" in frame.columns:
        row_column = list(frame.columns).index("data_row") + 1
        for cells in ws.iter_rows(min_row=2, min_col=row_column, max_col=row_column):
            cells[0].number_format = INT_FORMAT
    # Source cells are literal text: never execute source values as Excel formulas.
    for row in ws:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    _style_header_row(ws, len(frame.columns))
    for cell in ws[1]:
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    ws.auto_filter.ref = ws.dimensions
    _autosize_columns(ws)
    if name == "Input Issues":
        ws.column_dimensions["F"].width = 78
        for row in range(2, ws.max_row + 1):
            ws.row_dimensions[row].height = 48
    else:
        for row in range(2, ws.max_row + 1):
            ws.row_dimensions[row].height = 30


def generate_diagnostic_report(invoices: pd.DataFrame, payments: pd.DataFrame,
                               issues: list[dict], output_path: str | Path,
                               report_path: str | Path, invoice_source: str = "",
                               payment_source: str = "") -> None:
    """Write investigation evidence, never partial financial reconciliation totals."""
    output_path, report_path = Path(output_path), Path(report_path)
    if output_path.resolve() == report_path.resolve():
        raise ValueError("Diagnostic path must differ from the financial report path")
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Run Status", "BLOCKED—reconciliation not completed"])
    ws.append(["Generated (UTC)", datetime.now(timezone.utc).isoformat(timespec="seconds")])
    ws.append(["Input issue count", len(issues)])
    ws.append(["Invoice data records loaded", len(invoices)])
    ws.append(["Payment data records loaded", len(payments)])
    ws.append(["Invoice source", str(invoice_source)])
    ws.append(["Payment source", str(payment_source)])
    ws.append(["Normal report path", str(report_path)])
    ws.append(["Normal report status", "NOT REFRESHED. Any existing normal report belongs to an earlier run."])
    ws.append(["Required action", "Correct the source issues and rerun. No payments were applied in this blocked run."])
    ws.append(["Financial results", "Not calculated. Unknown amounts are not zero; no financial totals are presented."])
    ws.append(["Evidence", "Source tabs preserve loaded cell text. Input Issues lists 1-based data rows, excluding the CSV header."])
    ws.append(["Input scope", "If an input could not be read, its source tab is unavailable; inspect the file-level issue."])
    ws.append(["Diagnostics scope", "This workbook describes this blocked run only; it is not proof of the latest run status after a rerun."])
    for row in ws:
        row[0].font = Font(bold=True)
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row[0].row].height = 34
    for row_number in (3, 4, 5):
        ws.cell(row_number, 2).number_format = INT_FORMAT
    _style_header_row(ws, 2)
    for cell in ws[1]:
        cell.fill = ALERT_FILL
        cell.font = Font(bold=True, color="9C0006")
    ws.row_dimensions[1].height = 38
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 100
    _write_diagnostic_table(wb, "Input Issues", pd.DataFrame(issues, columns=ISSUE_COLUMNS))
    _write_diagnostic_table(wb, "Source Invoices", invoices)
    _write_diagnostic_table(wb, "Source Payments", payments)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=".reconciliation_diagnostic_", suffix=".xlsx", dir=output_path.parent)
    os.close(fd)
    try:
        wb.save(temp_name)
        os.replace(temp_name, output_path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def reconcile_payments(
    invoices: pd.DataFrame, payments: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    """Match payments to invoices, return DataFrames keyed by reconciliation status.

    Returns a dict with keys:
      - matched
      - partial
      - unmatched_payments
      - unpaid_invoices
      - overpayments (only populated if balance_due < -MATCH_TOLERANCE)
    """
    # Check before numeric coercion, grouping, merging, or payment application.
    if "invoice_id" in invoices:
        # Keep the existing library exception and repeated-ID tests compatible.
        validate_unique_invoice_ids(invoices)
    critical_issues = collect_input_issues(invoices, payments)
    if critical_issues:
        raise InputValidationError(critical_issues)
    invoices = invoices.copy()
    payments = payments.copy()

    invoices["amount_usd"] = pd.to_numeric(invoices["amount_usd"], errors="raise")
    invoices["days_overdue"] = pd.to_numeric(invoices["days_overdue"], errors="coerce").fillna(0).astype(int)
    payments["amount_paid"] = pd.to_numeric(payments["amount_paid"], errors="raise")

    # Classify individual receipts BEFORE grouping. groupby drops null keys;
    # blank references must never be applied, even if an invoice ID is blank too.
    missing_reference = (
        payments["invoice_id"].isna()
        | payments["invoice_id"].astype("string").str.strip().eq("").fillna(False)
    )
    reference_exists = payments["invoice_id"].isin(invoices["invoice_id"].dropna())
    matched_reference = ~missing_reference & reference_exists

    # Preserve one row per unmatched receipt, including repeated receipt records.
    # Do not guess an invoice or normalize nonblank IDs during this classification.
    unmatched_payments = payments.loc[~matched_reference].copy()
    unmatched_payments["unmatched_reason"] = "Invoice reference not found"
    unmatched_payments.loc[
        missing_reference.loc[~matched_reference], "unmatched_reason"
    ] = "Missing invoice reference"

    # Only receipts with a known, nonblank invoice reference may be applied.
    payments_by_invoice = (
        payments.loc[matched_reference].groupby("invoice_id")
        .agg(total_paid=("amount_paid", "sum"), payment_count=("payment_id", "count"))
        .reset_index()
    )
    merged = invoices.merge(payments_by_invoice, on="invoice_id", how="left")
    merged["total_paid"] = merged["total_paid"].fillna(0.0).round(2)
    merged["payment_count"] = merged["payment_count"].fillna(0).astype(int)
    merged["balance_due"] = (merged["amount_usd"] - merged["total_paid"]).round(2)

    # Matched: balance is within tolerance of zero
    matched_mask = (merged["balance_due"].abs() <= MATCH_TOLERANCE) & (merged["total_paid"] > 0)
    matched = merged[matched_mask].copy()

    # Partial: payment received but balance still owed
    partial_mask = (merged["total_paid"] > 0) & (merged["balance_due"] > MATCH_TOLERANCE)
    partial = merged[partial_mask].copy()

    # Overpayments: payment exceeds invoice amount beyond tolerance
    overpaid_mask = (merged["total_paid"] > 0) & (merged["balance_due"] < -MATCH_TOLERANCE)
    overpayments = merged[overpaid_mask].copy()

    # Unpaid: no payment received
    unpaid_mask = merged["total_paid"] == 0
    unpaid_invoices = merged[unpaid_mask].drop(
        columns=["total_paid", "payment_count", "balance_due"]
    ).copy()
    unpaid_invoices["control_flag"] = unpaid_invoices["status"].apply(
        lambda s: "REVIEW: Marked Paid with $0 Receipt"
        if str(s).strip().lower() == "paid"
        else "Open Unpaid"
    )

    results = {
        "matched": matched.sort_values("customer_name"),
        "partial": partial.sort_values("balance_due", ascending=False),
        "unmatched_payments": unmatched_payments.sort_values("payment_date", ascending=False),
        "unpaid_invoices": unpaid_invoices.sort_values("days_overdue", ascending=False),
    }
    if not overpayments.empty:
        results["overpayments"] = overpayments.sort_values("balance_due")
    return results


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
            else:
                display_len = len(str(val))
            max_length = max(max_length, display_len)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 5, 45)


def _format_sheet_columns(ws, df: pd.DataFrame) -> None:
    currency_cols = {"amount_usd", "total_paid", "balance_due", "amount_paid"}
    int_cols = {"days_overdue", "payment_count"}

    for col_idx, col_name in enumerate(df.columns, start=1):
        fmt = None
        if col_name in currency_cols:
            fmt = CURRENCY_FORMAT
        elif col_name in int_cols:
            fmt = INT_FORMAT

        if fmt:
            for row_idx in range(2, len(df) + 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    cell.number_format = fmt


def _write_sheet(
    wb: Workbook,
    name: str,
    df: pd.DataFrame,
    empty_message: str,
    highlight_flag_col: str | None = None,
) -> None:
    """Add a worksheet for a DataFrame with header styling, formats, and auto-fit widths."""
    ws = wb.create_sheet(name) if name != wb.active.title else wb.active
    if df.empty:
        ws.append([empty_message])
        return
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        clean_row = [None if pd.isna(v) else v for v in row]
        ws.append(clean_row)
    _style_header_row(ws, len(df.columns))
    _format_sheet_columns(ws, df)

    if highlight_flag_col and highlight_flag_col in df.columns:
        flag_idx = list(df.columns).index(highlight_flag_col) + 1
        for row_idx in range(2, len(df) + 2):
            flag_val = str(ws.cell(row=row_idx, column=flag_idx).value or "")
            if flag_val.startswith("REVIEW"):
                for col_idx in range(1, len(df.columns) + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = ALERT_FILL

    _autosize_columns(ws)


def generate_reconciliation_report(
    results: dict[str, pd.DataFrame], output_path: str | Path
) -> None:
    """Write the reconciliation summary and detail sheets to an Excel workbook."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.active.title = "Summary"
    ws_summary = wb.active

    # ----- Section 1: Reconciliation Bucket Summary -----
    ws_summary.append(["Reconciliation Bucket", "Row Count", "Total Amount"])

    bucket_specs = [
        ("Matched (Settled in Full)", "matched", "amount_usd"),
        ("Partial (Remaining Balance)", "partial", "balance_due"),
        ("Unmatched Payments (Orphan Cash)", "unmatched_payments", "amount_paid"),
        ("Unpaid Invoices ($0 Received)", "unpaid_invoices", "amount_usd"),
    ]
    if "overpayments" in results and not results["overpayments"].empty:
        bucket_specs.append(("Overpayments (Credit Balance)", "overpayments", "balance_due"))

    for idx, (label, key, amount_col) in enumerate(bucket_specs, start=2):
        df = results[key]
        total = round(float(df[amount_col].sum()), 2) if not df.empty else 0.0
        ws_summary.append([label, len(df), total])
        ws_summary.cell(row=idx, column=2).number_format = INT_FORMAT
        ws_summary.cell(row=idx, column=3).number_format = CURRENCY_FORMAT

    _style_header_row(ws_summary, 3, row_idx=1)

    # ----- Section 2: Accounting Control Tie-Out -----
    matched_df = results["matched"]
    partial_df = results["partial"]
    unmatched_df = results["unmatched_payments"]
    unpaid_df = results["unpaid_invoices"]

    gross_billed = round(
        float(matched_df["amount_usd"].sum() + partial_df["amount_usd"].sum() + unpaid_df["amount_usd"].sum()),
        2,
    )
    total_cash_received = round(
        float(matched_df["total_paid"].sum() + partial_df["total_paid"].sum() + unmatched_df["amount_paid"].sum()),
        2,
    )
    applied_cash = round(
        float(matched_df["total_paid"].sum() + partial_df["total_paid"].sum()),
        2,
    )
    unapplied_cash = round(float(unmatched_df["amount_paid"].sum()), 2)

    paid_status_defects = unpaid_df[
        unpaid_df["status"].astype(str).str.strip().str.lower() == "paid"
    ]
    defect_amount = round(float(paid_status_defects["amount_usd"].sum()), 2)
    gross_unsettled = round(
        float(partial_df["balance_due"].sum() + unpaid_df["amount_usd"].sum()),
        2,
    )
    net_open_ar = round(gross_unsettled - defect_amount, 2)

    start_control_row = len(bucket_specs) + 3
    ws_summary.append([])  # spacer
    ws_summary.append(["Control Tie-Out Metric", "Count", "Amount"])
    _style_header_row(ws_summary, 3, row_idx=start_control_row)

    control_rows = [
        ("Total Gross Invoices Billed", len(matched_df) + len(partial_df) + len(unpaid_df), gross_billed, False),
        ("Total Cash Receipts Imported", int(matched_df["payment_count"].sum() + partial_df["payment_count"].sum() + len(unmatched_df)), total_cash_received, False),
        ("  Less: Unapplied / Orphan Cash", len(unmatched_df), -unapplied_cash, False),
        ("Total Cash Applied to Subledger", int(matched_df["payment_count"].sum() + partial_df["payment_count"].sum()), applied_cash, False),
        ("Gross Unsettled Subledger Balance", len(partial_df) + len(unpaid_df), gross_unsettled, False),
        ("  Less: 'Paid'-Status Discrepancies ($0 Cash)", len(paid_status_defects), -defect_amount, False),
        ("Net Open AR (Ties to Aging Report)", len(partial_df) + len(unpaid_df) - len(paid_status_defects), net_open_ar, True),
    ]

    for offset, (label, cnt, amt, is_subtotal) in enumerate(control_rows, start=1):
        r = start_control_row + offset
        ws_summary.append([label, cnt, amt])
        ws_summary.cell(row=r, column=2).number_format = INT_FORMAT
        ws_summary.cell(row=r, column=3).number_format = CURRENCY_FORMAT
        if is_subtotal:
            for col_idx in (1, 2, 3):
                cell = ws_summary.cell(row=r, column=col_idx)
                cell.font = SUBTOTAL_FONT
                cell.fill = SUBTOTAL_FILL

    _autosize_columns(ws_summary)

    # Detail sheets
    _write_sheet(wb, "Matched", results["matched"], "No matched payments")
    _write_sheet(wb, "Partial", results["partial"], "No partial payments")
    _write_sheet(wb, "Unmatched Payments", results["unmatched_payments"], "No unmatched payments")
    _write_sheet(
        wb,
        "Unpaid Invoices",
        results["unpaid_invoices"],
        "No unpaid invoices",
        highlight_flag_col="control_flag",
    )
    if "overpayments" in results and not results["overpayments"].empty:
        _write_sheet(wb, "Overpayments", results["overpayments"], "No overpayments")

    wb.save(output_path)


def main() -> int:
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
    # Preserve identifier text, including leading zeros.
    invoices, invoice_read_issues = _read_source_csv(args.invoices, "Invoices")
    print(f"  Loaded {len(invoices)} invoices")

    print(f"Loading payments from {args.payments}...")
    payments, payment_read_issues = _read_source_csv(args.payments, "Payments")
    print(f"  Loaded {len(payments)} payments")

    critical_issues = invoice_read_issues + payment_read_issues + collect_input_issues(invoices, payments)
    if critical_issues:
        diagnostics = diagnostic_output_path(args.output)
        print(f"BLOCKED—reconciliation not completed: {len(critical_issues)} critical input issue(s).")
        print("No payments were applied. No normal reconciliation report will be generated.")
        print(f"Normal report NOT REFRESHED: {args.output}")
        print("Any existing normal report belongs to an earlier run; do not treat it as this run's result.")
        try:
            generate_diagnostic_report(invoices, payments, critical_issues, diagnostics, args.output,
                                       args.invoices, args.payments)
        except Exception as exc:
            print(f"ERROR: Could not write diagnostic workbook: {exc}")
            print("Check folder permissions and close the diagnostic workbook if open in Excel.")
            return 2
        print(f"Diagnostic workbook: {diagnostics}")
        print("Correct the input issues listed there, then rerun.")
        return 2

    print("Reconciling...")
    results = reconcile_payments(invoices, payments)
    for key, df in results.items():
        print(f"  {key}: {len(df)} rows")

    print(f"Writing report to {args.output}...")
    generate_reconciliation_report(results, args.output)

    print("Done. Open the workbook to review the reconciliation.")
    previous_diagnostics = diagnostic_output_path(args.output)
    if previous_diagnostics.exists():
        print(f"Note: earlier diagnostic workbook remains at {previous_diagnostics}; "
              "it is not the status of this successful run.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except DuplicateInvoiceIDError as exc:
        raise SystemExit(
            f"ERROR: {exc}\n"
            "Report not generated. Any existing output workbook has not been updated."
        ) from None
