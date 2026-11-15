"""
Receipt-reference reconciliation: analytical coverage, not ERP application.

Uses one row per invoice. Receipt amounts are intended for the referenced invoice;
matching proves only that the invoice exists in the supplied extract. No allocation,
posting, refund, credit memo, or write-off is made. Every source-status invoice is
included. Positive residuals and excess receipts are shown separately, while the
signed invoice-minus-receipt difference is retained. One-cent differences remain;
amounts with sub-cent precision are blocked rather than silently rounded.

Normal output contains Summary, Reconciliation Checks, Fully Covered, Partial
Coverage, Receipts Without Invoices, No Positive Receipts, Excess Receipts, and
Paid Status Exceptions. Review subsets must not be added to partition totals.
The legacy internal names total_paid/balance_due and result keys remain for API
compatibility only; workbook labels use the agreed reference-matching definitions.

Critical invalid inputs produce separate diagnostics (exit 2), without refreshing
the normal report. A failed financial check produces a separate failed-checks
workbook (exit 3). Measurable exceptions yield REVIEW with completed checks (exit 0).
Repeated invoice IDs block; missing receipt references remain visible unmatched.
Missing invoice-ID values, receipt-ID duplication, dates, and other broader data
validation remain separate work. All supplied dates are used, not historical cutoffs.

Run from the project root:
    python src/reconcile_payments.py
"""

import argparse
import math
import os
import tempfile
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, localcontext
from collections import Counter
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

# Exact-cent policy: never classify a one-cent residual as settled.
MATCH_TOLERANCE = 0.0  # Compatibility constant; classifications use integer cents.

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


def _to_cents(value) -> int:
    decimal = Decimal(str(value).strip())
    with localcontext() as context:
        context.prec = max(50, len(decimal.as_tuple().digits) + abs(decimal.as_tuple().exponent) + 5)
        return int(decimal * 100)


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
    with localcontext() as context:
        context.prec = max(50, len(decimal.as_tuple().digits) + abs(decimal.as_tuple().exponent) + 5)
        if decimal * 100 != (decimal * 100).to_integral_value():
            return "More than two decimal places unsupported; no automatic rounding or write-off"
    if _to_cents(decimal) != _to_cents(float(parsed)):
        return "Amount loses cent precision in the current numeric engine"
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



def reconcile_payments(invoices: pd.DataFrame, payments: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Match receipt INTENT by invoice reference; no ERP application or posting.

    Monetary arithmetic is in integer cents. Legacy dictionary/column names
    total_paid and balance_due remain internal API aliases, not accounting claims.
    Every invoice participates irrespective of its source status.
    """
    if "invoice_id" in invoices:
        validate_unique_invoice_ids(invoices)
    issues = collect_input_issues(invoices, payments)
    if issues:
        raise InputValidationError(issues)
    inv = invoices.copy().reset_index(drop=True)
    pay = payments.copy().reset_index(drop=True)
    inv["invoice_data_row"] = range(1, len(inv) + 1)
    pay["receipt_data_row"] = range(1, len(pay) + 1)
    inv["amount_usd"] = [_to_cents(v) / 100 for v in inv["amount_usd"]]
    pay["amount_paid"] = [_to_cents(v) / 100 for v in pay["amount_paid"]]
    inv["days_overdue"] = pd.to_numeric(inv["days_overdue"], errors="coerce").fillna(0).astype(int)
    refs = pay["invoice_id"]
    missing = refs.isna() | refs.astype("string").str.strip().eq("").fillna(False)
    exists = refs.isin(inv["invoice_id"].dropna())
    matched_reference = ~missing & exists
    matched_receipts = pay.loc[matched_reference].copy()
    unmatched = pay.loc[~matched_reference].copy()
    unmatched["unmatched_reason"] = "Invoice reference not found"
    unmatched.loc[missing.loc[~matched_reference], "unmatched_reason"] = "Missing invoice reference"

    receipt_totals, receipt_counts = {}, {}
    # Count source records, not nonblank receipt IDs. Duplicated receipt-ID
    # validation is separate work; never silently deduplicate input rows.
    for _, row in matched_receipts.iterrows():
        key = row["invoice_id"]
        receipt_totals[key] = receipt_totals.get(key, 0) + _to_cents(row["amount_paid"])
        receipt_counts[key] = receipt_counts.get(key, 0) + 1
    received, differences, positives, excesses, covered, counts = [], [], [], [], [], []
    for _, row in inv.iterrows():
        ref = row["invoice_id"]
        amount = _to_cents(row["amount_usd"])
        receipt = receipt_totals.get(ref, 0) if pd.notna(ref) else 0
        residual = amount - receipt
        received.append(receipt / 100)
        differences.append(residual / 100)
        positives.append(max(residual, 0) / 100)
        excesses.append(max(-residual, 0) / 100)
        covered.append(min(amount, receipt) / 100)
        counts.append(receipt_counts.get(ref, 0) if pd.notna(ref) else 0)
    inv["total_paid"] = received
    inv["payment_count"] = counts
    inv["balance_due"] = differences
    inv["positive_residual_usd"] = positives
    inv["excess_receipts_usd"] = excesses
    inv["analytical_coverage_usd"] = covered
    inv["control_flag"] = ""
    paid = inv["status"].astype(str).str.strip().str.lower().eq("paid")
    unsupported = paid & inv["positive_residual_usd"].gt(0)
    inv.loc[unsupported, "control_flag"] = "REVIEW: Paid status not supported by supplied receipt total"
    inv.loc[inv["excess_receipts_usd"].gt(0), "control_flag"] = "REVIEW: Excess receipts; no allocation or posting made"
    positive_receipts = inv["total_paid"].gt(0)
    fully = positive_receipts & inv["balance_due"].eq(0)
    partial = positive_receipts & inv["balance_due"].gt(0)
    excess = positive_receipts & inv["balance_due"].lt(0)
    no_positive = ~positive_receipts
    return {
        "matched": inv.loc[fully].sort_values("customer_name").copy(),
        "partial": inv.loc[partial].sort_values("balance_due", ascending=False).copy(),
        "unmatched_payments": unmatched.sort_values("payment_date", ascending=False).copy(),
        "unpaid_invoices": inv.loc[no_positive].sort_values("days_overdue", ascending=False).copy(),
        "overpayments": inv.loc[excess].sort_values("balance_due").copy(),
        "paid_status_exceptions": inv.loc[unsupported].copy(),
        "all_invoices": inv,
        "matched_receipts": matched_receipts,
        "source_invoices": invoices.copy().reset_index(drop=True),
        "source_payments": payments.copy().reset_index(drop=True),
    }


def _sum_cents(frame: pd.DataFrame, column: str) -> int:
    return sum(_to_cents(value) for value in frame[column])


def build_reconciliation_checks(results: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Compare actual report populations to independent validated source totals."""
    source_i, source_p = results["source_invoices"], results["source_payments"]
    inv, receipts, without = results["all_invoices"], results["matched_receipts"], results["unmatched_payments"]
    checks = []
    def add(label, unit, expected, actual, explanation):
        difference = actual - expected
        checks.append(dict(Check=label, Unit=unit,
                           Expected=expected / 100 if unit == "USD" else expected,
                           Actual=actual / 100 if unit == "USD" else actual,
                           Difference=difference / 100 if unit == "USD" else difference,
                           Status="PASS" if difference == 0 else "FAIL", Explanation=explanation))
    gross = _sum_cents(source_i, "amount_usd")
    all_receipts = _sum_cents(source_p, "amount_paid")
    reference_receipts = _sum_cents(receipts, "amount_paid")
    add("Receipt count completeness", "records", len(source_p), len(receipts) + len(without),
        "Reference-matched plus receipts without corresponding invoices.")
    expected_records = Counter(range(1, len(source_p) + 1))
    actual_records = Counter(list(receipts["receipt_data_row"]) + list(without["receipt_data_row"]))
    add("Receipt record identity", "errors", 0, sum((expected_records - actual_records).values()) + sum((actual_records - expected_records).values()),
        "Every source receipt row must appear exactly once, including missing references.")
    # Independently classify original source rows, not the derived receipt output.
    known = set(source_i["invoice_id"].dropna())
    expected_matched_rows = []
    expected_matched_amount = 0
    expected_by_invoice = {}
    for position, (_, row) in enumerate(source_p.iterrows(), start=1):
        ref = row["invoice_id"]
        if pd.notna(ref) and str(ref).strip() and ref in known:
            expected_matched_rows.append(position)
            cents = _to_cents(row["amount_paid"])
            expected_matched_amount += cents
            expected_by_invoice[ref] = expected_by_invoice.get(ref, 0) + cents
    add("Receipt reference membership", "errors", 0,
        sum((Counter(expected_matched_rows) - Counter(receipts["receipt_data_row"])).values()) +
        sum((Counter(receipts["receipt_data_row"]) - Counter(expected_matched_rows)).values()),
        "Reference matching proves existence only, not ERP application.")
    add("Total receipt amount completeness", "USD", all_receipts,
        reference_receipts + _sum_cents(without, "amount_paid"), "All supplied receipts retained, including excess receipts.")
    add("Reference-matched receipt amount", "USD", expected_matched_amount, reference_receipts,
        "Derived receipt population agrees with the original receipt references and amounts.")
    add("Original invoice amount completeness", "USD", gross, _sum_cents(inv, "amount_usd"),
        "Includes every invoice regardless of source status or receipt coverage.")
    invoice_records = Counter(inv["invoice_data_row"])
    add("Invoice record identity", "errors", 0,
        sum((Counter(range(1, len(source_i) + 1)) - invoice_records).values()) +
        sum((invoice_records - Counter(range(1, len(source_i) + 1))).values()),
        "All original invoice records retained exactly once.")
    add("Receipt totals linked to invoices", "USD", reference_receipts, _sum_cents(inv, "total_paid"),
        "Full intended receipt totals, including excess; no claim of accounting application.")
    association_errors = sum(
        _to_cents(row["total_paid"]) != expected_by_invoice.get(row["invoice_id"], 0)
        for _, row in inv.iterrows()
    )
    add("Per-invoice receipt association", "errors", 0, association_errors,
        "Receipt amounts agree to their original intended invoice; no cross-invoice offsets.")
    row_errors = sum(_to_cents(r["amount_usd"]) - _to_cents(r["total_paid"]) != _to_cents(r["balance_due"])
                     for _, r in inv.iterrows())
    add("Per-invoice signed difference", "errors", 0, row_errors,
        "Original amount minus reference-matched receipts equals signed difference, for each invoice.")
    add("Signed invoice difference total", "USD", gross - reference_receipts, _sum_cents(inv, "balance_due"),
        "No Paid-status deduction from the arithmetic population.")
    add("Positive versus excess split", "USD", _sum_cents(inv, "balance_due"),
        _sum_cents(inv, "positive_residual_usd") - _sum_cents(inv, "excess_receipts_usd"),
        "Positive residuals minus excess receipts; excess never offsets another invoice's positive exposure.")
    add("Analytical coverage plus excess", "USD", reference_receipts,
        _sum_cents(inv, "analytical_coverage_usd") + _sum_cents(inv, "excess_receipts_usd"),
        "Coverage is min(invoice, referenced receipts), not an approved allocation or ERP posting.")
    category_rows = []
    for key in ("matched", "partial", "unpaid_invoices", "overpayments"):
        category_rows += list(results[key]["invoice_data_row"])
    add("Invoice category membership", "errors", 0,
        sum((invoice_records - Counter(category_rows)).values()) + sum((Counter(category_rows) - invoice_records).values()),
        "Each invoice appears in exactly one coverage category; review tabs are overlapping views.")
    category_amount = sum(_sum_cents(results[k], "amount_usd") for k in ("matched", "partial", "unpaid_invoices", "overpayments"))
    add("Invoice category amount completeness", "USD", gross, category_amount,
        "Includes the original invoice amounts for excess-receipt cases.")
    return pd.DataFrame(checks)


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
    for row in ws:
        for cell in row:
            if isinstance(cell.value, str):cell.data_type = "s"
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



def generate_reconciliation_report(results: dict[str, pd.DataFrame], output_path: str | Path) -> None:
    """Write reference-matching analysis, never allocations or balance-sheet postings."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    checks = build_reconciliation_checks(results)
    failed = int(checks["Status"].eq("FAIL").sum())
    review_items = len(results["unmatched_payments"]) + len(results["overpayments"]) + len(results["paid_status_exceptions"])
    status = "FAIL—do not rely on financial results" if failed else ("REVIEW" if review_items else "PASS")
    inv, source_i, source_p = results["all_invoices"], results["source_invoices"], results["source_payments"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    # Summary is presentation-only: it uses the same validated reconciliation
    # results. It does not read an aging workbook or assert an external tie-out.
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 61
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 25

    def section(title, fourth=None):
        ws.append([])
        ws.append([title, "Invoice count" if fourth else "Count", "Original invoice amount" if fourth else "Amount (USD)", fourth])
        r = ws.max_row
        if not fourth:
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        _style_header_row(ws, 4, row_idx=r)
        for cell in ws[r]:
            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 32

    def amount_row(label, count, cents, highlight=False):
        ws.append([label, count, cents / 100, None])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(r, 2).number_format = INT_FORMAT
        ws.cell(r, 2).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(r, 3).number_format = CURRENCY_FORMAT
        ws.cell(r, 3).alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[r].height = 30
        if highlight:
            for cell in ws[r]:
                cell.font = SUBTOTAL_FONT
                cell.fill = SUBTOTAL_FILL

    def note(text, height=32):
        ws.append([text])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = height

    ws.append(["Reconciliation status", status, None, None])
    ws.merge_cells("B1:D1")
    _style_header_row(ws, 4)
    ws.cell(1, 1).alignment = Alignment(vertical="center")
    ws.cell(1, 2).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    ws.append(["Generated (UTC)", datetime.now(timezone.utc).isoformat(timespec="seconds")])
    ws.merge_cells("B2:D2")
    ws.append(["Failed checks", failed])
    ws.append(["Review items", review_items])
    for r in (3, 4):
        ws.cell(r, 2).number_format = INT_FORMAT
    note("PASS: tested relationships agree. REVIEW: measurable exceptions need review. Neither confirms ERP application or completeness of the source extracts.")

    paid_mask = inv["status"].astype(str).str.strip().str.lower().eq("paid")
    paid_invoices, nonpaid_invoices = inv.loc[paid_mask], inv.loc[~paid_mask]
    signed_all = _sum_cents(inv, "balance_due")
    signed_paid = _sum_cents(paid_invoices, "balance_due")
    signed_nonpaid = _sum_cents(nonpaid_invoices, "balance_due")

    section("1. Connection to AR aging")
    amount_row("Calculated signed invoice difference—all statuses", len(inv), signed_all)
    amount_row("Less: signed differences on invoices marked Paid", len(paid_invoices), -signed_paid)
    amount_row("Non-Paid balance—comparable to AR aging", len(nonpaid_invoices), signed_nonpaid, highlight=True)
    note("Bridge: all-status difference minus Paid-status difference = non-Paid balance. Counts include zero-balance invoices. This is comparable to signed non-Paid AR, not overdue AR or positive-only collections.", 34)
    note("Expected to match aging only with the same inputs and balance rules. The script does not open or independently check an aging workbook. Removing Paid differences here does not resolve or write off those exceptions.", 34)

    section("2. Receipt totals (receipt records, not invoice balances)")
    amount_row("Total supplied receipts", len(source_p), _sum_cents(source_p, "amount_paid"), highlight=True)
    amount_row("Invoice-reference-matched receipts", len(results["matched_receipts"]), _sum_cents(results["matched_receipts"], "amount_paid"))
    amount_row("Receipts Without Corresponding Invoices", len(results["unmatched_payments"]), _sum_cents(results["unmatched_payments"], "amount_paid"))
    note("Total supplied receipts = reference-matched receipts + receipts without invoices. Matching confirms only that the referenced invoice exists in this extract; it does not prove ERP application.")

    section("3. Invoice amounts and balance components")
    amount_row("Original invoice amounts—all source statuses", len(source_i), _sum_cents(source_i, "amount_usd"))
    amount_row("Positive calculated residuals—all statuses", int(inv["positive_residual_usd"].gt(0).sum()), _sum_cents(inv, "positive_residual_usd"))
    amount_row("Excess receipts requiring review", len(results["overpayments"]), _sum_cents(inv, "excess_receipts_usd"))
    amount_row("Analytical covered portion—not ERP application", int(inv["analytical_coverage_usd"].gt(0).sum()), _sum_cents(inv, "analytical_coverage_usd"))
    note("Signed difference = original invoice amounts minus reference-matched receipts. It also equals positive residuals minus excess receipts. These rows explain the balance; do not add them all together.", 34)
    note("Covered portion = the smaller of an invoice amount and its matching receipt total. Excess is kept separately; no cross-invoice allocation, accounting posting, write-off, credit memo, or refund is made.", 34)

    section("4. Invoice coverage categories (all source statuses)", "Remaining difference (signed)")
    for label, key in [("Fully covered by receipts", "matched"), ("Partial receipt coverage", "partial"),
                       ("No positive receipts", "unpaid_invoices"), ("Excess receipts", "overpayments")]:
        frame = results[key]
        ws.append([label, len(frame), _sum_cents(frame, "amount_usd") / 100, _sum_cents(frame, "balance_due") / 100])
    ws.append(["Total—all invoice categories", len(inv), _sum_cents(inv, "amount_usd") / 100, signed_all / 100])
    total_row = ws.max_row
    for r in range(total_row - 4, total_row + 1):
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(r, 2).number_format = INT_FORMAT
        for c in (3, 4):
            ws.cell(r, c).number_format = CURRENCY_FORMAT
        ws.row_dimensions[r].height = 26
    for cell in ws[total_row]:
        cell.font = SUBTOTAL_FONT
        cell.fill = SUBTOTAL_FILL
    note("Each invoice appears in one coverage category. Original amounts show the invoice value before receipts; remaining differences show what remains after receipts. Negative differences mean excess receipts, not approved credit postings.", 34)

    section("5. Paid-status exceptions (review subset only)")
    amount_row("Paid-status positive residuals—review subset", len(results["paid_status_exceptions"]), _sum_cents(results["paid_status_exceptions"], "positive_residual_usd"))
    note("These invoices are already included above. Do not add this review subset again or treat it automatically as collectible AR. If Paid invoices have excess receipts, their signed differences also affect the bridge and appear in Excess Receipts.", 34)
    note("Review items at the top counts receipts without invoices, excess-receipt invoices, and Paid-status positive residuals. The detail tabs show the records requiring investigation.")
    ws.append([])
    note("Report basis: all supplied matching receipts are used, regardless of receipt date. One-cent differences remain visible. The report does not reconstruct historical balances or determine balance-sheet account classification.", 34)
    check_ws = wb.create_sheet("Reconciliation Checks")
    check_ws.append(list(checks.columns))
    for values in checks.itertuples(index=False, name=None):
        check_ws.append(list(values))
    _style_header_row(check_ws, len(checks.columns))
    check_ws.auto_filter.ref = check_ws.dimensions
    for col, width in {"A":40,"B":12,"C":23,"D":23,"E":20,"F":12,"G":85}.items():
        check_ws.column_dimensions[col].width = width
    for row in check_ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for cell in row[2:5]:
            cell.number_format = CURRENCY_FORMAT if row[1].value == "USD" else INT_FORMAT
        check_ws.row_dimensions[row[0].row].height = 36
        if row[5].value == "FAIL":
            for cell in row:cell.fill = ALERT_FILL
    labels = {"invoice_data_row":"Invoice data row", "receipt_data_row":"Receipt data row", "invoice_id":"Invoice ID",
              "customer_name":"Customer name", "invoice_date":"Invoice date", "due_date":"Due date", "status":"Source status",
              "amount_usd":"Original invoice amount", "total_paid":"Reference-matched receipts", "payment_count":"Receipt record count",
              "balance_due":"Calculated signed difference", "positive_residual_usd":"Positive calculated residual",
              "excess_receipts_usd":"Excess receipts—review", "analytical_coverage_usd":"Analytical covered portion",
              "control_flag":"Review note", "payment_id":"Receipt ID", "payment_date":"Receipt date",
              "amount_paid":"Receipt amount", "unmatched_reason":"Review reason", "days_overdue":"Supplied days overdue", "po_number":"PO number"}
    invoice_columns = ["invoice_data_row", "invoice_id", "customer_name", "status", "amount_usd", "total_paid",
                       "balance_due", "positive_residual_usd", "excess_receipts_usd", "analytical_coverage_usd",
                       "payment_count", "days_overdue", "control_flag"]
    specs = [("Fully Covered", "matched"), ("Partial Coverage", "partial"),
             ("Receipts Without Invoices", "unmatched_payments"), ("No Positive Receipts", "unpaid_invoices"),
             ("Excess Receipts", "overpayments"), ("Paid Status Exceptions", "paid_status_exceptions")]
    currency_headers = {labels[c] for c in ("amount_usd", "total_paid", "balance_due", "positive_residual_usd", "excess_receipts_usd", "analytical_coverage_usd", "amount_paid")}
    count_headers = {labels[c] for c in ("invoice_data_row", "receipt_data_row", "payment_count", "days_overdue")}
    for title, key in specs:
        frame = results[key]
        if key == "unmatched_payments":
            columns = ["receipt_data_row", "payment_id", "invoice_id", "payment_date", "amount_paid", "unmatched_reason"]
        else:
            columns = invoice_columns
        # Preserve additional source fields (dates, PO numbers, etc.).
        columns = columns + [column for column in frame.columns if column not in columns]
        display = frame[columns].rename(columns=labels)
        _write_sheet(wb, title, display, "No records in this category.", highlight_flag_col="Review note")
        sheet = wb[title]
        if not display.empty:
            sheet.auto_filter.ref = sheet.dimensions
            sheet.row_dimensions[1].height = 42
            for cell in sheet[1]:cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            for col_index, name in enumerate(display.columns, start=1):
                if name in currency_headers:fmt = CURRENCY_FORMAT
                elif name in count_headers:fmt = INT_FORMAT
                else:continue
                for cells in sheet.iter_rows(min_row=2, min_col=col_index, max_col=col_index):cells[0].number_format = fmt
            for column in ("Review note", "Review reason"):
                if column in display:
                    col_index = list(display.columns).index(column) + 1
                    sheet.column_dimensions[get_column_letter(col_index)].width = 65
                    for cells in sheet.iter_rows(min_row=2, min_col=col_index, max_col=col_index):cells[0].alignment = Alignment(wrap_text=True, vertical="top")
                    for r in range(2, sheet.max_row + 1):sheet.row_dimensions[r].height = 30
    # Atomic replace prevents a partially written normal report on save failure.
    fd, temp_name = tempfile.mkstemp(prefix=".reconciliation_", suffix=".xlsx", dir=output_path.parent)
    os.close(fd)
    try:
        wb.save(temp_name)
        os.replace(temp_name, output_path)
    finally:
        if os.path.exists(temp_name):os.unlink(temp_name)


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
    for key in ("matched", "partial", "unmatched_payments", "unpaid_invoices", "overpayments", "paid_status_exceptions"):
        print(f"  {key}: {len(results[key])} rows")
    checks = build_reconciliation_checks(results)
    failures = int(checks["Status"].eq("FAIL").sum())
    if failures:
        failure_path = Path(args.output).with_name(Path(args.output).stem + "_failed_checks.xlsx")
        generate_reconciliation_report(results, failure_path)
        print(f"FAIL: {failures} control check(s) failed. Normal report NOT REFRESHED: {args.output}")
        print(f"Investigation workbook: {failure_path}")
        return 3

    print(f"Writing report to {args.output}...")
    generate_reconciliation_report(results, args.output)

    review_items = len(results["unmatched_payments"]) + len(results["overpayments"]) + len(results["paid_status_exceptions"])
    print(f"Control checks: PASS. Overall report status: {'REVIEW' if review_items else 'PASS'}.")
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
