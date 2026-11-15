"""
AR Aging Automation - core analysis script.

Loads an invoice extract (and optional payment receipts), validates data integrity
and cross-field controls, calculates net open AR aging buckets, flags high-risk
customers, and writes a formatted multi-sheet Excel report, including a separate
All Invoices, Payment Receipts, Paid Status Exceptions, and Credit Balances tabs. Collections and risk
use only positive balances not marked Paid. Source statuses are not overwritten.

Run from the project root:
    python src/ar_aging_automation.py
    python src/ar_aging_automation.py --payments data/payment_receipts.csv
    python src/ar_aging_automation.py --payments data/payment_receipts.csv --aging-date 2026-11-14

The aging-date option recalculates overdue days only, not historical balances.
Report Basis notes are generated at the bottom of Summary.
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
    df = pd.read_csv(file_path, dtype={"invoice_id": "string"})
    receipt_records = None
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

    # 6. Aging-only recalculation (--aging-date; --as-of-date remains an alias)
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
        pay = pd.read_csv(payments_path, dtype={"invoice_id": "string", "payment_id": "string"})
        # Preserve every supplied payment record before numeric coercion/aggregation.
        receipt_records = pay.to_dict("records")
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

    # DataFrame.copy() in calculate_aging_buckets preserves these source records.
    df.attrs["payment_receipt_records"] = receipt_records
    df.attrs["aging_date"] = as_of_date
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


def collectible_invoices(df: pd.DataFrame) -> pd.DataFrame:
    """Positive cent-rounded balances not marked Paid; not a credit-policy decision.

    Unsupported Paid records remain in Paid Status Exceptions. Missing/invalid
    balances are ineligible, with broader input validation handled separately.
    """
    is_paid = df["status"].astype(str).str.strip().str.lower().eq("paid")
    return df.loc[~is_paid & df["net_open_usd"].round(2).gt(0)].copy()


def flag_risk_accounts(
    open_df: pd.DataFrame,
    min_overdue_count: int = RISK_MIN_OVERDUE_COUNT,
    min_overdue_amount: float = RISK_MIN_OVERDUE_AMOUNT,
) -> pd.DataFrame:
    """Return one row per customer flagged as high-risk based on net open overdue balance."""
    eligible = collectible_invoices(open_df)
    overdue = eligible[eligible["days_overdue"] > 0]
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
    int_cols = {"days_overdue", "overdue_count", "max_days_overdue", "invoice_count", "payment_receipt_count"}
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



def build_all_invoices(df: pd.DataFrame, issues: list[dict], has_payments: bool) -> pd.DataFrame:
    """One row per input invoice record; no duplication from individual receipts.

    Duplicate invoice IDs, if supplied, remain separate flagged source records.
    A calculated positive balance is not confirmation of collectibility.
    """
    columns = ["invoice_id", "customer_name", "invoice_date", "due_date", "po_number", "status",
               "amount_usd", "amount_paid", "net_open_usd", "days_overdue", "aging_bucket"]
    result = df[columns].copy()
    paid = df["status"].astype(str).str.strip().str.lower().eq("paid")
    balance = df["net_open_usd"].round(2)
    result["review_category"] = "Invalid/missing balance—review validation notes"
    result.loc[~paid & balance.gt(0), "review_category"] = "Positive non-Paid balance—collections candidate"
    result.loc[balance.eq(0), "review_category"] = "Zero remaining balance"
    result.loc[paid, "review_category"] = (
        "Paid—supported by supplied payment total" if has_payments else
        "Paid—payment support not evaluated"
    )
    if has_payments:
        unsupported = paid & (df["payment_receipt_count"].eq(0) | balance.gt(0))
        result.loc[unsupported, "review_category"] = "Paid status—payment not supported by receipts"
    result.loc[balance.lt(0), "review_category"] = "Credit balance—review, not collections"
    notes = {}
    for issue in issues:
        key = str(issue.get("invoice_id", ""))
        text = f"{issue.get('field', '')}: {issue.get('problem', '')}"
        if text not in notes.setdefault(key, []):
            notes[key].append(text)
    result["validation_notes"] = result["invoice_id"].map(lambda x: " | ".join(notes.get(str(x), [])))
    result["payment_support_scope"] = (
        "Supplied receipts only; positive balance is not proof of collectibility"
        if has_payments else
        "Not evaluated—no receipts file supplied; amount_paid defaults to zero"
    )
    return result


def build_payment_receipts(df: pd.DataFrame, records: list[dict]) -> pd.DataFrame:
    """Keep receipt-level grain, including orphans and the original amount values.

    Matched means only that the invoice reference exists, not that the receipt
    is valid/unique. Payment dates are preserved as supplied, not cutoff-filtered.
    """
    columns = ["payment_id", "invoice_id", "payment_date", "amount_paid"]
    pay = pd.DataFrame(records, columns=columns)
    known_ids = set(df["invoice_id"].dropna())
    # Lookup without a merge: duplicate invoice rows must not multiply receipts.
    names = {}
    for invoice_id, customer in zip(df["invoice_id"], df["customer_name"]):
        if pd.notna(invoice_id) and pd.notna(customer):
            names.setdefault(invoice_id, set()).add(str(customer))
    def customer_for(invoice_id):
        customers = names.get(invoice_id, set())
        if len(customers) == 1:
            return next(iter(customers))
        return "Multiple customers—review invoice IDs" if customers else None
    matched = pay["invoice_id"].notna() & pay["invoice_id"].isin(known_ids)
    pay["customer_name"] = pay["invoice_id"].map(customer_for)
    pay["invoice_match_result"] = matched.map({True: "Matched", False: "No matching invoice"})
    # Store valid dates as Excel dates, retaining invalid source text for review.
    parsed_dates = pd.to_datetime(pay["payment_date"], format="%Y-%m-%d", errors="coerce")
    pay["payment_date"] = [parsed if pd.notna(parsed) else original
                           for original, parsed in zip(pay["payment_date"], parsed_dates)]
    return pay


def write_reference_tab(wb, name: str, data: pd.DataFrame) -> None:
    """Filterable working sheet, preserving values rather than hiding bad inputs."""
    ws = wb.create_sheet(name)
    _write_dataframe(ws, data)
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 34
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for column in ("review_category", "validation_notes", "payment_support_scope"):
        if column in data.columns:
            idx = data.columns.get_loc(column) + 1
            ws.column_dimensions[get_column_letter(idx)].width = 60 if column == "validation_notes" else 46
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                row[0].alignment = Alignment(wrap_text=True, vertical="top")
    if name == "All Invoices":
        for row_idx in range(2, len(data) + 2):
            ws.row_dimensions[row_idx].height = 60
    ws.freeze_panes = "C2" if name == "All Invoices" else "A2"



def write_report_basis(ws, has_payments: bool, aging_date: Optional[str] = None) -> None:
    """Append explanatory notes below Summary, without shifting existing tables."""
    if aging_date:
        aging_basis = (
            f"Overdue days recalculated from due dates as of {aging_date}. "
            "If a due date is unavailable, supplied overdue days are retained. "
            "See Validation Issues for detected date problems."
        )
    else:
        aging_basis = (
            "Uses overdue days supplied in the invoice CSV. "
            "No reporting date has been established for those values."
        )
    if has_payments:
        payment_basis = "Uses all supplied matching receipts, regardless of payment date."
        balance_basis = "Original invoice amount minus all supplied matching payments."
        paid_treatment = (
            "Invoices marked Paid are excluded from collections and risk. "
            "Missing or insufficient receipt support is listed in Paid Status Exceptions."
        )
    else:
        payment_basis = "Not evaluated—no payment file supplied; no receipts deducted."
        balance_basis = (
            "Calculated remaining balances equal original invoice amounts because "
            "recorded payments default to zero. Actual payment-adjusted balances are not established."
        )
        paid_treatment = (
            "Invoices marked Paid are excluded from collections and risk. "
            "Their payment support is not evaluated without a receipts file."
        )
    ws.append([])
    ws.append(["Report Basis", "Description"])
    header_row = ws.max_row
    ws.merge_cells(start_row=header_row, start_column=2, end_row=header_row, end_column=3)
    _style_header_row(ws, 3, row_idx=header_row)
    for label, text in (
        ("Aging basis", aging_basis),
        ("Payment basis", payment_basis),
        ("Balance calculation", balance_basis),
        ("Paid-status treatment", paid_treatment),
        ("Limitation", "This workbook does not reconstruct historical balances as of a specified date."),
    ):
        ws.append([label, text])
        row = ws.max_row
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row, 1).font = Font(bold=True)
        ws.cell(row, 1).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row, 2).alignment = Alignment(vertical="top", wrap_text=True)
        # Existing Summary widths remain unchanged; allow space for wrapped notes.
        ws.row_dimensions[row].height = 18 * max(2, (len(text) + 69) // 70) + 6


def generate_summary_report(
    df: pd.DataFrame,
    issues: list[dict],
    output_path: str | Path,
    has_payments: bool = False,
    payment_receipt_records: Optional[list[dict]] = None,
    aging_date: Optional[str] = None,
) -> None:
    """Write eight review tabs. Receipts are preserved at source-record grain.

    Invoice and payment tabs are alternative views, not amounts to add together.
    Derived values are generated by this script; edits in Excel do not recalculate them.
    Historical cutoff handling and broader validation remain separate work.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    is_paid = df["status"].astype(str).str.strip().str.lower() == "paid"
    paid_df = df[is_paid].copy()
    # Retain signed non-Paid balances for the existing accounting summary.
    open_df = df[~is_paid].copy()
    collectible_df = collectible_invoices(df)
    credit_df = df.loc[df["net_open_usd"].round(2).lt(0)].copy()
    nonpaid_credits = open_df.loc[open_df["net_open_usd"].round(2).lt(0)]
    zero_nonpaid = open_df.loc[open_df["net_open_usd"].round(2).eq(0)]
    collectible_ar = float(collectible_df["net_open_usd"].sum())

    gross_billed = float(df["amount_usd"].sum())
    settled_paid_invoices = float(paid_df["amount_usd"].sum())
    partial_payments_applied = float(open_df["amount_paid"].sum())
    net_open_ar = float(open_df["net_open_usd"].sum())

    overdue_df = collectible_df[collectible_df["days_overdue"] > 0]
    total_overdue = float(overdue_df["net_open_usd"].sum())
    pct_overdue = (total_overdue / collectible_ar * 100) if collectible_ar else 0.0

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
        ws_summary.append(["Less: Payments on non-Paid invoices", round(-partial_payments_applied, 2), partial_count])
        net_row_idx = 5
    else:
        net_row_idx = 4

    ws_summary.append(["Net AR: non-Paid, including credits", round(net_open_ar, 2), len(open_df)])
    ws_summary.append(["Total Overdue (>0 Days)", round(total_overdue, 2), len(overdue_df)])
    ws_summary.append(["% of Collectible AR Overdue", pct_overdue / 100, None])
    ws_summary.append([])  # spacer

    for r in range(2, net_row_idx + 2):
        ws_summary.cell(row=r, column=2).number_format = CURRENCY_FORMAT
        ws_summary.cell(row=r, column=3).number_format = INT_FORMAT
    ws_summary.cell(row=net_row_idx + 2, column=2).number_format = "0.0%"
    ws_summary.cell(row=net_row_idx + 2, column=2).alignment = Alignment(horizontal="right")

    for col_idx in (1, 2, 3):
        cell = ws_summary.cell(row=net_row_idx, column=col_idx)
        cell.font = SUBTOTAL_FONT
        cell.fill = SUBTOTAL_FILL

    bucket_header_row = net_row_idx + 4
    ws_summary.append(["Aging: non-Paid, including credits", "Invoice Count", "Net Open Amount"])
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
    ws_summary.append(["Total non-Paid AR (signed)", int(bucket_summary["invoice_count"].sum()), round(float(bucket_summary["total_amount"].sum()), 2)])
    ws_summary.cell(row=total_row_idx, column=2).number_format = INT_FORMAT
    ws_summary.cell(row=total_row_idx, column=3).number_format = CURRENCY_FORMAT
    for col_idx in (1, 2, 3):
        cell = ws_summary.cell(row=total_row_idx, column=col_idx)
        cell.font = SUBTOTAL_FONT
        cell.fill = SUBTOTAL_FILL

    # Explain the collections population without hiding zero/credit balances.
    ws_summary.append([])
    ws_summary.append(["Collections eligibility: non-Paid", "Amount", "Invoice Count"])
    _style_header_row(ws_summary, 3, row_idx=ws_summary.max_row)
    for label, population in (
        ("Positive remaining balances", collectible_df),
        ("Credit balances (signed)", nonpaid_credits),
        ("Zero remaining balances", zero_nonpaid),
    ):
        ws_summary.append([label, round(float(population["net_open_usd"].sum()), 2), len(population)])
        ws_summary.cell(ws_summary.max_row, 2).number_format = CURRENCY_FORMAT
        ws_summary.cell(ws_summary.max_row, 3).number_format = INT_FORMAT
    ws_summary.append([
        "Only positive non-Paid balances enter collections/risk. Signed AR and aging above "
        "retain credits and zero-balance records; the overdue percentage uses positive balances only."
    ])
    eligibility_note = ws_summary.max_row
    ws_summary.merge_cells(start_row=eligibility_note, start_column=1, end_row=eligibility_note, end_column=3)
    ws_summary.cell(eligibility_note, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws_summary.row_dimensions[eligibility_note].height = 60

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
    effective_aging_date = aging_date if aging_date is not None else df.attrs.get("aging_date")
    write_report_basis(ws_summary, has_payments, effective_aging_date)

    # ----- Sheet 2: Detail >30d -----
    detail_df = collectible_df[collectible_df["days_overdue"] > 30].sort_values("days_overdue", ascending=False)
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
        ws_detail.append(["No collectible invoices >30 days overdue"])

    # ----- Sheet 3: Risk Accounts -----
    risk_df = flag_risk_accounts(collectible_df)
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

    write_reference_tab(wb, "All Invoices", build_all_invoices(df, issues, has_payments))
    if has_payments:
        records = payment_receipt_records
        if records is None:
            records = df.attrs.get("payment_receipt_records")
        if records is None:
            raise ValueError("Receipt-level records are required to create Payment Receipts.")
        write_reference_tab(wb, "Payment Receipts", build_payment_receipts(df, records))
    else:
        ws_receipts = wb.create_sheet("Payment Receipts")
        ws_receipts.append(["No payment file supplied—not evaluated."])
        ws_receipts.column_dimensions["A"].width = 65

    write_paid_status_exceptions(wb, exceptions, has_payments)
    # Preserve every negative balance, including Paid-status overpayments.
    ws_credit = wb.create_sheet("Credit Balances")
    if not credit_df.empty:
        credit_columns = ["invoice_id", "customer_name", "invoice_date", "due_date", "status",
                          "amount_usd", "amount_paid", "net_open_usd", "aging_bucket"]
        _write_dataframe(ws_credit, credit_df[credit_columns])
        ws_credit.auto_filter.ref = ws_credit.dimensions
    else:
        ws_credit.append(["No negative remaining balances identified."])
    # Public move_sheet API, preserving sheet content and formatting.
    desired_order = ["Summary", "All Invoices", "Payment Receipts", "Detail >30d", "Risk Accounts",
                     "Paid Status Exceptions", "Credit Balances", "Validation Issues"]
    for target, name in enumerate(desired_order):
        wb.move_sheet(name, offset=target - wb.sheetnames.index(name))
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
        "--aging-date", "--as-of-date",
        dest="aging_date",
        default=None,
        help=("Optional aging date (YYYY-MM-DD) to recalculate overdue days only; "
              "does not filter invoices or receipts to historical balances. "
              "--as-of-date is retained as an alias."),
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
        as_of_date=args.aging_date,
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
        aging_date=args.aging_date,
    )
    print("Done. Open the workbook to review the report.")


if __name__ == "__main__":
    main()
