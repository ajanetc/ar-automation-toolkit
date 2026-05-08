"""
Convert the mock CSV data into a single Excel workbook with two sheets.

Reads:
  - data/invoices_sample.csv
  - data/payment_receipts.csv

Writes:
  - output/mock_data.xlsx  (sheets: Invoices, Payments)
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PROJECT_ROOT = Path(__file__).resolve().parent.parent
INVOICES_CSV = PROJECT_ROOT / "data" / "invoices_sample.csv"
PAYMENTS_CSV = PROJECT_ROOT / "data" / "payment_receipts.csv"
OUTPUT_PATH = PROJECT_ROOT / "output" / "mock_data.xlsx"

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")


def write_sheet(ws, df: pd.DataFrame) -> None:
    """Write a DataFrame to a worksheet with styled headers and autosized columns."""
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))

    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=10,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 40)


def main() -> None:
    print(f"Loading {INVOICES_CSV.name}...")
    invoices = pd.read_csv(INVOICES_CSV)
    print(f"  {len(invoices)} rows")

    print(f"Loading {PAYMENTS_CSV.name}...")
    payments = pd.read_csv(PAYMENTS_CSV)
    print(f"  {len(payments)} rows")

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    ws_invoices = wb.active
    ws_invoices.title = "Invoices"
    write_sheet(ws_invoices, invoices)

    ws_payments = wb.create_sheet("Payments")
    write_sheet(ws_payments, payments)

    wb.save(OUTPUT_PATH)
    print(f"✅ Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()