"""
Generate realistic mock AR data for the AR Automation Toolkit.

Outputs two CSVs into the data/ folder:
  - invoices_sample.csv:  300 invoices with mixed statuses and aging
  - payment_receipts.csv: ~180 payment records (some partial, some orphan)

Intentionally seeds bad data so validation logic has something to flag:
  - 7 rows missing PO numbers
  - 3 rows with malformed date strings
  - 4 orphan payments that don't match any invoice
"""

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

from faker import Faker

# Reproducibility: same seed = same data every run
fake = Faker()
Faker.seed(42)
random.seed(42)

# ------- Config -------
NUM_INVOICES = 300
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data"
TODAY = datetime(2026, 4, 26)  # anchor "today" so aging is reproducible

STATUSES = ["Paid", "Open", "Overdue"]
STATUS_WEIGHTS = [0.60, 0.25, 0.15]  # ~60% paid, 25% open, 15% overdue


def random_invoice_date() -> datetime:
    """Random date within the past 150 days."""
    days_ago = random.randint(0, 150)
    return TODAY - timedelta(days=days_ago)


def generate_invoices(n: int) -> list[dict]:
    """Generate n invoice records."""
    customers = [fake.company() for _ in range(40)]  # ~40 unique customers
    invoices = []

    for i in range(1, n + 1):
        invoice_date = random_invoice_date()
        # Net-30 terms
        due_date = invoice_date + timedelta(days=30)
        status = random.choices(STATUSES, weights=STATUS_WEIGHTS, k=1)[0]

        # Days overdue: positive if past due, 0 if not yet due or paid
        if status == "Overdue":
            days_overdue = (TODAY - due_date).days
            if days_overdue < 1:
                days_overdue = random.randint(1, 120)
        elif status == "Open" and TODAY > due_date:
            days_overdue = (TODAY - due_date).days
        else:
            days_overdue = 0

        invoice = {
            "invoice_id": f"INV-{10000 + i}",
            "customer_name": random.choice(customers),
            "invoice_date": invoice_date.strftime("%Y-%m-%d"),
            "due_date": due_date.strftime("%Y-%m-%d"),
            "amount_usd": round(random.uniform(500, 25_000), 2),
            "status": status,
            "days_overdue": days_overdue,
            "po_number": f"PO-{random.randint(100000, 999999)}",
        }
        invoices.append(invoice)

    # ----- Inject intentional bad data for validation testing -----
    # 7 rows with missing PO numbers
    missing_po_indices = random.sample(range(len(invoices)), 7)
    for idx in missing_po_indices:
        invoices[idx]["po_number"] = ""

    # 3 rows with malformed dates
    bad_date_indices = random.sample(range(len(invoices)), 3)
    bad_formats = ["04/15/2026", "2026-13-01", "Apr 15, 2026"]
    for idx, bad_fmt in zip(bad_date_indices, bad_formats):
        invoices[idx]["invoice_date"] = bad_fmt

    return invoices


def generate_payments(invoices: list[dict]) -> list[dict]:
    """Generate payment receipts for paid + some open invoices."""
    payments = []
    payment_id = 1

    for inv in invoices:
        # Skip rows with bad dates so we don't try to parse them
        try:
            invoice_date = datetime.strptime(inv["invoice_date"], "%Y-%m-%d")
        except ValueError:
            continue

        if inv["status"] == "Paid":
            # Full payment, sometime between invoice date and today
            days_to_pay = random.randint(5, 45)
            payment_date = invoice_date + timedelta(days=days_to_pay)
            payments.append({
                "payment_id": f"PMT-{20000 + payment_id}",
                "invoice_id": inv["invoice_id"],
                "payment_date": payment_date.strftime("%Y-%m-%d"),
                "amount_paid": inv["amount_usd"],
            })
            payment_id += 1

        elif inv["status"] == "Open" and random.random() < 0.20:
            # 20% of open invoices have a partial payment
            partial = round(inv["amount_usd"] * random.uniform(0.2, 0.7), 2)
            payment_date = invoice_date + timedelta(days=random.randint(10, 25))
            payments.append({
                "payment_id": f"PMT-{20000 + payment_id}",
                "invoice_id": inv["invoice_id"],
                "payment_date": payment_date.strftime("%Y-%m-%d"),
                "amount_paid": partial,
            })
            payment_id += 1

    # Add 4 orphan payments (no matching invoice)
    for _ in range(4):
        payments.append({
            "payment_id": f"PMT-{20000 + payment_id}",
            "invoice_id": f"INV-{random.randint(99000, 99999)}",  # bogus
            "payment_date": (TODAY - timedelta(days=random.randint(5, 40))).strftime("%Y-%m-%d"),
            "amount_paid": round(random.uniform(500, 5000), 2),
        })
        payment_id += 1

    return payments


def write_csv(rows: list[dict], output_path: Path) -> None:
    """Write a list of dicts to CSV using the first row's keys as headers."""
    if not rows:
        return
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    print(f"Generating {NUM_INVOICES} invoices...")
    invoices = generate_invoices(NUM_INVOICES)

    print("Generating payment receipts...")
    payments = generate_payments(invoices)

    invoices_path = OUTPUT_DIR / "invoices_sample.csv"
    payments_path = OUTPUT_DIR / "payment_receipts.csv"

    write_csv(invoices, invoices_path)
    write_csv(payments, payments_path)

    print(f"✅ Wrote {len(invoices)} invoices to {invoices_path}")
    print(f"✅ Wrote {len(payments)} payments to {payments_path}")


if __name__ == "__main__":
    main()