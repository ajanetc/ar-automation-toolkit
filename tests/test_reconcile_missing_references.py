"""Place in tests/. Run: python -m unittest discover -s tests -p 'test_reconcile_*.py' -v"""
import importlib.util
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

root = Path(__file__).resolve().parent
script = root.parent / "src" / "reconcile_payments.py"
if not script.exists():
    script = root / "reconcile_payments.py"
spec = importlib.util.spec_from_file_location("missing_reference_under_test", script)
recon = importlib.util.module_from_spec(spec)
spec.loader.exec_module(recon)


def invoices(ids=("INV-1",)):
    return pd.DataFrame([dict(invoice_id=i, customer_name="Test Customer", amount_usd=100,
                              days_overdue=40, status="Open") for i in ids])


def receipts(refs, amounts=None):
    amounts = [10] * len(refs) if amounts is None else amounts
    return pd.DataFrame([dict(invoice_id=ref, payment_id=f"P{n}", amount_paid=amount,
                              payment_date="2026-01-01")
                         for n, (ref, amount) in enumerate(zip(refs, amounts), start=1)],
                        columns=["invoice_id", "payment_id", "amount_paid", "payment_date"])


class MissingReferenceTests(unittest.TestCase):
    def test_null_empty_and_whitespace_references_are_preserved(self):
        pay = receipts([None, pd.NA, "", " ", "\t"])
        result = recon.reconcile_payments(invoices(), pay)
        unmatched = result["unmatched_payments"]
        self.assertEqual(len(unmatched), len(pay))
        self.assertEqual(set(unmatched.payment_id), set(pay.payment_id))
        self.assertTrue(unmatched.unmatched_reason.eq("Missing invoice reference").all())
        self.assertEqual(unmatched.amount_paid.sum(), 50)
        self.assertEqual(len(result["unpaid_invoices"]), 1)

    def test_unknown_reference_has_distinct_reason(self):
        result = recon.reconcile_payments(invoices(), receipts([None, "UNKNOWN"]))
        byid = result["unmatched_payments"].set_index("payment_id")
        self.assertEqual(byid.loc["P1", "unmatched_reason"], "Missing invoice reference")
        self.assertEqual(byid.loc["P2", "unmatched_reason"], "Invoice reference not found")

    def test_known_multiple_receipts_still_apply_once_each(self):
        result = recon.reconcile_payments(invoices(), receipts(["INV-1", "INV-1", None, "UNKNOWN"], [20, 30, 7, 9]))
        row = result["partial"].iloc[0]
        self.assertEqual(row.total_paid, 50)
        self.assertEqual(row.payment_count, 2)
        self.assertEqual(row.balance_due, 50)
        self.assertEqual(len(result["unmatched_payments"]), 2)
        self.assertEqual(result["unmatched_payments"].amount_paid.sum(), 16)

    def test_blank_invoice_record_cannot_absorb_blank_receipt(self):
        # Invoice-file missing-ID validation is deferred; do not apply receipts to it.
        for ref in (None, "", " "):
            with self.subTest(ref=ref):
                result = recon.reconcile_payments(invoices(("INV-1", ref)), receipts([ref], [25]))
                self.assertEqual(len(result["unmatched_payments"]), 1)
                self.assertEqual(len(result["unpaid_invoices"]), 2)
                self.assertTrue(result["partial"].empty)

    def test_nonblank_id_not_silently_trimmed(self):
        pay = receipts([" INV-1 "], [25])
        row = recon.reconcile_payments(invoices(), pay)["unmatched_payments"].iloc[0]
        self.assertEqual(row.invoice_id, " INV-1 ")
        self.assertEqual(row.unmatched_reason, "Invoice reference not found")

    def test_source_frames_are_not_modified(self):
        inv, pay = invoices(), receipts([None, "INV-1"])
        inv_before, pay_before = inv.copy(deep=True), pay.copy(deep=True)
        recon.reconcile_payments(inv, pay)
        pd.testing.assert_frame_equal(inv, inv_before)
        pd.testing.assert_frame_equal(pay, pay_before)

    def test_empty_payment_file_keeps_all_invoices_unpaid(self):
        result = recon.reconcile_payments(invoices(), receipts([]))
        self.assertTrue(result["unmatched_payments"].empty)
        self.assertIn("unmatched_reason", result["unmatched_payments"].columns)
        self.assertEqual(len(result["unpaid_invoices"]), 1)

    def test_distinct_unmatched_receipts_with_same_reference_and_amount_are_not_dropped(self):
        pay = receipts([None], [25])
        pay = pd.concat([pay, pay], ignore_index=True)
        pay["payment_id"] = ["P-UNMATCHED-1", "P-UNMATCHED-2"]
        unmatched = recon.reconcile_payments(invoices(), pay)["unmatched_payments"]
        self.assertEqual(len(unmatched), 2)
        self.assertEqual(unmatched.amount_paid.sum(), 50)

    def test_summary_accounts_for_missing_reference_cash_and_counts(self):
        pay = receipts(["INV-1", None, "UNKNOWN", " "], [40, 7, 9, 11])
        result = recon.reconcile_payments(invoices(), pay)
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "report.xlsx"
            recon.generate_reconciliation_report(result, output)
            wb = load_workbook(output, data_only=True)
            summary = {row[0]: row[1:3] for row in wb["Summary"].values if row[0]}
            self.assertEqual(summary["Total supplied receipts"], (4, 67))
            self.assertEqual(summary["Receipts Without Corresponding Invoices"], (3, 27))
            self.assertEqual(summary["Invoice-reference-matched receipts"], (1, 40))
            self.assertEqual(wb["Receipts Without Invoices"].max_row - 1, 3)
            self.assertIn("Review reason", [c.value for c in wb["Receipts Without Invoices"][1]])
            # Independent source arithmetic, not the report's category sums.
            self.assertEqual(Decimal(str(summary["Total supplied receipts"][1])),
                             sum((Decimal(str(x)) for x in pay.amount_paid), Decimal(0)))


if __name__ == "__main__":
    unittest.main()
