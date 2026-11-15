"""Run from the project root: python -m unittest discover -s tests -p 'test_reconcile_invoice_ids.py' -v"""
import importlib.util
import sys
import unittest
from pathlib import Path
import pandas as pd

# Put this file in tests/ and the revised script in src/.
# Also supports running beside the script for isolated review.
root = Path(__file__).resolve().parent
script = root.parent / "src" / "reconcile_payments.py"
if not script.exists():
    script = root / "reconcile_payments.py"
spec = importlib.util.spec_from_file_location("reconcile_under_test", script)
recon = importlib.util.module_from_spec(spec)
spec.loader.exec_module(recon)


class InvoiceIDValidationTests(unittest.TestCase):
    def test_unique_ids_allowed(self):
        frame = pd.DataFrame({"invoice_id": ["INV-1", "INV-2"]})
        recon.validate_unique_invoice_ids(frame)

    def test_repeated_ids_identified_with_all_data_rows(self):
        frame = pd.DataFrame({"invoice_id": ["INV-1", "INV-2", "INV-1", "INV-2"]}, index=[8, 9, 21, 40])
        with self.assertRaises(recon.DuplicateInvoiceIDError) as caught:
            recon.validate_unique_invoice_ids(frame)
        message = str(caught.exception)
        self.assertIn("'INV-1': data rows 1, 3", message)
        self.assertIn("'INV-2': data rows 2, 4", message)
        self.assertIn("legitimate invoice/product lines", message)

    def test_product_lines_are_not_automatically_summed(self):
        frame = pd.DataFrame({"invoice_id": ["INV-1", "INV-1"], "amount_usd": [100, 200]})
        original = frame.copy(deep=True)
        with self.assertRaises(recon.DuplicateInvoiceIDError):
            recon.validate_unique_invoice_ids(frame)
        pd.testing.assert_frame_equal(frame, original)

    def test_check_runs_before_any_payment_processing(self):
        invoices = pd.DataFrame({"invoice_id": ["INV-1", "INV-1"]})
        # Deliberately no amount fields: duplicate validation must happen first.
        with self.assertRaises(recon.DuplicateInvoiceIDError):
            recon.reconcile_payments(invoices, pd.DataFrame())

    def test_multiple_receipts_for_one_invoice_are_allowed(self):
        invoices = pd.DataFrame({"invoice_id": ["INV-1"], "amount_usd": [100],
                                 "days_overdue": [30], "customer_name": ["Test"], "status": ["Open"]})
        payments = pd.DataFrame({"invoice_id": ["INV-1", "INV-1"], "payment_id": ["P1", "P2"],
                                 "amount_paid": [20, 30], "payment_date": ["2026-01-01", "2026-01-02"]})
        result = recon.reconcile_payments(invoices, payments)
        self.assertEqual(len(result["partial"]), 1)
        self.assertEqual(result["partial"].iloc[0]["total_paid"], 50)

    def test_leading_zero_identifiers_remain_distinct(self):
        recon.validate_unique_invoice_ids(pd.DataFrame({"invoice_id": ["001", "1"]}))

    def test_missing_and_blank_ids_are_outside_this_check(self):
        # This check does not approve missing IDs; their validation is a later task.
        recon.validate_unique_invoice_ids(pd.DataFrame({"invoice_id": [None, None, "", " "]}))


if __name__ == "__main__":
    unittest.main()
