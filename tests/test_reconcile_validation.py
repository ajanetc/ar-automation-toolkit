"""Place in tests/. Run: python -m unittest discover -s tests -p 'test_reconcile_*.py' -v"""
import contextlib
import importlib.util
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from openpyxl import load_workbook

root = Path(__file__).resolve().parent
script = root.parent / "src" / "reconcile_payments.py"
if not script.exists():
    script = root / "reconcile_payments.py"
spec = importlib.util.spec_from_file_location("validation_under_test", script)
recon = importlib.util.module_from_spec(spec)
spec.loader.exec_module(recon)


def invoice_frame(amount="100"):
    return pd.DataFrame([dict(invoice_id="001", customer_name="Test Customer", amount_usd=amount,
                              days_overdue="40", status="Open")])


def payment_frame(amount="40", reference="001"):
    return pd.DataFrame([dict(payment_id="P1", invoice_id=reference, amount_paid=amount,
                              payment_date="2026-01-01")])


class CriticalValidationTests(unittest.TestCase):
    def make_inputs(self, folder, invoice=None, payment=None):
        folder = Path(folder)
        invoice = invoice_frame() if invoice is None else invoice
        payment = payment_frame() if payment is None else payment
        ip, pp = folder / "invoices.csv", folder / "payments.csv"
        invoice.to_csv(ip, index=False)
        payment.to_csv(pp, index=False)
        return ip, pp

    def run_cli(self, folder, invoice=None, payment=None, name="reconciliation_report.xlsx"):
        ip, pp = self.make_inputs(folder, invoice, payment)
        out = Path(folder) / name
        proc = subprocess.run([sys.executable, str(script), "--invoices", str(ip),
                               "--payments", str(pp), "-o", str(out)], capture_output=True, text=True)
        return proc, out, recon.diagnostic_output_path(out)

    def test_valid_positive_zero_and_overpayment_inputs_allowed(self):
        for ia, pa in [(100, 40), (0, 0), (100, 120), ("100.50", "0.00")]:
            with self.subTest(ia=ia, pa=pa):
                self.assertEqual(recon.collect_input_issues(invoice_frame(ia), payment_frame(pa)), [])

    def test_missing_and_unreadable_amounts_are_not_zero(self):
        for value in [None, pd.NA, "", " ", "abc", "1,000.00", True]:
            with self.subTest(value=value):
                issues = recon.collect_input_issues(invoice_frame(value), payment_frame(value))
                self.assertEqual(len(issues), 2)
                self.assertEqual({x["source"] for x in issues}, {"Invoices", "Payments"})

    def test_nonfinite_values_blocked(self):
        for value in [float("nan"), float("inf"), "NaN", "Infinity", "-inf", "1e999"]:
            with self.subTest(value=value):
                self.assertTrue(recon.collect_input_issues(invoice_frame(value), payment_frame()))

    def test_negative_amounts_blocked_with_explanation(self):
        issues = recon.collect_input_issues(invoice_frame(-100), payment_frame(-40))
        self.assertEqual(len(issues), 2)
        self.assertTrue(all("credit/reversal" in x["problem"] for x in issues))

    def test_missing_required_columns_identified(self):
        issues = recon.collect_input_issues(invoice_frame().drop(columns="amount_usd"),
                                             payment_frame().drop(columns="payment_id"))
        self.assertEqual({x["field"] for x in issues}, {"amount_usd", "payment_id"})
        self.assertTrue(all(x["data_row"] is None for x in issues))

    def test_collects_all_errors_without_mutating_sources(self):
        inv = pd.concat([invoice_frame("abc"), invoice_frame(-10)], ignore_index=True)
        pay = payment_frame("NaN")
        original_i, original_p = inv.copy(deep=True), pay.copy(deep=True)
        issues = recon.collect_input_issues(inv, pay)
        self.assertEqual(len(issues), 5)  # 3 amounts plus both repeated-ID rows
        pd.testing.assert_frame_equal(inv, original_i)
        pd.testing.assert_frame_equal(pay, original_p)

    def test_library_call_blocks_before_payment_application(self):
        with self.assertRaises(recon.InputValidationError) as caught:
            recon.reconcile_payments(invoice_frame(), payment_frame("abc"))
        self.assertEqual(caught.exception.issues[0]["field"], "amount_paid")

    def test_cli_diagnostics_preserve_source_and_previous_report(self):
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder) / "reconciliation_report.xlsx"
            out.write_bytes(b"PREVIOUS REPORT MUST REMAIN UNCHANGED")
            inv = invoice_frame("=1+1")
            inv.loc[0, "customer_name"] = "=HYPERLINK(\"https://example.com\",\"not a formula\")"
            proc, report, diagnostic = self.run_cli(folder, inv, payment_frame("NaN"))
            self.assertEqual(proc.returncode, 2, proc.stderr)
            self.assertIn("NOT REFRESHED", proc.stdout)
            self.assertEqual(report.read_bytes(), b"PREVIOUS REPORT MUST REMAIN UNCHANGED")
            wb = load_workbook(diagnostic, data_only=False)
            self.assertEqual(wb.sheetnames, ["Summary", "Input Issues", "Source Invoices", "Source Payments"])
            self.assertIn("BLOCKED", wb["Summary"]["B1"].value)
            self.assertEqual(wb["Input Issues"].max_row - 1, 2)
            source = wb["Source Invoices"]
            cols = {c.value: c.column for c in source[1]}
            self.assertEqual(source.cell(2, cols["amount_usd"]).value, "=1+1")
            self.assertEqual(source.cell(2, cols["amount_usd"]).data_type, "s")
            self.assertEqual(source.cell(2, cols["invoice_id"]).value, "001")
            self.assertEqual(wb["Source Payments"]["C2"].value, "NaN")
            self.assertFalse(any(c.data_type == "f" for ws in wb for row in ws for c in row))

    def test_blocked_run_does_not_create_normal_report(self):
        with tempfile.TemporaryDirectory() as folder:
            proc, report, diagnostic = self.run_cli(folder, payment=payment_frame(""))
            self.assertEqual(proc.returncode, 2)
            self.assertFalse(report.exists())
            self.assertTrue(diagnostic.exists())

    def test_duplicate_ids_produce_diagnostics_too(self):
        with tempfile.TemporaryDirectory() as folder:
            inv = pd.concat([invoice_frame(), invoice_frame()], ignore_index=True)
            proc, report, diagnostic = self.run_cli(folder, invoice=inv)
            self.assertEqual(proc.returncode, 2)
            self.assertEqual(load_workbook(diagnostic)["Input Issues"].max_row - 1, 2)
            self.assertFalse(report.exists())

    def test_measurable_unmatched_receipts_do_not_block(self):
        for reference in ["UNKNOWN", ""]:
            with self.subTest(reference=reference), tempfile.TemporaryDirectory() as folder:
                proc, report, diagnostic = self.run_cli(folder, payment=payment_frame("25", reference))
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertTrue(report.exists())
                self.assertFalse(diagnostic.exists())
                self.assertEqual(load_workbook(report)["Receipts Without Invoices"].max_row - 1, 1)

    def test_custom_output_has_separate_diagnostic_path(self):
        self.assertEqual(recon.diagnostic_output_path("out/custom.xlsx"), Path("out/custom_diagnostics.xlsx"))
        with tempfile.TemporaryDirectory() as folder:
            proc, report, diagnostic = self.run_cli(folder, payment=payment_frame("bad"), name="custom.xlsx")
            self.assertEqual(proc.returncode, 2)
            self.assertNotEqual(report, diagnostic)
            self.assertTrue(diagnostic.exists())
            self.assertFalse(report.exists())

    def test_diagnostic_write_failure_still_returns_failure(self):
        with tempfile.TemporaryDirectory() as folder:
            ip, pp = self.make_inputs(folder, payment=payment_frame("bad"))
            out = Path(folder) / "reconciliation_report.xlsx"
            out.write_bytes(b"KEEP")
            argv = ["reconcile_payments.py", "--invoices", str(ip), "--payments", str(pp), "-o", str(out)]
            with patch.object(sys, "argv", argv), patch.object(recon, "generate_diagnostic_report", side_effect=PermissionError("test locked workbook")), contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(recon.main(), 2)
            self.assertIn("Could not write diagnostic workbook", output.getvalue())
            self.assertEqual(out.read_bytes(), b"KEEP")

    def test_successful_rerun_identifies_old_diagnostics(self):
        with tempfile.TemporaryDirectory() as folder:
            proc, report, diagnostic = self.run_cli(folder, payment=payment_frame("bad"))
            previous = diagnostic.read_bytes()
            proc, report, diagnostic = self.run_cli(folder)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("not the status of this successful run", proc.stdout)
            self.assertTrue(report.exists())
            self.assertEqual(diagnostic.read_bytes(), previous)


if __name__ == "__main__":
    unittest.main()
