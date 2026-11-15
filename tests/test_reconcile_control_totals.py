"""Save in tests/. Run: python -m unittest discover -s tests -p 'test_reconcile_*.py' -v"""
import contextlib
import importlib.util
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from openpyxl import load_workbook

root = Path(__file__).resolve().parent
script = root.parent / 'src' / 'reconcile_payments.py'
if not script.exists():
    script = root / 'reconcile_payments.py'
spec = importlib.util.spec_from_file_location('controls_under_test', script)
recon = importlib.util.module_from_spec(spec)
spec.loader.exec_module(recon)


def invoices(amounts=('100',), statuses=None):
    return pd.DataFrame([dict(invoice_id=f'I{i}', customer_name=f'Customer {i}',
                             amount_usd=a, days_overdue='40', status=(statuses or ['Open'] * len(amounts))[i])
                         for i, a in enumerate(amounts)])


def receipts(amounts=('40',), refs=None):
    return pd.DataFrame([dict(payment_id=f'P{i}', invoice_id=(refs or ['I0'] * len(amounts))[i],
                             amount_paid=a, payment_date='2026-01-01') for i, a in enumerate(amounts)],
                        columns=['payment_id', 'invoice_id', 'amount_paid', 'payment_date'])


class ReceiptControlTests(unittest.TestCase):
    def checked(self, inv, pay):
        result = recon.reconcile_payments(inv, pay)
        checks = recon.build_reconciliation_checks(result)
        self.assertTrue(checks.Status.eq('PASS').all(), checks.to_string())
        return result

    def summary_workbook(self, inv, pay):
        result = self.checked(inv, pay)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'report.xlsx'
            recon.generate_reconciliation_report(result, path)
            return load_workbook(path, data_only=True)

    def test_summary_bridge_uses_signed_paid_differences_not_only_exceptions(self):
        wb = self.summary_workbook(invoices(('100', '100', '100'), ['Paid', 'Paid', 'Open']),
                                   receipts(('40', '120', '30'), ['I0', 'I1', 'I2']))
        m = {row[0]: row[1:3] for row in wb['Summary'].values if row[0]}
        self.assertEqual(m['Calculated signed invoice difference—all statuses'], (3, 110))
        self.assertEqual(m['Less: signed differences on invoices marked Paid'], (2, -40))
        self.assertEqual(m['Non-Paid balance—comparable to AR aging'], (1, 70))
        self.assertEqual(m['Paid-status positive residuals—review subset'], (1, 60))

    def test_summary_bridge_paid_excess_adjustment_is_positive(self):
        wb = self.summary_workbook(invoices(('100', '100'), ['Paid', 'Open']), receipts(('120',)))
        m = {row[0]: row[1:3] for row in wb['Summary'].values if row[0]}
        self.assertEqual(m['Calculated signed invoice difference—all statuses'][1], 80)
        self.assertEqual(m['Less: signed differences on invoices marked Paid'][1], 20)
        self.assertEqual(m['Non-Paid balance—comparable to AR aging'][1], 100)

    def test_summary_coverage_shows_original_and_remaining_with_complete_totals(self):
        wb = self.summary_workbook(invoices(('100', '100', '100', '100')),
                                   receipts(('100', '40', '120'), ['I0', 'I1', 'I3']))
        m = {row[0]: row[1:] for row in wb['Summary'].values if row[0]}
        self.assertEqual(m['4. Invoice coverage categories (all source statuses)'],
                         ('Invoice count', 'Original invoice amount', 'Remaining difference (signed)'))
        self.assertEqual(m['Partial receipt coverage'], (1, 100, 60))
        self.assertEqual(m['Excess receipts'], (1, 100, -20))
        self.assertEqual(m['Total—all invoice categories'], (4, 400, 140))

    def test_summary_sections_distinguish_receipts_balances_and_review(self):
        wb = self.summary_workbook(invoices(), receipts())
        labels = [row[0] for row in wb['Summary'].values if row[0]]
        for section in ['1. Connection to AR aging', '2. Receipt totals (receipt records, not invoice balances)',
                        '3. Invoice amounts and balance components', '5. Paid-status exceptions (review subset only)']:
            self.assertIn(section, labels)
        text = ' '.join(str(x) for x in labels)
        self.assertIn('does not open or independently check an aging workbook', text)
        self.assertIn('do not add them all together', text)

    def test_summary_counts_fully_covered_paid_invoices_in_bridge(self):
        wb = self.summary_workbook(invoices(('100', '100'), ['Paid', 'Open']),
                                   receipts(('100', '40'), ['I0', 'I1']))
        m = {row[0]: row[1:3] for row in wb['Summary'].values if row[0]}
        self.assertEqual(m['Less: signed differences on invoices marked Paid'], (1, 0))
        self.assertEqual(m['Non-Paid balance—comparable to AR aging'], (1, 60))

    def test_additional_source_columns_preserved_in_workbook(self):
        inv, pay = invoices(), receipts(('120',))
        inv['due_date'] = '2026-01-01'
        inv['po_number'] = '00001'
        result = self.checked(inv, pay)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'report.xlsx'
            recon.generate_reconciliation_report(result, path)
            wb = load_workbook(path)
            rows = list(wb['Excess Receipts'].values)
            row = dict(zip(rows[0], rows[1]))
            self.assertEqual(row['Due date'], '2026-01-01')
            self.assertEqual(row['PO number'], '00001')

    def test_cent_precision_loss_input_is_blocked(self):
        with self.assertRaises(recon.InputValidationError):
            recon.reconcile_payments(invoices(('100000000000000.01',)), receipts())

    def test_overpayment_keeps_full_receipt_and_separate_positive_excess(self):
        r = self.checked(invoices(), receipts(('120',)))
        row = r['all_invoices'].iloc[0]
        self.assertEqual(tuple(row[c] for c in ['total_paid', 'balance_due', 'positive_residual_usd',
                                                'excess_receipts_usd', 'analytical_coverage_usd']),
                         (120, -20, 0, 20, 100))
        self.assertEqual(len(r['overpayments']), 1)
        self.assertEqual(r['matched_receipts'].amount_paid.sum(), 120)

    def test_one_cent_shortfall_is_not_fully_covered(self):
        r = self.checked(invoices(), receipts(('99.99',)))
        self.assertEqual(len(r['partial']), 1)
        self.assertEqual(len(r['matched']), 0)
        self.assertEqual(r['partial'].iloc[0].positive_residual_usd, .01)

    def test_one_cent_excess_is_not_discarded(self):
        r = self.checked(invoices(), receipts(('100.01',)))
        self.assertEqual(len(r['overpayments']), 1)
        self.assertEqual(r['overpayments'].iloc[0].excess_receipts_usd, .01)

    def test_exact_cents_multiple_receipts(self):
        r = self.checked(invoices(('.30',)), receipts(('.10', '.20')))
        self.assertEqual(len(r['matched']), 1)
        self.assertEqual(r['matched'].iloc[0].balance_due, 0)

    def test_paid_status_does_not_remove_financial_arithmetic(self):
        r = self.checked(invoices(('100', '200'), ['Paid', 'Open']), receipts(('40', '220'), ['I0', 'I1']))
        self.assertEqual(r['all_invoices'].balance_due.sum(), 40)
        self.assertEqual(r['all_invoices'].positive_residual_usd.sum(), 60)
        self.assertEqual(r['all_invoices'].excess_receipts_usd.sum(), 20)
        self.assertEqual(len(r['paid_status_exceptions']), 1)
        self.assertEqual(r['paid_status_exceptions'].iloc[0].positive_residual_usd, 60)

    def test_no_cross_invoice_offset(self):
        r = self.checked(invoices(('100', '100')), receipts(('120',), ['I0']))
        self.assertEqual(r['all_invoices'].positive_residual_usd.sum(), 100)
        self.assertEqual(r['all_invoices'].excess_receipts_usd.sum(), 20)
        self.assertEqual(r['all_invoices'].balance_due.sum(), 80)

    def test_no_receipts_retains_all_invoices_including_zero(self):
        r = self.checked(invoices(('0', '100')), receipts(()))
        self.assertEqual(len(r['unpaid_invoices']), 2)
        self.assertEqual(r['all_invoices'].balance_due.sum(), 100)

    def test_missing_and_unknown_refs_retained_distinctly(self):
        r = self.checked(invoices(), receipts(('10', '20', '30'), ['I0', '', 'UNKNOWN']))
        self.assertEqual(r['matched_receipts'].amount_paid.sum(), 10)
        self.assertEqual(r['unmatched_payments'].amount_paid.sum(), 50)
        self.assertEqual(set(r['unmatched_payments'].unmatched_reason),
                         {'Missing invoice reference', 'Invoice reference not found'})

    def test_repeated_receipt_ids_not_silently_deduplicated(self):
        pay = receipts(('20', '30'))
        pay['payment_id'] = 'SAME'
        r = self.checked(invoices(), pay)
        self.assertEqual(r['all_invoices'].iloc[0].payment_count, 2)
        self.assertEqual(r['all_invoices'].iloc[0].total_paid, 50)

    def test_dropped_receipt_fails_controls(self):
        r = self.checked(invoices(), receipts(('10', '20')))
        r['matched_receipts'] = r['matched_receipts'].iloc[:1]
        self.assertTrue(recon.build_reconciliation_checks(r).Status.eq('FAIL').any())

    def test_duplicate_receipt_output_fails_identity(self):
        r = self.checked(invoices(), receipts(('10',)))
        r['matched_receipts'] = pd.concat([r['matched_receipts']] * 2)
        checks = recon.build_reconciliation_checks(r).set_index('Check')
        self.assertEqual(checks.loc['Receipt record identity', 'Status'], 'FAIL')

    def test_equal_count_equal_amount_replacement_fails_identity(self):
        r = self.checked(invoices(), receipts(('10', '10')))
        r['matched_receipts'].loc[1, 'receipt_data_row'] = 1
        checks = recon.build_reconciliation_checks(r).set_index('Check')
        self.assertEqual(checks.loc['Receipt record identity', 'Status'], 'FAIL')

    def test_swapped_invoice_receipts_fail_association(self):
        r = self.checked(invoices(('100', '100')), receipts(('30', '40'), ['I0', 'I1']))
        r['all_invoices']['total_paid'] = [40, 30]
        checks = recon.build_reconciliation_checks(r).set_index('Check')
        self.assertEqual(checks.loc['Per-invoice receipt association', 'Status'], 'FAIL')

    def test_subcent_inputs_block_no_silent_rounding(self):
        for amount in ['100.001', '0.009']:
            with self.subTest(amount=amount), self.assertRaises(recon.InputValidationError):
                recon.reconcile_payments(invoices(), receipts((amount,)))

    def test_all_invoice_coverage_categories_complete_once(self):
        r = self.checked(invoices(('100', '100', '100', '100')),
                         receipts(('100', '40', '120'), ['I0', 'I1', 'I3']))
        self.assertEqual([len(r[k]) for k in ('matched', 'partial', 'unpaid_invoices', 'overpayments')], [1, 1, 1, 1])
        self.assertEqual(r['all_invoices'].amount_usd.sum(), 400)

    def test_workbook_definitions_totals_and_exact_cent_formats(self):
        r = self.checked(invoices(('100', '100'), ['Paid', 'Open']), receipts(('120', '10'), ['I0', 'UNKNOWN']))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'report.xlsx'
            recon.generate_reconciliation_report(r, path)
            wb = load_workbook(path, data_only=True)
            self.assertEqual(wb.sheetnames, ['Summary', 'Reconciliation Checks', 'Fully Covered', 'Partial Coverage',
                                           'Receipts Without Invoices', 'No Positive Receipts', 'Excess Receipts', 'Paid Status Exceptions'])
            metrics = {row[0]: row[1:3] for row in wb['Summary'].values if row[0]}
            self.assertEqual(metrics['Reconciliation status'][0], 'REVIEW')
            self.assertEqual(metrics['Total supplied receipts'], (2, 130))
            self.assertEqual(metrics['Invoice-reference-matched receipts'], (1, 120))
            self.assertEqual(metrics['Receipts Without Corresponding Invoices'], (1, 10))
            self.assertEqual(metrics['Calculated signed invoice difference—all statuses'], (2, 80))
            self.assertEqual(metrics['Positive calculated residuals—all statuses'], (1, 100))
            self.assertEqual(metrics['Excess receipts requiring review'], (1, 20))
            text = ' '.join(str(c.value) for ws in wb for row in ws for c in row)
            for old in ['Cash Applied to Subledger', 'Net Open AR Ties to Aging', 'Matched Settled in Full']:
                self.assertNotIn(old, text)
            self.assertEqual(wb['Excess Receipts']['I2'].number_format, recon.CURRENCY_FORMAT)

    def test_zero_receipt_record_count_preserved(self):
        r = self.checked(invoices(), receipts(('0',)))
        self.assertEqual(r['unpaid_invoices'].iloc[0].payment_count, 1)
        self.assertEqual(len(r['matched_receipts']), 1)

    def test_failed_controls_do_not_refresh_normal_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            ip, pp, out = tmp/'i.csv', tmp/'p.csv', tmp/'report.xlsx'
            invoices().to_csv(ip, index=False)
            receipts().to_csv(pp, index=False)
            out.write_bytes(b'previous normal report')
            original = recon.build_reconciliation_checks
            def failure(results):
                checks = original(results)
                checks.loc[0, 'Status'] = 'FAIL'
                return checks
            with patch.object(sys, 'argv', ['recon', '--invoices', str(ip), '--payments', str(pp), '-o', str(out)]), \
                 patch.object(recon, 'build_reconciliation_checks', side_effect=failure), \
                 contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(recon.main(), 3)
            self.assertEqual(out.read_bytes(), b'previous normal report')
            failed = load_workbook(tmp/'report_failed_checks.xlsx', data_only=True)
            self.assertTrue(failed['Summary']['B1'].value.startswith('FAIL'))

    def test_normal_cli_with_review_items_completes_exit_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            ip, pp, out = tmp/'i.csv', tmp/'p.csv', tmp/'report.xlsx'
            invoices().to_csv(ip, index=False)
            receipts(('120',)).to_csv(pp, index=False)
            with patch.object(sys, 'argv', ['recon', '--invoices', str(ip), '--payments', str(pp), '-o', str(out)]), \
                 contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(recon.main(), 0)
            self.assertEqual(load_workbook(out)['Summary']['B1'].value, 'REVIEW')


if __name__ == '__main__':
    unittest.main()
