# AR Automation Toolkit

Turn invoice and receipt exports into review-ready Excel reports for AR aging, collections prioritization, and invoice-to-receipt reference reconciliation.

AR teams need more than a balance report: they need to identify remaining balances, investigate unsupported Paid statuses, and trace receipts that do not correspond to invoices in an extract. This Python toolkit makes those review steps repeatable while preserving source records and checking reconciliation completeness. It complements ERP reporting rather than replacing it.

**Built with:** Python · pandas · openpyxl
**Portfolio project:** synthetic data · 67 automated reconciliation tests · 15 reconciliation control checks

## 1. Reports and sample results

| Report                                                                                  | What it helps reviewers do                                                                                                          |
| --------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| **AR Aging** — `aging_summary.xlsx`                                            | Review balances by aging bucket, prioritize positive non-Paid balances, identify risk accounts, and inspect Paid-status exceptions. |
| **Invoice-to-Receipt Reference Reconciliation** — `reconciliation_report.xlsx` | Trace receipt references, calculate coverage, investigate exceptions, and review control results.                                   |

The reports answer different questions: aging supports collections review; reconciliation explains receipt coverage and checks record and amount completeness.

### Sample balance bridge

The supplied synthetic sample illustrates why reporting populations matter:

| Measure                                           |                     USD |
| ------------------------------------------------- | ----------------------: |
| Net Unsettled Subledger Balance (All Statuses AR) |           $1,390,822.67 |
| Less: Unsubstantiated Paid-Status Balances        |            −$69,108.08 |
| Net Open AR (Active Non-Paid Subledger)           | **$1,321,714.59** |

Bridge Explanation: 

* **Exception Preservation:** Reconciles workbook scope differences without auto-clearing unsubstantiated `Paid`-status exceptions, which remain isolated for review
* **Metric Separation:** Keeps Net Open AR (all active non-`Paid`balances, including credits), Past-Due AR (past due date), and Collectible Debit Balances (positive-only collections worklists) distinct.
* **Independent Control:** Derives the tie-out directly from source extracts rather than reading the aging workbook, making cross-report agreement an independent completeness check

## 2. How it works

```text
Invoice CSV + Receipt CSV
           │
           ├── AR aging → balances, aging buckets, collections worklists
           │
           └── Reconciliation → input validation → reference matching
                                → coverage categories + 15 controls
           │
           └── Excel reports → finance review
```

Reconciliation preserves full receipt amounts and signed differences. It separates fully covered invoices, partial coverage, no positive receipts, excess receipts, and receipts without corresponding invoices. One-cent differences remain visible.

**Example:** a $10,000 invoice with $9,950 in matching receipts has $50 of remaining coverage to review.

## 3. Quickstart

Use **Python 3.11+**. Run commands from the repository root.

```bash
git clone https://github.com/ajanetc/ar-automation-toolkit.git
cd ar-automation-toolkit
```

Create and activate a virtual environment:

```bash
# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

```powershell
# Windows PowerShell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies, generate both reports, and run the reconciliation tests:

```bash
python -m pip install -r requirements.txt
python src/reconcile_payments.py
python src/ar_aging_automation.py --payments data/payment_receipts.csv
python -m unittest discover -s tests -p "test_reconcile_*.py" -v
```

Default inputs are `data/invoices_sample.csv` and `data/payment_receipts.csv`; reports are written to `output/`. The current reconciliation suite contains **67 tests**; this is not complete aging or end-to-end coverage. Fresh-clone and cross-platform setup verification remain publication tasks.

For custom paths:

```bash
python src/reconcile_payments.py --invoices data/invoices_sample.csv --payments data/payment_receipts.csv -o output/reconciliation_report.xlsx
python src/ar_aging_automation.py data/invoices_sample.csv --payments data/payment_receipts.csv -o output/aging_summary.xlsx
```

The aging invoice path is **positional**, not `--input`. Use either script's `--help` for options. Close Excel workbooks before regenerating them if your system prevents replacement.

## 4. Inputs and accounting rules

| Input                    | Required columns                                                                                                               | Row basis                                                 |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------- |
| `invoices_sample.csv`  | `invoice_id`, `customer_name`, `invoice_date`, `due_date`, `amount_usd`, `status`, `days_overdue`, `po_number` | One row per invoice, not per product line.                |
| `payment_receipts.csv` | `payment_id`, `invoice_id`, `payment_date`, `amount_paid`                                                              | One uniquely identified receipt intended for one invoice. |

- **Balance:** original invoice amount (`amount_usd`) minus supplied matching receipts. Do not subtract receipts again from an already-net ERP balance; receipt coverage must be appropriate for the original amounts.
- **Identifiers:** payment IDs must be nonblank and unique. Multiple receipts may reference one invoice. Reconciliation preserves leading zeros; nonblank identifiers are compared as supplied, without automatic whitespace/case normalization.
- **Dates:** intended format is `YYYY-MM-DD`. Default aging uses supplied `days_overdue`; receipt dates do not filter balances.
- **Collections:** positive non-Paid remaining balances only. Paid-status shortfalls and negative balances are shown separately; source statuses are not overwritten.
- **Aging buckets:** Current, 1–30, 31–60, 61–90, and 90+; `90+` means more than 90 days.
- **Risk accounts:** at least 3 overdue eligible invoices **or** $10,000 in overdue eligible balance.

Running aging without `--payments` does not evaluate receipt support.

## 5. Workbook guide

| Workbook       | Tabs                                                                                                                                                                                                                                      |
| -------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Aging          | **Summary** (including Report Basis), **All Invoices**, **Payment Receipts**, **Detail >30d**, **Risk Accounts**, **Paid Status Exceptions**, **Credit Balances**, **Validation Issues**  |
| Reconciliation | **Summary**, **Reconciliation Checks**, **Fully Covered**, **Partial Coverage**, **Receipts Without Invoices**, **No Positive Receipts**, **Excess Receipts**, **Paid Status Exceptions** |

**Detail >30d** contains positive non-Paid balances over 30 days overdue. Reconciliation coverage categories partition invoices; **Paid Status Exceptions** is an overlapping review subset, not an additional category to add to totals.

## 6. Validation, controls, and run statuses

Reconciliation blocks missing required columns, unreadable inputs, duplicate nonblank invoice IDs, blank/duplicate payment IDs, and invalid amounts—including unsupported negative, non-finite, or sub-cent values and amounts that cannot retain cent precision.

Its **15 control checks** cover source totals, record completeness, reference membership, per-invoice receipt association, and invoice-category completeness.

| Status  | Meaning                                                                                                  | Output / exit code                                  |
| ------- | -------------------------------------------------------------------------------------------------------- | --------------------------------------------------- |
| PASS    | Checks agree; no counted review exceptions remain.                                                       | Normal report; `0`                                |
| REVIEW  | Checks agree; receipts without invoices, excess receipts, or Paid-status positive residuals need review. | Normal report; `0`                                |
| BLOCKED | Critical input issues prevent reconciliation.                                                            | `reconciliation_diagnostics.xlsx`; `2`          |
| FAIL    | A financial/control check fails.                                                                         | `reconciliation_report_failed_checks.xlsx`; `3` |

These filenames assume the default output path. Diagnostics retain loaded source records and identify affected rows. A blocked or failed run **does not refresh the normal report**; check terminal messages and workbook timestamps before using existing files. Old diagnostics may also remain after a successful rerun.

The synthetic sample intentionally includes review exceptions. Passing tests or controls is not proof that source data is complete or correct, nor a measure of business savings.

## 7. Scope and limitations

- Personal, synthetic-data portfolio project—not a production accounting system, bank reconciliation, or general-ledger reconciliation.
- Reference matching and coverage analysis do not verify ERP posting. No automatic allocation, posting, write-off, refund, or credit-memo processing is performed. Unknown references may relate to invoices outside the extract.
- USD-oriented and invoice-level; split receipts, multi-currency accounting, negative receipts, and reversals require additional rules.
- `--as-of-date` recalculates overdue days only; it does not reconstruct historical AR, filter receipts, or restore historical statuses.
- Reconciliation's newer blocking and exact-cent safeguards are not all implemented in aging. Broader date validation, missing invoice-ID validation, and identifier normalization remain future work.

## 8. Project structure

```text
ar-automation-toolkit/
├── data/                  # Synthetic invoice and receipt CSVs
├── src/
│   ├── ar_aging_automation.py
│   └── reconcile_payments.py
├── tests/                 # Reconciliation tests and fixtures
├── output/                # Generated locally; Excel reports ignored by Git
├── requirements.txt
├── LICENSE
└── README.md
```

Generated Excel reports and `.DS_Store` files are excluded from Git.

## 9. Roadmap

**Next:** add selected report screenshots, verify setup from a fresh clone, expand aging/cross-report tests, and align validation safeguards across both scripts.

**Potential future extensions—not current features:** SQLite storage, ERP export/API adapters, explicit receipt-allocation rules, historical snapshots, and an MCP-compatible interface. Each requires defined accounting rules and additional testing before implementation.
