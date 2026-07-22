---
name: reconcile-subcontractor-invoices
description: Reconcile subcontractor or vendor invoice lines to customer invoices in QuickBooks, compare job costs with client service charges and full invoice totals, identify duplicate, ambiguous, or missing billing, and create an auditable Excel report. Use when a user provides a contractor invoice, vendor bill, job-cost statement, title-search invoice, or similar PDF/image and asks what clients were charged, whether jobs were billed, or what spread/margin was earned.
---

# Reconcile Subcontractor Invoices

Turn a vendor invoice plus QuickBooks customer-invoice data into a job-level reconciliation. Keep ambiguous matches visible and treat the workbook as a read-only accounting review.

## Required workflows

- Use the PDF skill to render and visually inspect PDF invoices.
- Use the Intuit QuickBooks connector for company information and customer invoices. If it is unavailable, request a QuickBooks invoice export; never invent billing data.
- Use the Spreadsheets skill for the final `.xlsx` file and follow its artifact-tool and verification requirements.
- Read [references/reconciliation-rules.md](references/reconciliation-rules.md) before matching or preparing builder input.

## Workflow

### 1. Capture the subcontractor invoice

1. Make the file locally readable. For an inaccessible mapped drive, try the exact path with the user's approved desktop permissions; otherwise request a direct upload.
2. Render every PDF page and inspect the image. Do not rely only on extracted text or OCR.
3. Capture every line as `searchDate`, `vendorType`, `vendorReference`, and `vendorCost`.
4. Reconcile the captured line count and cost sum to the printed invoice total before querying QuickBooks.
5. Record date-label inconsistencies, handwritten entries, blank file numbers, and uncertain characters.

### 2. Pull QuickBooks invoices

1. Call QuickBooks company information once to establish the company and connection.
2. Query customer invoices from the earliest vendor search date through the later of the vendor invoice date or the latest listed search date. Include a reasonable billing lag when the user has not specified one.
3. Request line items, custom file numbers, invoice numbers, customer, transaction date, total, balance, and invoice link.
4. Split the period into date chunks of seven days or fewer with a limit of 1,000. If any chunk returns exactly 1,000 records, split it again; do not assume the result is complete.
5. Keep QuickBooks read-only unless the user explicitly asks for an accounting change.

### 3. Match jobs conservatively

Apply the matching hierarchy in the reference file. Normalize identifiers for comparison, but preserve original values in the report.

- Use `Exact` for one unambiguous normalized file-number or invoice-number match.
- Use `Near-exact` only when the core identifier, date, client, and job description support a small suffix or punctuation variation.
- Use `Probable` for a strong date-plus-property/job-description match when the vendor reference is not a QuickBooks file number.
- Use `Review` for duplicates or conflicting candidates.
- Use `Unmatched` when evidence is insufficient. Never force a match to make totals reconcile.

Treat multiple active QuickBooks invoices for the same file as a duplicate-billing exception. Keep every candidate and its balance.

### 4. Calculate two comparisons

For each matched job, report:

- Vendor cost.
- Client service charge: the substantive search/service line, excluding copies, court, cover/run-down, card fees, and similar additions unless the vendor invoice specifically covers them.
- Client invoice total.
- Service spread and margin.
- Full-invoice spread and margin.

Use formulas defined in the reference file. Label results as gross spread versus the vendor, not net profit, because internal labor and other costs are not included.

For headline/base totals:

- Include all vendor costs, including unallocated lines, for a conservative spread.
- Exclude suspected duplicate customer invoices and show them separately.
- State any alternative total if the duplicate remains active in QuickBooks.

### 5. Build the workbook

1. Prepare JSON using [references/reconciliation-rules.md](references/reconciliation-rules.md).
2. Copy `scripts/build_reconciliation.mjs` into a writable work directory.
3. Call the workspace-dependency loader, create a `node_modules` junction in that work directory pointing to its bundled Node modules, and run:

   `node build_reconciliation.mjs input.json output.xlsx preview-dir`

4. Inspect headline formulas and totals, scan for formula errors, and visually review every rendered sheet.
5. Fix material clipping, unreadable text, blank sheets, broken formulas, or incorrect totals before delivery.
6. Put the final workbook under the task's `outputs/` directory. Do not deliver previews or temporary JSON.

## Required exceptions

Always surface:

- Duplicate active customer invoices.
- Unmatched or blank vendor references.
- Near-exact and probable matches.
- Vendor lines with no corresponding client billing.
- Client invoices with zero or missing service charges.
- Source totals, record counts, query range, and retrieval date.

Do not modify, void, credit, send, or create QuickBooks transactions without a separate explicit user request.

## Resources

- `references/reconciliation-rules.md`: matching hierarchy, calculations, and builder input schema.
- `scripts/build_reconciliation.mjs`: deterministic Excel reconciliation builder.
