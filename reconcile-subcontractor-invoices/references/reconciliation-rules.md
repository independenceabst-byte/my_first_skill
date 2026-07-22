# Reconciliation Rules and Builder Input

## Matching hierarchy

Compare identifiers after uppercasing and removing all non-alphanumeric characters. Preserve original identifiers in outputs.

1. **Exact file match**: normalized vendor reference equals the QuickBooks custom file number.
2. **Exact invoice match**: normalized vendor reference equals the customer invoice number and the date/job context is plausible.
3. **Near-exact identifier**: small suffix, punctuation, or OCR difference with the same core number plus supporting date, client, and description.
4. **Description match**: distinctive property/address/party description plus a close date. Do not use a date alone.
5. **Unmatched**: no defensible match.

If more than one active QuickBooks invoice matches the same vendor job, use `Review`, retain both candidates, and exclude the suspected duplicate from the base total.

## Calculations

- `serviceSpread = clientServiceCharge - vendorCost`
- `serviceMargin = serviceSpread / clientServiceCharge`
- `invoiceSpread = clientInvoiceTotal - vendorCost`
- `invoiceMargin = invoiceSpread / clientInvoiceTotal`

Return blank margins when the denominator is zero or missing.

Headline spreads subtract the full vendor invoice total, including unallocated vendor costs. Therefore headline spreads can differ from the sum of matched row spreads.

## Service-line selection

Use the primary search/service line that corresponds to the vendor work, such as `FULL`, `DEVELOPER`, `PRESENT OWNER`, `40 YEAR`, `CONTINUATION`, or a comparable core service.

Do not include additions such as copies, surrogate/court fees, cover or run-down, card fees, recording, shipping, or taxes unless the subcontractor line clearly covers that addition. Preserve both the service charge and the complete invoice total in the report.

## JSON input schema

All amounts are numbers, not currency strings. Use ISO dates (`YYYY-MM-DD`). Use `null` for unknown numeric values and an empty string for unknown text.

```json
{
  "reportTitle": "Subcontractor Invoice Reconciliation",
  "companyName": "Company name from QuickBooks",
  "vendorName": "Vendor or subcontractor",
  "vendorInvoice": {
    "sourcePath": "original file name or path",
    "invoiceDate": "2026-07-20",
    "periodLabel": "2026-06-08 through 2026-07-20",
    "total": 2956.50
  },
  "quickbooks": {
    "queryStart": "2026-06-08",
    "queryEnd": "2026-07-21",
    "retrievedDate": "2026-07-21",
    "invoiceCount": 2355
  },
  "jobs": [
    {
      "searchDate": "2026-06-10",
      "vendorType": "Full",
      "vendorReference": "ABC-123",
      "vendorCost": 75,
      "matchStatus": "Exact",
      "qbDate": "2026-06-10",
      "qbInvoice": "123",
      "qbFile": "ABC-123",
      "client": "Client Name",
      "jobDescription": "OWNER / MUNICIPALITY",
      "service": "FULL",
      "serviceCharge": 115,
      "invoiceTotal": 143.90,
      "balance": 0,
      "potentialDuplicate": 0,
      "notes": "",
      "qbUrl": "https://qbo.intuit.com/..."
    }
  ],
  "exceptions": [
    {
      "priority": "High",
      "issue": "Potential duplicate customer billing",
      "vendorReference": "ABC-123",
      "qbReference": "ABC-123",
      "amount": 143.90,
      "status": "Both invoices active",
      "recommendedAction": "Confirm intended invoice before voiding or crediting.",
      "sourceRef": "QBO transaction ID or concise reference"
    }
  ],
  "sourceNotes": [
    {
      "sourceType": "Subcontractor invoice",
      "source": "Vendor name",
      "period": "Invoice and search period",
      "records": 39,
      "amount": 2956.50,
      "reference": "file name",
      "notes": "Extraction or date-label notes"
    }
  ]
}
```

Allowed `matchStatus` values: `Exact`, `Near-exact`, `Probable`, `Review`, and `Unmatched`.

The sum of `jobs[].vendorCost` must equal `vendorInvoice.total` within one cent. Add unmatched vendor lines as jobs rather than dropping them.
