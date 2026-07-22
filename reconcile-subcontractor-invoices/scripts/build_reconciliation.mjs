import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [inputPath, outputPath, previewDir] = process.argv.slice(2);
if (!inputPath || !outputPath || !previewDir) {
  throw new Error("Usage: node build_reconciliation.mjs input.json output.xlsx preview-dir");
}

const input = JSON.parse(await fs.readFile(inputPath, "utf8"));
const jobs = input.jobs || [];
if (!jobs.length) throw new Error("Input must contain at least one job.");

const allowedStatuses = new Set(["Exact", "Near-exact", "Probable", "Review", "Unmatched"]);
for (const [index, job] of jobs.entries()) {
  if (!Number.isFinite(job.vendorCost)) throw new Error("jobs[" + index + "].vendorCost must be numeric.");
  if (!allowedStatuses.has(job.matchStatus)) throw new Error("Invalid matchStatus at jobs[" + index + "].");
}
const vendorTotal = jobs.reduce((sum, job) => sum + Number(job.vendorCost || 0), 0);
if (!input.vendorInvoice || !Number.isFinite(input.vendorInvoice.total)) {
  throw new Error("vendorInvoice.total must be numeric.");
}
if (Math.abs(vendorTotal - input.vendorInvoice.total) > 0.01) {
  throw new Error("Vendor job costs do not reconcile to vendorInvoice.total.");
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Summary");
const recon = workbook.worksheets.add("Reconciliation");
const sources = workbook.worksheets.add("Sources");

const colors = {
  navy: "#17324D",
  teal: "#0F766E",
  white: "#FFFFFF",
  dark: "#374151",
  border: "#D1D5DB",
  lightBorder: "#E5E7EB",
  blue: "#EAF2F8",
  green: "#E8F5E9",
  tealPale: "#DFF3EF",
  amber: "#FFF4D6",
  red: "#FDE8E7",
  gray: "#F3F4F6",
};
const moneyFmt = '$#,##0.00;[Red]($#,##0.00);-';
const pctFmt = '0.0%;[Red](0.0%);-';
const startRow = 6;
const endRow = startRow + jobs.length - 1;
const totalsRow = endRow + 2;
const exceptions = (input.exceptions && input.exceptions.length) ? input.exceptions : [{
  priority: "None",
  issue: "No exceptions supplied",
  vendorReference: "",
  qbReference: "",
  amount: null,
  status: "Clear",
  recommendedAction: "No action required.",
  sourceRef: "",
}];
const duplicateTotal = jobs.reduce((sum, job) => sum + Number(job.potentialDuplicate || 0), 0);

function setColumnWidths(sheet, rowCount, widths) {
  for (let i = 0; i < widths.length; i++) {
    sheet.getRangeByIndexes(0, i, rowCount, 1).format.columnWidth = widths[i];
  }
}
function isoDate(value) {
  return value ? new Date(value + "T12:00:00") : null;
}

// Summary
summary.showGridLines = false;
summary.getRange("A1:H2").merge();
summary.getRange("A1:H2").values = [[input.reportTitle || "Subcontractor Invoice Reconciliation"]];
summary.getRange("A1:H2").format = {
  fill: colors.navy,
  font: { bold: true, color: colors.white, size: 20, name: "Arial" },
  verticalAlignment: "center",
};
summary.getRange("A3:H3").merge();
summary.getRange("A3:H3").values = [[
  (input.companyName || "QuickBooks company") + " | " +
  (input.vendorName || "Vendor") + " | Vendor invoice " +
  (input.vendorInvoice.invoiceDate || "date not supplied") + " | QuickBooks through " +
  ((input.quickbooks && input.quickbooks.queryEnd) || "date not supplied")
]];
summary.getRange("A3:H3").format = {
  fill: colors.blue,
  font: { color: colors.dark, italic: true, size: 10, name: "Arial" },
  verticalAlignment: "center",
};

const cards = [
  ["A5:B5", "A6:B7", "Vendor invoice total", "=SUM('Reconciliation'!$D$" + startRow + ":$D$" + endRow + ")", colors.blue, moneyFmt],
  ["C5:D5", "C6:D7", "Client service charges", "=SUM('Reconciliation'!$L$" + startRow + ":$L$" + endRow + ")", colors.tealPale, moneyFmt],
  ["E5:F5", "E6:F7", "Client invoice totals", "=SUM('Reconciliation'!$M$" + startRow + ":$M$" + endRow + ")", colors.green, moneyFmt],
  ["G5:H5", "G6:H7", "Matched lines (of " + jobs.length + ")", "=COUNT('Reconciliation'!$M$" + startRow + ":$M$" + endRow + ")", colors.gray, "0"],
  ["A9:B9", "A10:B11", "Service spread", "=C6-A6", colors.tealPale, moneyFmt],
  ["C9:D9", "C10:D11", "Service spread %", "=IFERROR(A10/C6,0)", colors.tealPale, pctFmt],
  ["E9:F9", "E10:F11", "Total-invoice spread", "=E6-A6", colors.green, moneyFmt],
  ["G9:H9", "G10:H11", "Total-invoice spread %", "=IFERROR(E10/E6,0)", colors.green, pctFmt],
];
for (const [labelRange, valueRange, label, formula, fill, numberFormat] of cards) {
  summary.getRange(labelRange).merge();
  summary.getRange(labelRange).values = [[label]];
  summary.getRange(labelRange).format = {
    fill,
    font: { bold: true, color: colors.dark, size: 10, name: "Arial" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: colors.border },
  };
  summary.getRange(valueRange).merge();
  summary.getRange(valueRange).formulas = [[formula]];
  summary.getRange(valueRange).format = {
    fill,
    font: { bold: true, color: colors.navy, size: 18, name: "Arial" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    numberFormat,
    borders: { preset: "outside", style: "thin", color: colors.border },
  };
}

summary.getRange("A13:H14").merge();
summary.getRange("A13:H14").values = [[
  duplicateTotal > 0
    ? "REVIEW: Potential duplicate active customer invoices total $" + duplicateTotal.toFixed(2) + " and are excluded from headline client totals."
    : "No potential duplicate customer invoices were supplied. Review all Near-exact, Probable, Review, and Unmatched rows before posting."
]];
summary.getRange("A13:H14").format = {
  fill: duplicateTotal > 0 ? colors.red : colors.amber,
  font: { bold: true, color: duplicateTotal > 0 ? "#9B1C1C" : "#92400E", size: 10, name: "Arial" },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "outside", style: "medium", color: duplicateTotal > 0 ? "#E57373" : "#D6A64F" },
};

summary.getRange("A16:H16").merge();
summary.getRange("A16:H16").values = [["Exceptions and judgment calls"]];
summary.getRange("A16:H16").format = {
  fill: colors.navy,
  font: { bold: true, color: colors.white, size: 12, name: "Arial" },
};
summary.getRange("A17:H17").values = [[
  "Priority", "Issue", "Vendor reference", "QBO reference", "Amount", "Status", "Recommended action", "Source reference"
]];
summary.getRange("A17:H17").format = {
  fill: colors.teal,
  font: { bold: true, color: colors.white, name: "Arial" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
const exceptionStart = 18;
const exceptionEnd = exceptionStart + exceptions.length - 1;
summary.getRange("A" + exceptionStart + ":H" + exceptionEnd).values = exceptions.map((x) => [
  x.priority || "", x.issue || "", x.vendorReference || "", x.qbReference || "",
  x.amount == null ? null : Number(x.amount), x.status || "", x.recommendedAction || "", x.sourceRef || ""
]);
summary.getRange("A" + exceptionStart + ":H" + exceptionEnd).format = {
  font: { name: "Arial", size: 9 },
  wrapText: true,
  verticalAlignment: "top",
  borders: {
    insideHorizontal: { style: "thin", color: colors.border },
    insideVertical: { style: "thin", color: colors.lightBorder },
    bottom: { style: "thin", color: colors.border },
  },
};
summary.getRange("E" + exceptionStart + ":E" + exceptionEnd).format.numberFormat = moneyFmt;
summary.getRange("E" + exceptionStart + ":F" + exceptionEnd).format.horizontalAlignment = "center";

const noteRow = exceptionEnd + 2;
summary.getRange("A" + noteRow + ":H" + (noteRow + 2)).merge();
summary.getRange("A" + noteRow + ":H" + (noteRow + 2)).values = [[
  "Client service charges use the substantive QBO service line. Client invoice totals can also include copies, court, cover/run-down, card fees, and other additions. Spreads compare those amounts only with the vendor invoice and are not net profit after internal labor or other costs. Headline spreads include unmatched vendor costs."
]];
summary.getRange("A" + noteRow + ":H" + (noteRow + 2)).format = {
  fill: colors.gray,
  font: { color: colors.dark, size: 10, name: "Arial" },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: colors.border },
};
setColumnWidths(summary, noteRow + 2, [13, 22, 20, 22, 15, 18, 31, 34]);
summary.getRange("A1:H2").format.rowHeight = 34;
summary.getRange("A3:H3").format.rowHeight = 24;
summary.getRange("A6:H7").format.rowHeight = 28;
summary.getRange("A10:H11").format.rowHeight = 28;
summary.getRange("A13:H14").format.rowHeight = 30;
summary.getRange("A17:H17").format.rowHeight = 34;
summary.getRange("A" + exceptionStart + ":H" + exceptionEnd).format.rowHeight = 42;
summary.getRange("A" + noteRow + ":H" + (noteRow + 2)).format.rowHeight = 28;
summary.freezePanes.freezeRows(3);

// Reconciliation
recon.showGridLines = false;
recon.getRange("A1:U2").merge();
recon.getRange("A1:U2").values = [["Job-by-Job Reconciliation"]];
recon.getRange("A1:U2").format = {
  fill: colors.navy,
  font: { bold: true, color: colors.white, size: 18, name: "Arial" },
  verticalAlignment: "center",
};
recon.getRange("A3:U3").merge();
recon.getRange("A3:U3").values = [[
  "Base totals exclude potential duplicate invoices. Keep ambiguous matches visible and edit source inputs only after verification."
]];
recon.getRange("A3:U3").format = {
  fill: colors.blue,
  font: { italic: true, color: colors.dark, size: 10, name: "Arial" },
  verticalAlignment: "center",
};
recon.getRange("A5:U5").values = [[
  "Search Date", "Vendor Reference", "Type", "Vendor Cost", "Match",
  "Client", "Job Description", "QBO Invoice", "QBO File", "QBO Date",
  "Service", "Client Service Charge", "Client Invoice Total",
  "Spread vs Service", "Service Margin", "Spread vs Invoice", "Invoice Margin",
  "QBO Balance", "Potential Duplicate", "Notes", "QuickBooks URL"
]];
recon.getRange("A5:U5").format = {
  fill: colors.teal,
  font: { bold: true, color: colors.white, size: 9, name: "Arial" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};

recon.getRange("A" + startRow + ":U" + endRow).values = jobs.map((j) => [
  isoDate(j.searchDate), j.vendorReference || "", j.vendorType || "", Number(j.vendorCost),
  j.matchStatus, j.client || "", j.jobDescription || "", j.qbInvoice || "", j.qbFile || "",
  isoDate(j.qbDate), j.service || "", j.serviceCharge == null ? null : Number(j.serviceCharge),
  j.invoiceTotal == null ? null : Number(j.invoiceTotal), null, null, null, null,
  j.balance == null ? null : Number(j.balance), j.potentialDuplicate ? Number(j.potentialDuplicate) : null,
  j.notes || "", j.qbUrl || ""
]);
recon.getRange("N" + startRow).formulas = [['=IF(L' + startRow + '="","",L' + startRow + '-D' + startRow + ')']];
recon.getRange("N" + startRow + ":N" + endRow).fillDown();
recon.getRange("O" + startRow).formulas = [['=IFERROR(N' + startRow + '/L' + startRow + ',"")']];
recon.getRange("O" + startRow + ":O" + endRow).fillDown();
recon.getRange("P" + startRow).formulas = [['=IF(M' + startRow + '="","",M' + startRow + '-D' + startRow + ')']];
recon.getRange("P" + startRow + ":P" + endRow).fillDown();
recon.getRange("Q" + startRow).formulas = [['=IFERROR(P' + startRow + '/M' + startRow + ',"")']];
recon.getRange("Q" + startRow + ":Q" + endRow).fillDown();

recon.getRange("A" + totalsRow + ":C" + totalsRow).merge();
recon.getRange("A" + totalsRow + ":C" + totalsRow).values = [["Totals / base scenario"]];
recon.getRange("D" + totalsRow).formulas = [["=SUM(D" + startRow + ":D" + endRow + ")"]];
recon.getRange("L" + totalsRow).formulas = [["=SUM(L" + startRow + ":L" + endRow + ")"]];
recon.getRange("M" + totalsRow).formulas = [["=SUM(M" + startRow + ":M" + endRow + ")"]];
recon.getRange("N" + totalsRow).formulas = [["=L" + totalsRow + "-D" + totalsRow]];
recon.getRange("O" + totalsRow).formulas = [["=IFERROR(N" + totalsRow + "/L" + totalsRow + ",0)"]];
recon.getRange("P" + totalsRow).formulas = [["=M" + totalsRow + "-D" + totalsRow]];
recon.getRange("Q" + totalsRow).formulas = [["=IFERROR(P" + totalsRow + "/M" + totalsRow + ",0)"]];
recon.getRange("R" + totalsRow).formulas = [["=SUM(R" + startRow + ":R" + endRow + ")"]];
recon.getRange("S" + totalsRow).formulas = [["=SUM(S" + startRow + ":S" + endRow + ")"]];
recon.getRange("A" + totalsRow + ":U" + totalsRow).format = {
  fill: colors.blue,
  font: { bold: true, color: colors.navy, name: "Arial" },
  borders: { top: { style: "medium", color: colors.navy }, bottom: { style: "double", color: colors.navy } },
};

recon.getRange("A" + startRow + ":U" + endRow).format = {
  font: { name: "Arial", size: 9 },
  verticalAlignment: "top",
  borders: {
    insideHorizontal: { style: "thin", color: colors.lightBorder },
    insideVertical: { style: "thin", color: "#EDF0F2" },
  },
};
recon.getRange("A" + startRow + ":A" + endRow).format.numberFormat = "yyyy-mm-dd";
recon.getRange("J" + startRow + ":J" + endRow).format.numberFormat = "yyyy-mm-dd";
recon.getRange("D" + startRow + ":D" + totalsRow).format.numberFormat = moneyFmt;
recon.getRange("L" + startRow + ":N" + totalsRow).format.numberFormat = moneyFmt;
recon.getRange("O" + startRow + ":O" + totalsRow).format.numberFormat = pctFmt;
recon.getRange("P" + startRow + ":P" + totalsRow).format.numberFormat = moneyFmt;
recon.getRange("Q" + startRow + ":Q" + totalsRow).format.numberFormat = pctFmt;
recon.getRange("R" + startRow + ":S" + totalsRow).format.numberFormat = moneyFmt;
recon.getRange("D" + startRow + ":D" + endRow).format.fill = colors.blue;
recon.getRange("L" + startRow + ":M" + endRow).format.fill = colors.green;
recon.getRange("N" + startRow + ":Q" + endRow).format.fill = colors.tealPale;
recon.getRange("S" + startRow + ":S" + endRow).format.fill = colors.red;
recon.getRange("T" + startRow + ":T" + endRow).format.wrapText = true;
recon.getRange("U" + startRow + ":U" + endRow).format.font = { color: "#2563EB", size: 8, name: "Arial" };

for (const status of ["Review", "Unmatched"]) {
  recon.getRange("E" + startRow + ":E" + endRow).conditionalFormats.add("containsText", {
    text: status, format: { fill: colors.red, font: { bold: true, color: "#9B1C1C" } },
  });
}
for (const status of ["Probable", "Near-exact"]) {
  recon.getRange("E" + startRow + ":E" + endRow).conditionalFormats.add("containsText", {
    text: status, format: { fill: colors.amber, font: { bold: true, color: "#92400E" } },
  });
}
const table = recon.tables.add("A5:U" + endRow, true, "SubcontractorReconciliationTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
setColumnWidths(recon, totalsRow, [12, 21, 10, 12, 12, 27, 31, 17, 22, 12, 25, 15, 15, 15, 12, 15, 12, 14, 15, 46, 22]);
recon.getRange("A5:U5").format.rowHeight = 42;
recon.getRange("A" + startRow + ":U" + endRow).format.rowHeight = 22;
for (let i = 0; i < jobs.length; i++) {
  if (jobs[i].notes) recon.getRangeByIndexes(startRow - 1 + i, 0, 1, 21).format.rowHeight = 42;
}
recon.freezePanes.freezeRows(5);
recon.freezePanes.freezeColumns(3);

// Sources
sources.showGridLines = false;
sources.getRange("A1:G2").merge();
sources.getRange("A1:G2").values = [["Sources, Scope, and Matching Method"]];
sources.getRange("A1:G2").format = {
  fill: colors.navy,
  font: { bold: true, color: colors.white, size: 18, name: "Arial" },
  verticalAlignment: "center",
};
sources.getRange("A4:G4").values = [[
  "Source Type", "Source / Company", "Period / As-of", "Records", "Amount", "Reference", "Notes"
]];
sources.getRange("A4:G4").format = {
  fill: colors.teal,
  font: { bold: true, color: colors.white, name: "Arial" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
const sourceNotes = (input.sourceNotes && input.sourceNotes.length) ? input.sourceNotes : [
  {
    sourceType: "Subcontractor invoice",
    source: input.vendorName || "Vendor",
    period: input.vendorInvoice.periodLabel || input.vendorInvoice.invoiceDate || "",
    records: jobs.length,
    amount: input.vendorInvoice.total,
    reference: input.vendorInvoice.sourcePath || "",
    notes: "Rendered and visually reviewed before matching.",
  },
  {
    sourceType: "QuickBooks Online",
    source: input.companyName || "",
    period: input.quickbooks ? input.quickbooks.queryStart + " through " + input.quickbooks.queryEnd : "",
    records: input.quickbooks ? input.quickbooks.invoiceCount : null,
    amount: null,
    reference: "Customer invoices and line items",
    notes: input.quickbooks ? "Retrieved " + input.quickbooks.retrievedDate + "." : "",
  },
  {
    sourceType: "Matching method",
    source: "Job-level reconciliation",
    period: input.vendorInvoice.periodLabel || "",
    records: jobs.filter((j) => j.invoiceTotal != null).length,
    amount: duplicateTotal || null,
    reference: "File number, invoice number, date, and description",
    notes: "Potential duplicates are excluded from the base totals.",
  },
];
const sourceStart = 5;
const sourceEnd = sourceStart + sourceNotes.length - 1;
sources.getRange("A" + sourceStart + ":G" + sourceEnd).values = sourceNotes.map((s) => [
  s.sourceType || "", s.source || "", s.period || "", s.records == null ? null : Number(s.records),
  s.amount == null ? null : Number(s.amount), s.reference || "", s.notes || ""
]);
sources.getRange("A" + sourceStart + ":G" + sourceEnd).format = {
  font: { name: "Arial", size: 10 },
  wrapText: true,
  verticalAlignment: "top",
  borders: { insideHorizontal: { style: "thin", color: colors.border }, bottom: { style: "thin", color: colors.border } },
};
sources.getRange("D" + sourceStart + ":D" + sourceEnd).format.numberFormat = "#,##0";
sources.getRange("E" + sourceStart + ":E" + sourceEnd).format.numberFormat = moneyFmt;
setColumnWidths(sources, sourceEnd, [20, 34, 28, 12, 14, 48, 62]);
sources.getRange("A4:G4").format.rowHeight = 34;
sources.getRange("A" + sourceStart + ":G" + sourceEnd).format.rowHeight = 54;
sources.freezePanes.freezeRows(4);

// Verification and previews
const summaryCheck = await workbook.inspect({
  kind: "table",
  range: "Summary!A5:H14",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 8,
  maxChars: 6000,
});
const totalCheck = await workbook.inspect({
  kind: "table",
  range: "Reconciliation!A" + totalsRow + ":S" + totalsRow,
  include: "values,formulas",
  tableMaxRows: 2,
  tableMaxCols: 19,
  maxChars: 4000,
});
const errorCheck = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 4000,
});
console.log(summaryCheck.ndjson);
console.log(totalCheck.ndjson);
console.log(errorCheck.ndjson);

for (const sheetName of ["Summary", "Reconciliation", "Sources"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: sheetName === "Reconciliation" ? 0.8 : 1.2,
    format: "png",
  });
  await fs.writeFile(path.join(previewDir, sheetName.toLowerCase() + ".png"), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
await fs.rm(outputPath + ".inspect.ndjson", { force: true });

console.log(JSON.stringify({
  outputPath,
  vendorTotal,
  jobCount: jobs.length,
  matchedCount: jobs.filter((j) => j.invoiceTotal != null).length,
  duplicateTotal,
}, null, 2));
