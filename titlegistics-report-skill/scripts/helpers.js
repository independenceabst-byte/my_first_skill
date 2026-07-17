/**
 * TitleGistics Report — Reusable Helpers
 * Require this file at the top of any report build script:
 *   const H = require('./titlegistics-report/scripts/helpers');
 * Then destructure what you need.
 */

const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, Footer,
} = require('docx');

// ─── Design Tokens ────────────────────────────────────────────────────────────
const NAVY    = "0A1A2E";
const GOLD    = "C9A84C";
const WHITE   = "FFFFFF";
const GRAY_BG = "F5F5F5";
const RED_BG  = "FCEBEB";
const RED_TXT = "A32D2D";
const BLUE_BG = "E6F1FB";
const BLUE_TXT= "185FA5";
const GREEN_TXT = "2A7A2A";
const AMBER_TXT = "885500";
const AMBER_BG  = "FDFBF0";

// ─── Border Presets ───────────────────────────────────────────────────────────
const noBorder   = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders  = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const thinBorders= { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const goldBorder = { style: BorderStyle.SINGLE, size: 8, color: GOLD };
const boxBorders = {
  top:    { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" },
  left:   { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" },
  right:  { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" },
  insideH: noBorder, insideV: noBorder,
};

// ─── Primitives ───────────────────────────────────────────────────────────────

/** TextRun with Arial font */
const r = (text, opts = {}) => new TextRun({
  text, font: "Arial",
  size:    opts.size    || 20,
  bold:    opts.bold    || false,
  color:   opts.color   || "000000",
  italics: opts.italics || false,
  allCaps: opts.allCaps || false,
});

/** Paragraph */
const p = (children, opts = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.before || 0, after: opts.after || 0, line: opts.line || 240 },
});

/** TableCell */
const cl = (children, opts = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width:   { size: opts.width || 4680, type: WidthType.DXA },
  borders: opts.borders || thinBorders,
  shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
  margins: { top: opts.mt||80, bottom: opts.mb||80, left: opts.ml||120, right: opts.mr||120 },
  verticalAlign: opts.va || VerticalAlign.TOP,
  columnSpan: opts.span,
});

/** Vertical spacer paragraph */
const sp = (pts = 120) => p(r(""), { before: pts });

// ─── Section Header ───────────────────────────────────────────────────────────
const sectionHeader = (title, badge = "") => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({ children: [
    cl([p([
      r(title.toUpperCase(), { bold: true, size: 18, color: WHITE, allCaps: true }),
      ...(badge ? [r("   " + badge, { size: 16, color: "DDDDDD" })] : []),
    ], { before: 60, after: 60 })],
    { width: 9360, borders: noBorders, shading: NAVY, ml: 160, mr: 160, mt: 100, mb: 100 })
  ]})]
});

// ─── Two-column field row (4-cell: label | value | label | value) ─────────────
const twoCol = (l1, v1, l2, v2) => new TableRow({ children: [
  cl([p(r(l1, { size: 18, bold: true, color: "555555" }), { before: 40, after: 40 })], { width: 1800, borders: noBorders, shading: GRAY_BG, ml: 120, mr: 80 }),
  cl([p(r(v1 || "—", { size: 18 }), { before: 40, after: 40 })], { width: 2880, borders: noBorders, ml: 80 }),
  cl([p(r(l2, { size: 18, bold: true, color: "555555" }), { before: 40, after: 40 })], { width: 1800, borders: noBorders, shading: GRAY_BG, ml: 120, mr: 80 }),
  cl([p(r(v2 || "—", { size: 18 }), { before: 40, after: 40 })], { width: 2880, borders: noBorders, ml: 80 }),
]});

// ─── Single field row (2-cell: label | value) ─────────────────────────────────
const fieldRow = (label, value) => new TableRow({ children: [
  cl([p(r(label, { size: 18, bold: true, color: "555555" }), { before: 40, after: 40 })], { width: 2200, borders: noBorders, shading: GRAY_BG, ml: 120, mr: 80 }),
  cl([p(r(value || "—", { size: 18 }), { before: 40, after: 40 })], { width: 7160, borders: noBorders, ml: 80 }),
]});

// ─── Deed Block ───────────────────────────────────────────────────────────────
/**
 * @param {number} num         - Deed number in chain
 * @param {string} type        - e.g. "Bargain and Sale"
 * @param {string} grantor
 * @param {string} grantee
 * @param {string} deedDate    - "MM/DD/YYYY" or "—"
 * @param {string} recDate     - "MM/DD/YYYY" or "—"
 * @param {string} book
 * @param {string} page
 * @param {string} inst        - Instrument number or "—"
 * @param {string} amount      - e.g. "$65,000.00" or null
 * @param {string} notes       - italic note or null
 * @param {boolean} isCurrent  - true = current vesting deed
 */
const deedBlock = (num, type, grantor, grantee, deedDate, recDate, book, page, inst, amount, notes, isCurrent) => {
  const badgeBg  = isCurrent ? BLUE_BG  : GRAY_BG;
  const badgeTxt = isCurrent ? "CURRENT VESTING" : "PRIOR";
  const badgeClr = isCurrent ? BLUE_TXT : "555555";

  const rows = [
    new TableRow({ children: [
      cl([p([r(`Deed #${num}`, { bold: true, size: 20 }), r(type ? `  —  ${type}` : "", { size: 18, color: "666666" })], { before: 60, after: 60 })], { width: 6960, borders: noBorders, ml: 120 }),
      cl([p(r(badgeTxt, { bold: true, size: 16, color: badgeClr }), { align: AlignmentType.RIGHT, before: 60, after: 60 })], { width: 2400, borders: noBorders, shading: badgeBg, mr: 120 }),
    ]}),
  ];

  if (amount) rows.push(new TableRow({ children: [
    cl([p(r(amount, { bold: true, size: 22, color: NAVY }), { before: 20, after: 40 })], { width: 9360, borders: noBorders, ml: 120, span: 2 })
  ]}));

  rows.push(new TableRow({ children: [
    cl([p(r("Grantor", { size: 16, bold: true, color: "777777" }), { before: 20, after: 10 }), p(r(grantor || "—", { size: 18 }), { after: 40 })], { width: 4680, borders: noBorders, ml: 120 }),
    cl([p(r("Grantee", { size: 16, bold: true, color: "777777" }), { before: 20, after: 10 }), p(r(grantee || "—", { size: 18 }), { after: 40 })], { width: 4680, borders: noBorders }),
  ]}));

  rows.push(new TableRow({ children: [
    cl([p(r("Deed Date",   { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(deedDate || "—", { size: 18 }))], { width: 2340, borders: noBorders, ml: 120 }),
    cl([p(r("Recorded",    { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(recDate  || "—", { size: 18 }))], { width: 2340, borders: noBorders }),
    cl([p(r("Book / Page", { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(book && page ? `${book} / ${page}` : "—", { size: 18 }))], { width: 2340, borders: noBorders }),
    cl([p(r("Inst. No.",   { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(inst || "—", { size: 18 }))], { width: 2340, borders: noBorders }),
  ]}));

  if (notes) rows.push(new TableRow({ children: [
    cl([p(r(notes, { size: 17, italics: true, color: "555555" }), { before: 20, after: 60 })], { width: 9360, borders: noBorders, ml: 120, span: 2 })
  ]}));

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680], rows, borders: boxBorders });
};

// ─── Mortgage Block — NONE ────────────────────────────────────────────────────
/**
 * Use when mortgages = NONE for period searched.
 * Shows Record Owner and "None for period searched." ONLY.
 * Never list prior cancelled instruments.
 */
const mortgageNone = (recordOwner) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  rows: [
    new TableRow({ children: [
      cl([p([r("Mortgage #1", { bold: true, size: 20 })], { before: 60, after: 60 })], { width: 6960, borders: noBorders, ml: 120 }),
      cl([p(r("NONE", { bold: true, size: 18, color: "555555" }), { align: AlignmentType.RIGHT, before: 60, after: 60 })], { width: 2400, borders: noBorders, shading: GRAY_BG, mr: 120 }),
    ]}),
    new TableRow({ children: [
      cl([
        p(r("Record Owner", { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }),
        p(r(recordOwner, { size: 18 }), { after: 20 }),
        p(r("None for period searched.", { size: 18, italics: true, color: "555555" }), { after: 60 }),
      ], { width: 9360, borders: noBorders, ml: 120, span: 2 }),
    ]}),
  ],
  borders: boxBorders,
});

// ─── Mortgage Block — OPEN ────────────────────────────────────────────────────
/**
 * Use when there is an open mortgage of record.
 * @param {number} num
 * @param {string} mortgagor
 * @param {string} mortgagee
 * @param {string} amount      - e.g. "$2,500,000.00"
 * @param {string} mortDate
 * @param {string} recDate
 * @param {string} book
 * @param {string} page
 * @param {string} inst
 * @param {string} notes       - e.g. assignment of leases/rents, payoff note
 */
const mortgageOpen = (num, mortgagor, mortgagee, amount, mortDate, recDate, book, page, inst, notes) => {
  const rows = [
    new TableRow({ children: [
      cl([p([r(`Mortgage #${num}`, { bold: true, size: 20 })], { before: 60, after: 60 })], { width: 6960, borders: noBorders, ml: 120 }),
      cl([p(r("OPEN", { bold: true, size: 18, color: RED_TXT }), { align: AlignmentType.RIGHT, before: 60, after: 60 })], { width: 2400, borders: noBorders, shading: RED_BG, mr: 120 }),
    ]}),
  ];

  if (amount) rows.push(new TableRow({ children: [
    cl([p(r(amount, { bold: true, size: 22, color: RED_TXT }), { before: 20, after: 40 })], { width: 9360, borders: noBorders, ml: 120, span: 2 })
  ]}));

  rows.push(new TableRow({ children: [
    cl([p(r("Mortgagor", { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(mortgagor || "—", { size: 18 }), { after: 40 })], { width: 4680, borders: noBorders, ml: 120 }),
    cl([p(r("Mortgagee", { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(mortgagee || "—", { size: 18 }), { after: 40 })], { width: 4680, borders: noBorders }),
  ]}));

  rows.push(new TableRow({ children: [
    cl([p(r("Mortgage Date", { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(mortDate || "—", { size: 18 }))], { width: 2340, borders: noBorders, ml: 120 }),
    cl([p(r("Recorded",      { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(recDate  || "—", { size: 18 }))], { width: 2340, borders: noBorders }),
    cl([p(r("Book / Page",   { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(book && page ? `${book} / ${page}` : "—", { size: 18 }))], { width: 2340, borders: noBorders }),
    cl([p(r("Inst. No.",     { size: 16, bold: true, color: "777777" }), { before: 0, after: 10 }), p(r(inst || "—", { size: 18 }))], { width: 2340, borders: noBorders }),
  ]}));

  if (notes) rows.push(new TableRow({ children: [
    cl([p(r(notes, { size: 17, italics: true, color: "555555" }), { before: 20, after: 60 })], { width: 9360, borders: noBorders, ml: 120, span: 2 })
  ]}));

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680], rows, borders: boxBorders });
};

// ─── Standard Report Header ───────────────────────────────────────────────────
const reportHeader = (orderNum, customer, searchType, refNum, boardDate) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [5200, 4160],
  rows: [new TableRow({ children: [
    cl([
      p(r("Independence Search & Abstract Co., Inc.", { bold: true, size: 24, color: NAVY }), { before: 80, after: 60 }),
      p(r("45 East Main St #207, Freehold, NJ 07728",            { size: 17, color: "444444" }), { after: 30 }),
      p(r("Office: (732) 431-1223  \u00B7  Fax: (732) 431-5757", { size: 17, color: "444444" }), { after: 30 }),
      p(r("Serving all 21 counties in the State of New Jersey",  { size: 16, italics: true, color: "666666" }), { after: 80 }),
    ], { width: 5200, borders: noBorders }),
    cl([
      p(r("SEARCH REPORT", { bold: true, size: 28, color: NAVY, allCaps: true }), { align: AlignmentType.RIGHT, before: 80, after: 60 }),
      p([r("Customer: ",    { bold: true, size: 18, color: "555555" }), r(customer   || "—", { size: 18 })], { align: AlignmentType.RIGHT, after: 30 }),
      p([r("Search Type: ", { bold: true, size: 18, color: "555555" }), r(searchType || "—", { size: 18 })], { align: AlignmentType.RIGHT, after: 30 }),
      p([r("Order #: ",     { bold: true, size: 18, color: "555555" }), r(orderNum   || "—", { size: 18, color: GOLD, bold: true })], { align: AlignmentType.RIGHT, after: 30 }),
      ...(refNum ? [p([r("Ref #: ", { bold: true, size: 18, color: "555555" }), r(refNum, { size: 18 })], { align: AlignmentType.RIGHT, after: 30 })] : []),
      p([r("Board Date: ",  { bold: true, size: 18, color: "555555" }), r(boardDate  || "—", { size: 18 })], { align: AlignmentType.RIGHT, after: 80 }),
    ], { width: 4160, borders: noBorders }),
  ]})],
  borders: { bottom: goldBorder, top: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
});

// ─── Standard Footer ──────────────────────────────────────────────────────────
const reportFooter = (orderNum) => new Footer({ children: [
  p([r(`Independence Search & Abstract Co., Inc.  \u00B7  Order #${orderNum}  \u00B7  CONFIDENTIAL`, { size: 16, color: "777777" })], { align: AlignmentType.CENTER, before: 60 }),
  p([r("Powered By TitleGistics\u2122 Real Estate Title AI Technology \u00A9", { size: 14, color: "999999", italics: true })], { align: AlignmentType.CENTER, after: 40 }),
]});

// ─── Standard Title Summary Box ───────────────────────────────────────────────
const titleSummaryBox = (extraNote = "") => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({ children: [
    cl([
      p(r("TITLE EXAMINATION SUMMARY", { bold: true, size: 18, color: NAVY, allCaps: true }), { before: 80, after: 60 }),
      p(r(
        "Title appears marketable and insurable subject to a thorough Title Examination and final acceptance by a Competent Title Examiner " +
        "and the Title Agency/Underwriter and the obtaining of NJ UC Judgment clearance as approved by the Title Agency/Underwriter and " +
        "the receipt of appropriate Affidavit(s) of Title for possible NJ UC Judgment Search matches." +
        (extraNote ? " " + extraNote : ""),
        { size: 18 }), { after: 80 }),
    ], { width: 9360, borders: noBorders, shading: "E8EDF3", ml: 200, mr: 200, mt: 60, mb: 60 })
  ]})]
});

// ─── Standard Disclaimer ──────────────────────────────────────────────────────
const disclaimer = () => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({ children: [
    cl([
      p(r("DISCLAIMER", { bold: true, size: 17, color: "555555" }), { before: 80, after: 40 }),
      p(r(
        "This report was prepared exclusively for Independence Search & Abstract Co., Inc. and its client(s). This report is not a guarantee of title, " +
        "a commitment to insure, or a policy of title insurance, and does not constitute a legal opinion on real estate title, which may only be rendered " +
        "by a licensed attorney. No warranty, expressed or implied, is made in connection with this report, including without limitation, merchantability or " +
        "fitness for a particular use or purpose. The liability of Independence Search & Abstract Co., Inc. shall not exceed the cost paid for the specific " +
        "search product giving rise to any claim. In no event shall Independence Search & Abstract Co., Inc. be liable for any indirect, incidental, " +
        "consequential, special, or punitive damages. The information herein has been compiled from publicly accessible databases and records maintained by " +
        "third-party governmental and quasi-governmental agencies beyond the care, custody, and control of Independence Search & Abstract Co., Inc.. " +
        "Independence Search & Abstract Co., Inc. assumes no responsibility for the accuracy, completeness, or currency of said public records, nor for any " +
        "consequences arising from inaccuracies or deficiencies therein.",
        { size: 16, color: "555555" }), { after: 80 }),
    ], { width: 9360, borders: { top: thinBorder, bottom: noBorder, left: noBorder, right: noBorder }, shading: GRAY_BG })
  ]})]
});

// ─── Document Config ──────────────────────────────────────────────────────────
const docConfig = (orderNum, sections) => ({
  styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      }
    },
    footers: { default: reportFooter(orderNum) },
    children: sections,
  }]
});

module.exports = {
  // tokens
  NAVY, GOLD, WHITE, GRAY_BG, RED_BG, RED_TXT, BLUE_BG, BLUE_TXT,
  GREEN_TXT, AMBER_TXT, AMBER_BG,
  // borders
  noBorder, noBorders, thinBorder, thinBorders, goldBorder, boxBorders,
  // primitives
  r, p, cl, sp,
  // components
  sectionHeader, twoCol, fieldRow,
  deedBlock, mortgageNone, mortgageOpen,
  reportHeader, reportFooter, titleSummaryBox, disclaimer, docConfig,
};
