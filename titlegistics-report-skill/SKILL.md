---
name: titlegistics-report
description: >
  Generates a formatted ISA TitleGistics-style Search Report (.docx) from raw NJ title search PDFs.
  Trigger when user uploads a title search PDF and asks to create a report, generate a TitleGistics
  report, "do similar report", or produce a Word doc from a searcher file. Handles scanned and text
  PDFs — extracts deed chain, mortgages, judgments, legal description, tax data, and misc instruments,
  then builds a Navy/Gold branded ISA Word report with all 13 sections: subject property, title summary,
  tax info, deed chain blocks, mortgage blocks (OPEN/NONE), judgments, miscellaneous instruments,
  additional instruments, legal description (typed verbatim from vesting deed), open items, and
  disclaimer. Always trigger for any NJ title search PDF + report request, even "do the same as before".
---

# TitleGistics Report Skill

Produces a fully formatted ISA Search Report (.docx) from raw NJ title search PDFs. The output matches the TitleGistics™ report style: navy/gold branding, structured sections, deed chain blocks, mortgage blocks, typed legal description from the vesting deed, open items, and disclaimer.

## Step 1 — Read the PDF

Use the `pdf` skill to inspect the source and the `documents` skill to render and validate the finished DOCX. The PDF may be scanned (Canon iR-ADV) or text-based. First obtain the page count and run a layout-preserving text extraction.

If text extraction returns blank pages (all `\f`), treat the PDF as scanned. Render pages 1–6 for the cover, run sheet, tax data, and order letter; render the final 10 pages for the legal description and current vesting deed. Inspect those page images visually.

## Step 2 — Data Extraction

Extract all fields from the pages. Key sources:

| Data | Source Page(s) |
|------|---------------|
| Order #, Municipality, Block/Lot, Record Owner, Mortgages summary, Judgments, Grants, Restrictions, Board Date, Copies | Page 1 — ISA cover sheet |
| Full deed chain (book/page), mortgage run, vesting deed(s), judgment types checked | Pages 2-6 — Searcher run sheets |
| Tax data (assessed value, land/improvements, property tax, lot size, year built) | NJTaxMaps.com property report page |
| Address, County, Block/Lot confirmation | Rise Abstract order letter |
| Legal description (Schedule A) | Last ~10 pages — current vesting deed |
| Instrument numbers, exact recorded dates | Mercer/Monmouth County Document Summary Sheets (highlighted recording info) |

**If supplemental recording info PDFs are attached**, read them for exact instrument numbers and recorded dates — these override anything inferred from searcher notes.

### Reading Searcher Run Sheets

The handwritten run sheets use shorthand:
- `6574-877` = Book 6574, Page 877
- `Canc 5/4/23` = Cancelled 5/4/2023 (prior mortgage, do NOT list in report)
- `FYI DNA PR` = For Your Information, Do Not Affect Period of Record
- `POA` = Power of Attorney instrument
- `LP` = Lis Pendens
- `Am` = Amendment
- `R` = Release
- `E n d` = end of chain for that column
- Owner names with date ranges = ownership periods for that vesting

Build the deed chain in chronological order from oldest to current.

## Step 3 — Mortgage Rule (CRITICAL)

**When mortgages status is NONE:**
- Show only: Badge "NONE", "Record Owner: [name]", "None for period searched."
- **NEVER list prior cancelled mortgage instruments, book/page numbers, or cancellation dates**

**When mortgages are OPEN:**
- Show full block with mortgagor, mortgagee, amount, dates, book/page, inst. no.
- Bold amount in red
- Note: "This mortgage must be satisfied or paid off at or before closing."

## Step 4 — Build the Report Script

Call `codex_app__load_workspace_dependencies` to locate the bundled Node runtime and its packages. Use that Node executable with `NODE_PATH` set to the returned Node packages directory; do not globally install `docx` or other packages.

Write the build script in the active workspace. Require this skill's bundled `scripts/helpers.js` using its absolute path (or copy it beside the build script), then use the patterns below.

### Design Tokens
```javascript
const NAVY    = "0A1A2E";
const GOLD    = "C9A84C";
const WHITE   = "FFFFFF";
const GRAY_BG = "F5F5F5";
const RED_BG  = "FCEBEB";
const RED_TXT = "A32D2D";
```

### Required Imports
```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, Footer,
} = require('docx');
```

### Core Helpers (copy exactly)

See `scripts/helpers.js` for the full set of reusable primitives: `r()`, `p()`, `cl()`, `sp()`, `sectionHeader()`, `twoCol()`, `fieldRow()`, `deedBlock()`, `mortgageNone()`, `mortgageOpen()`.

### Report Sections (in order)

1. **Report Header** — ISA logo text left, order info right, gold bottom border
2. **Subject Property** — twoCol grid: address, county, block/lot, tax map, record owner, vesting, effective date, completed, searcher run, copies
3. **Title Examination Summary** — light navy box with standard marketable/insurable language
4. **Tax Information** — 3-column grid: amount due/year, assessed value, land/improvements; second row: property type, lot size, year built
5. **Deed Chain** — one `deedBlock()` per instrument, separated by `sp(80)`. Label count in section header. Last deed = "CURRENT VESTING" badge (blue). Prior deeds = "PRIOR" badge (gray).
6. **Mortgages & Deeds of Trust** — `mortgageNone()` or `mortgageOpen()` per mortgage
7. **Judgments, UCC Filings & Liens** — table with NJ UC search result + checkmarks for all types searched
8. **Miscellaneous Instruments** — fieldRow table: Grants, Restrictions, Setback, Flood Cert, Tidelands, OFAC/Patriot, Declarations of Taking
9. **Additional Instruments** (if any) — FYI items, POAs, final judgments, vacations, lis pendens
10. **Legal Description** — typed verbatim from Schedule A of current vesting deed; cite Book/Page/CFN/Recorded date
11. **Open Items Requiring Action** — colored box: green if no open mortgages, red if open mortgage(s)
12. **Disclaimer** — standard ISA disclaimer text, gray background
13. **Footer** — "ISA · Order #[num] · CONFIDENTIAL" centered; "Powered By TitleGistics™..." below

### Page Setup
```javascript
page: {
  size: { width: 12240, height: 15840 },   // US Letter
  margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
}
```

### Output Path
```javascript
const path = require("path");
const outputDir = path.resolve(process.cwd(), "outputs");
fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(path.join(outputDir, "[ORDERNUM]_[NAME]_ISA_SEARCH_REPORT.docx"), buf);
```

## Step 5 — Build and Validate

Run the build script with the bundled Node runtime. Then use the `documents` skill's render-and-verify workflow on the produced `.docx`; inspect the rendered pages and correct layout, clipping, overflow, or XML issues before delivering the file.

## Critical Rules

- **Legal description**: Always type it in full from the actual deed pages — never summarize or abbreviate. Cite the source deed (Book/Page/CFN/Recorded date).
- **Mortgage NONE**: Show Record Owner + "None for period searched." only. No prior instrument details.
- **Instrument numbers**: If recording info PDFs are attached with highlighted data, use those exact numbers. They override inferred data.
- **FYI / DNA PR instruments**: List under Additional Instruments, note they do not affect title.
- **POA deeds**: Flag in deed notes: "POA instrument at Book X/Page Y — to be reviewed and confirmed valid by examining attorney."
- **Vacated judgments**: Note both the judgment and its vacation instrument.
- **Table column widths**: Always use DXA, always set both `columnWidths` on table AND `width` on each cell. They must sum correctly.
- **No unicode bullets** — use `LevelFormat.BULLET` with numbering config if lists needed.
- **US Letter page size** — docx defaults to A4; always override explicitly.

## Reference: Block/Lot Identification

- Monmouth County searches: Block/Lot from ISA cover sheet (e.g., Block 72 / Lot 35)
- Mercer County / Trenton City searches: Block/Lot from ISA cover sheet (e.g., Block 8203 / Lot 1, Block 2605 / Lot 29)
- Tax Map number also on cover sheet

## Reference: Deed Chain Patterns

**Trenton City searches** often have long chains (8–12+ instruments) with:
- Tax lien deed → municipality → investor chains
- Multiple quick flips in short windows
- POA conveyances
- Final judgment / vacation pairs

**Monmouth County searches** tend to have cleaner 3–6 instrument chains with:
- Individual → individual → LLC patterns
- 1031 exchange notes
- Life tenancy reservations
