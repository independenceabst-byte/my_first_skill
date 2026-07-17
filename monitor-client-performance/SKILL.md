---
name: monitor-client-performance
description: Monitor weekly job performance for the top 20 clients using the Google Drive workbooks CLIENT MASTER PERFORMANCE 2026 and NEW JOB COUNTIES 2026. Use when Codex needs to analyze client job volume, detect weekly declines of 30% or more, create or update a recurring client-performance alert, or review the monitor's results.
---

# Client Performance Monitor

Use this skill to produce a source-backed weekly client-performance readout and, when requested, maintain its recurring alert.

## Sources and metric contract

Find the following Google Drive spreadsheets by exact title before using them:

- `CLIENT MASTER PERFORMANCE 2026`: monthly client totals, month-over-month change, and monthly context.
- `NEW JOB COUNTIES 2026`: dated job-level records. Use this as the actuals source for weekly monitoring.

Confirm that each relevant `* Data` tab in `NEW JOB COUNTIES 2026` contains `DATE`, `CLIENT NAME`, and `VALUE`. Count performance as the sum of numeric `VALUE` by client and date. Treat one dated row as a job only when `VALUE` is valid; do not infer a value from blank cells.

Do not use `CLIENT MASTER PERFORMANCE 2026`'s monthly `Difference` or `% Change` fields as a weekly decline signal. Use that workbook only to provide monthly context or resolve naming differences.

## Weekly analysis

1. Determine the last completed Sunday in the user's timezone. Compare only complete Monday-Sunday weeks; never include a partial current week.
2. Gather the dated raw-job rows needed for the trailing 90 complete calendar days ending on that Sunday. Read every relevant monthly `* Data` tab when the period crosses months.
3. Rank clients by total jobs over that 90-day window and keep the top 20. Preserve the exact source names unless a documented naming map proves two variants are the same client.
4. For each top-20 client, compare the latest completed week with the immediately prior completed week:

   `decline % = (prior week jobs - latest week jobs) / prior week jobs`

   Flag an alert when the decline is at least 30%. Do not calculate a percentage decline when the prior week has zero jobs.
5. Use the monthly client-master sheet only as supplemental context. If its latest month is behind the raw-job data, state that the weekly calculation still uses the dated raw records.

## Quality checks

- Verify the selected date range and exclude headers, blank client names, invalid dates, and nonnumeric `VALUE` cells.
- Check source freshness: state the latest dated raw job and flag missing recent days before calling a result current.
- Do not deduplicate raw rows without a reliable job identifier. If duplicate risk is material, state it rather than silently removing rows.
- Keep calculation dates and calendar-week boundaries visible in the result.

## Weekly readout and alerts

Lead with the alert state. For every triggered client, show:

- Client name
- Latest completed-week jobs and prior-week jobs
- Percentage decline
- Both week ranges

Always include a concise weekly summary even if there are no alerts: calculation period, top-20 ranking basis, number of alerts, and material freshness or data-quality caveats. Link both source spreadsheets.

Use Data Analytics KPI reporting for a stakeholder-facing readout. Load metric diagnostics only if the user also asks why a client declined; do not present a hypothesis as a cause.

## Recurring monitor

When the user asks to create, change, pause, or resume the monitor:

1. Search for the automation tool first. Inspect existing automations and update a matching monitor instead of creating a duplicate.
2. Create a weekly local automation only when the user requests it. If no schedule is given, use Monday morning in the user's timezone and state the assumption.
3. Put the complete metric contract, source titles, week definition, 90-day top-20 rule, 30% threshold, and alert content in the automation prompt.
4. Instruct the automation to send a concise weekly summary and make a triggered alert prominent. It must never edit either source spreadsheet.

If the source layout changes or no fresh raw-job data is available, report the blocker and do not fabricate an alert.
