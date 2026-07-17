---
name: monitor-searcher-production
description: Monitor weekly job production for every person listed in the Office and Sub-contractors dashboards in Notion. Use when Codex needs to analyze each searcher's production, detect a 30% or greater week-over-week decline, create or update the Friday 3 PM ET alert, or review its results.
---

# Monitor Searcher Production

## Sources and roster

- Fetch the Notion **Dashboard Metrics** page at the start of each run.
- Treat the linked Office and Sub-contractors dashboard sections as the current roster. Do not hard-code names.
- Follow each person's dashboard link to its task datasource. Do not edit Notion pages or records.

## Metric contract

- Count a completed job when `Upload File` is non-empty.
- Use `Last Edited Time` as the completion timestamp proxy. State this caveat only if it could materially affect the result.
- At the Friday 3 PM ET run, compare the current Monday 12:00 AM–Friday 3:00 PM ET window against the equivalent Monday–Friday 3 PM window in the preceding week.
- Calculate decline as `(prior production - current production) / prior production`.
- Flag only a decline of 30% or more when prior-week production is greater than zero.

## Workflow

1. Fetch the dashboard roster and each linked task datasource.
2. Count qualifying jobs per person for both like-for-like windows.
3. Check that every datasource exposes `Upload File` and `Last Edited Time`; report a source blocker rather than producing a misleading comparison.
4. Return or send an alert only when at least one person meets the threshold, or when a source blocker prevents the check.

## Alert content

For each flagged person, include the Office or Sub-contractors group, current and prior production, percentage decline, comparison windows, and a link to the relevant Notion dashboard. Do not infer causes for a decline.

## Recurring monitor

Create or maintain one active recurring monitor for Friday at 3:00 PM in the America/New_York timezone. Reuse the existing monitor when updating the schedule or logic; do not create duplicates.
