---
name: add-report
description: Adds a report to Plant Watch so its figures get checked against the plant's data. Reads an uploaded PDF or a linked SharePoint/OneDrive report, shows the figures it found with the page and the quoted line, asks the person once what each figure means, and saves a template so later editions are read the same way. Use when someone says "check this report", "add this report", "here's our weekly pack", shares a report PDF or link, or when a report's layout has changed.
---

# Add a report

The script compares; you read. A reading you can't place is a question for the person,
never a guess.

Script: `python3 "<this skill's folder>/../../scripts/watch.py"`, state in
`plant-watch/state`.

## Read it

- **PDF**: read the file. Run `pdf-text <file>` too: if it returns text, quote lines
  exactly as it shows them (the run checks every quote against the page).
- **SharePoint / OneDrive link**: read it with the Microsoft 365 connector
  (`read_resource` on its `file:///` URI). Record its last-modified time in
  `plant-watch/state/links.json` as `{"<source id>": {"last_modified": "...", "as_of": "<date the report says it covers, if any>"}}`.
  If the link can't be read, record `{"error": "..."}` and tell the person.
- Other links (Power BI, Google Drive, a BI tool's web page): say plainly they can't be
  read directly yet, and ask for an export (PDF, Excel or CSV) instead.

## Confirm what it means, once

Show the figures as a short table: label as printed, value, page, the quoted line. Then
for each distinct label ask what it is: which metric (on the floor: downtime, good output,
scrap, planned, availability, performance, capacity hours; beyond it: orders booked, order
target, shipped, budget, % to budget, forecast, working days; the ids are in
`data/metrics.json`), which machines or lines (one, or the plant total), which period,
which unit. Also copy any footnote or definition note word for word.

Save the template to `plant-watch/state/templates/<series>.json`:
`{"series": "...", "labels": {"Downtime (min)": {"metric": "downtime", "unit": "min"}}, "period_rule": "the week in the title", "footnote_at_confirmation": "..."}`.

## Save the reading

Write `plant-watch/state/extractions/<file name without extension>.json`:

```json
{"file": "Weekly_OEE_pack_2026-11-20.pdf", "series": "Weekly OEE pack",
 "period": {"start": "2026-11-16", "end": "2026-11-20"}, "published": "2026-11-23",
 "footnote": "Downtime excludes changeovers, which are planned. ...",
 "figures": [{"metric": "downtime", "entity": "MILL-VX2-01", "value": 444, "unit": "min",
              "page": 1, "quote": "MILL-VX2-01 578 3 444 90.8%"},
             {"metric": "downtime", "entity": "total", "value": 1980, "unit": "min",
              "page": 1, "quote": "Plant total 7,107 95 1,980 93.1%"}]}
```

Values as plain numbers (no commas or units). `entity` is a machine id the data uses, or
`total`. Later editions of the same series: read them against the saved template without
asking again, unless a label is new or missing, then ask about that label only.

**A spreadsheet or linked report** (Excel, CSV, a SharePoint/OneDrive workbook) gets the same
reading, with three differences: cite each figure by `sheet` and `cell` instead of `page`
(the run checks the cell holds the value); give each figure its own `date` when the report
has a row per day (or `period` for a row per week); and add `"source": "<the link's source id>"`.
A linked report is refreshed in place, so it has one reading,
`plant-watch/state/extractions/<source id>.json`, rewritten whenever its last-modified time
changes.

```json
{"file": "daily_production_summary.xlsx", "source": "daily_summary", "series": "Daily production summary",
 "footnote": null,
 "figures": [{"metric": "downtime", "entity": "total", "date": "2026-11-12", "value": 629, "unit": "min",
              "sheet": "Daily", "cell": "D5"}]}
```

Read every row, not a sample: the run compares each figure with the data for its own day.
