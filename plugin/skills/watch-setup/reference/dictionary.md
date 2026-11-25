# dictionary.json (shape 2)

One entry per figure, confirmed by a person. The script judges every source against it, and
`export` saves the confirmed entries as a file the person keeps. A version-1 file (only
`status`, `meaning`, `record`, `exclude_categories`, `definition`) still reads.

```json
{"schema": "plant-watch.dictionary/2", "version": 4, "plant": "Northfork Precision",
 "metrics": {
  "downtime": {"label": "Downtime", "status": "confirmed",
               "meaning": "Every MES stop counts, changeovers and planned maintenance included.",
               "unit": "min", "grain": "machine x shift", "measure": "sum",
               "dated_by": "production day of the stop's start",
               "counts": ["every stop the MES records"], "excludes": [], "exclude_categories": [],
               "record": "mes_downtime",
               "owner": {"name": "Pat Lee", "role": "Plant manager"},
               "confirmed_by": "Pat Lee", "confirmed_on": "2026-11-25",
               "decisions": [
                 {"key": "changeovers_count", "question": "Do changeovers count as downtime?",
                  "answer": "Yes", "rationale": "They take capacity."},
                 {"key": "availability_clock", "question": "Availability against which clock?",
                  "answer": "Against scheduled time"}],
               "history": [{"version": 1, "on": "2026-11-18", "by": "Pat Lee", "status": "not_settled",
                            "why": "MES and shift log disagreed on changeovers"}]},
  "scrap": {"label": "Scrap", "status": "not_settled", "note": "Quality and the floor disagree on rework."}
 },
 "formulas": {
  "availability": {"label": "Availability", "status": "confirmed", "metric": "availability",
                   "definition": "[Run Hours]/([Planned Hours]-[PM Hours])",
                   "meaning": "Planned maintenance doesn't count against availability.",
                   "grain": "machine x month", "measure": "ratio",
                   "owner": {"name": "Pat Lee", "role": "Plant manager"},
                   "confirmed_by": "Pat Lee", "confirmed_on": "2026-11-25",
                   "variants_seen": [{"definition": "[Run Hours]/[Planned Hours]", "places": 9},
                                     {"definition": "[Run Hours]/([Planned Hours]-[PM Hours])", "places": 2}]}
 }}
```

What each field is for:
- `status`: `confirmed`, `not_settled` or `retired`. A not-settled figure is still reported,
  as a disagreement with the question to settle.
- `meaning`: the sentence a person would say. `counts` / `excludes`: what's in and out, in
  plain words. `exclude_categories`: reason codes that don't count (machine-checked).
- `unit`, `grain` (what one value is about: machine x shift, order line, month), `measure`
  (`sum`; `ratio`, where a quarter is total numerator over total denominator; or `period`, true
  only one period at a time), `dated_by` (which date puts a value in a period).
- `record`: the source id that is the record for this figure.
- `owner`: who is accountable for the number (name and role). Required once confirmed.
  `confirmed_by` / `confirmed_on`: who agreed and when (may be the same person).
- `decisions`: every question asked and the answer, with the reason when they gave one.
- `formulas`: figures defined by a spreadsheet formula, keyed by the label in lower case.
  `definition` is copied exactly from one of the inventory's definitions (the script compares by
  that text); `metric` names the Plant Watch metric it belongs to, if any; `variants_seen` keeps
  the other definitions the inventory listed.
- `version`: bump it whenever an entry changes, and append the old entry's status, meaning and
  why to its `history`. Never edit or delete a history line.

How to ask, one figure at a time:

> "Here's how your systems count downtime. The MES counts every stop, changeovers (CHG)
> and planned maintenance (PM) included. The shift log's Notes say that from 12 Nov it
> leaves out changeovers, and its total is 1,240 minutes lower. What should downtime mean
> at your plant? (It's fine to say 'not settled yet'.)"

Then: "Who is accountable for this number?"

Use only facts the inventory returned. Don't recommend an answer unless asked; if asked, say
what each choice does to the numbers. Run `check-dictionary` before saving; fix what it reports.
