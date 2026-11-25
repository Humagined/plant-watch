# sources.json

One entry per source. `glob` is relative to `folder`. Only include what the files
actually show; leave a field out rather than guess.

```json
{
  "plant": "Northfork Precision",
  "folder": "/Users/pat/Plant Exports",
  "sources": [
    {"id": "mes_downtime", "label": "MES downtime events", "glob": "exports/mes_downtime_events_*.csv",
     "kind": "events", "date": "production_day", "shift": "shift", "entity": "asset_id",
     "measures": {"downtime": "duration_min"}, "category": "reason_code", "record_for": ["downtime"]},
    {"id": "mes_output", "label": "MES production output", "glob": "exports/mes_production_output_*.csv",
     "kind": "rows", "date": "production_day", "shift": "shift", "entity": "asset_id",
     "key": "production_wo", "key_master": "erp_wo", "measures": {"good": "good_qty", "scrap": "scrap_qty"},
     "record_for": ["good", "scrap"]},
    {"id": "shift_log", "label": "Supervisor shift log", "glob": "spreadsheets/supervisor_shift_log_*.xlsx",
     "sheet": "Shift Log", "kind": "rows", "date": "Date", "shift": "Shift", "entity": "Machine",
     "key": "Job", "key_master": "erp_wo", "measures": {"good": "Good", "scrap": "Scrap", "downtime": "Downtime (min)"},
     "notes_sheet": "Notes", "stage": "prod::Production Reporting"},
    {"id": "reason_codes", "label": "MES reason codes", "glob": "exports/mes_reason_codes.csv",
     "kind": "codes", "code": "reason_code", "flags": {"counts_as_downtime": "downtime"}},
    {"id": "erp_wo", "label": "ERP production work orders", "glob": "erp/production_work_orders.csv",
     "kind": "master", "key": "production_wo", "stage": "prod::Work Order"},
    {"id": "weekly_pack", "label": "Weekly OEE pack", "glob": "reports/Weekly_OEE_pack_*.pdf",
     "kind": "report", "series": "Weekly OEE pack", "measures": ["good", "scrap", "downtime"]},
    {"id": "daily_summary", "label": "Daily production summary (SharePoint)", "kind": "link",
     "uri": "file:///<driveId>/<itemId>", "path": null, "expect": "daily",
     "measures": ["good", "scrap", "downtime"], "stage": "prod::Production Reporting"}
  ]
}
```

| kind | What it is | Needs |
|---|---|---|
| `events` | One row per event with a duration and a reason code | date, entity, measures, category |
| `rows` | Rows of quantities per day / shift / machine / job | date, entity, measures |
| `codes` | A code list with flags such as `counts_as_downtime` | code, flags |
| `master` | The list of valid keys (work orders) | key |
| `report` | PDF reports read by add-report | series |
| `link` | A report reached by link | uri (Microsoft 365) or path |

Metric names for `measures`: `downtime`, `good`, `scrap`, `planned`, `availability`,
`planned_minutes` (see `data/metrics.json`). `record_for` marks the source the plant
treats as the record; the dictionary can override it. `stage` places a source on a value
stream stage (`data/value_streams.json` keys, e.g. `prod::Downtime`).
