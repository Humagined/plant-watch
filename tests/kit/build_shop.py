"""Build the synthetic shop: Northfork Precision, a fictional two-cell CNC job shop.

Everything here is invented. Seeded, so every run produces the same files.

Baseline: every source agrees with every other. Downtime in the MES export and
in the supervisor's spreadsheet means the same thing (every stop the MES logs,
changeovers included). The definitional conflict for the test is planted
separately by generator/plant_conflict.py, after review.

    python generator/build_shop.py          # writes shop/
"""

import csv
import random
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent / "shop"
TZ = timezone(timedelta(hours=-7))  # Mountain Standard Time
EXPORTED_AT = datetime(2026, 11, 18, 5, 30, tzinfo=TZ)
SITE = "Northfork"
rng = random.Random(4471)

DAYS = [date(2026, 11, d) for d in (9, 10, 11, 12, 13, 16, 17)]  # Mon-Fri, Mon-Tue
SHIFTS = {"1": (6, 0), "2": (14, 30)}   # 8.5 h each, 30 min unpaid lunch
SHIFT_PLANNED_MIN = 480

ASSETS = [
    # asset_id, name, make, model, cell, work_center
    ("MILL-VX2-01", "Mill 1", "Northfork shop", "VX2", "Mill cell", "WC-MILL"),
    ("MILL-VX2-02", "Mill 2", "Northfork shop", "VX2", "Mill cell", "WC-MILL"),
    ("MILL-HX4-01", "Mill 3", "Northfork shop", "HX4", "Mill cell", "WC-MILL"),
    ("LATHE-TX2-01", "Lathe 1", "Northfork shop", "TX2", "Turn cell", "WC-TURN"),
    ("LATHE-TX2-02", "Lathe 2", "Northfork shop", "TX2", "Turn cell", "WC-TURN"),
    ("LATHE-CX1-01", "Lathe 3", "Northfork shop", "CX1", "Turn cell", "WC-TURN"),
]

PARTS = [
    # part_number, description, revision, std_cost, work_center, setup_hours, run_min_per_unit
    ("PN-1001-A", "Manifold block, 6061", "C", 38.40, "WC-MILL", 1.0, 6.0),
    ("PN-1014-B", "Sensor housing, 6061", "B", 21.75, "WC-MILL", 0.75, 4.0),
    ("PN-1022-A", "Mounting plate, 304 SS", "A", 17.10, "WC-MILL", 0.75, 3.2),
    ("PN-1040-C", "Pump cover, 6061", "D", 44.90, "WC-MILL", 1.25, 8.0),
    ("PN-2003-A", "Shaft, 4140", "B", 12.60, "WC-TURN", 0.5, 2.4),
    ("PN-2011-B", "Bushing, 932 bronze", "A", 6.85, "WC-TURN", 0.5, 1.5),
    ("PN-2027-A", "Fitting, 316 SS", "C", 9.30, "WC-TURN", 0.5, 2.0),
    ("PN-2035-C", "Spacer, 6061", "A", 3.95, "WC-TURN", 0.25, 1.2),
]
PART = {p[0]: p for p in PARTS}

CUSTOMERS = ["Example Customer A", "Example Customer B", "Example Customer C",
             "Example Customer D", "Example Customer E"]

# MES reason codes. In the MES, every one of these counts as downtime.
REASONS = [
    # code, description, category
    ("CHG", "Changeover / setup", "Setup"),
    ("PM", "Planned maintenance", "Planned"),
    ("MECH", "Mechanical fault", "Unplanned"),
    ("ELEC", "Electrical / control fault", "Unplanned"),
    ("MAT", "Waiting on material", "Unplanned"),
    ("OPR", "No operator", "Unplanned"),
    ("QA", "Quality hold / first-article", "Unplanned"),
    ("TOOL", "Tool change / breakage", "Unplanned"),
]
SMALL_STOPS = [("MAT", "Bar stock not staged"), ("OPR", "Operator covering other machine"),
               ("QA", "Waiting on first-article sign-off"), ("TOOL", "Broken end mill"),
               ("TOOL", "Insert change"), ("ELEC", "Door interlock fault, reset"),
               ("MECH", "Chip conveyor jam")]


def ts(d: date, minute_of_shift: int, shift: str) -> datetime:
    h, m = SHIFTS[shift]
    return datetime(d.year, d.month, d.day, h, m, tzinfo=TZ) + timedelta(minutes=minute_of_shift)


def iso(t: datetime) -> str:
    return t.isoformat(timespec="minutes")


def write_csv(path: Path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def build():
    # ---------- simulate: each machine runs jobs back to back, minute by minute ----------
    sales_orders, wo_rows = [], []
    plan_rows, output_rows, downtime_rows, sheet_rows = [], [], [], []
    counters = {"so": 51200, "wo": 88300, "ev": 0, "lot": 7100}

    def new_job(asset_id, wc, start_day):
        part = rng.choice([p for p in PARTS if p[4] == wc])
        counters["so"] += 1
        counters["wo"] += 1
        qty = rng.choice([500, 600, 800, 1000]) if part[6] < 3 else rng.choice([150, 200, 250, 300])
        so, wo = f"SO-{counters['so']}", f"WO-{counters['wo']}"
        promise = start_day + timedelta(days=rng.choice([5, 7, 10, 14]))
        sales_orders.append([so, rng.choice(CUSTOMERS), part[0], qty, "each", promise.isoformat(), "Open"])
        job = {"wo": wo, "part": part, "qty": qty, "done": 0.0, "so": so, "promise": promise,
               "start": start_day, "booked": 0}  # booked = whole units reported, what the ERP shows
        wo_rows.append(job)
        return job

    for asset_id, _, _, _, cell, wc in ASSETS:
        job = new_job(asset_id, wc, DAYS[0])
        job["done"] = job["qty"] * rng.uniform(0.0, 0.4)  # already part-way through on Monday
        job["booked"] = int(job["done"])
        for d in DAYS:
            for shift in ("1", "2"):
                events = []                      # [start, end, code, text]
                fixed = []
                if asset_id == "MILL-VX2-02" and d == date(2026, 11, 12) and shift == "1":
                    fixed.append((30, 60, "PM", "Quarterly PM: way lube, coolant"))
                if asset_id == "LATHE-TX2-02" and d == date(2026, 11, 16) and shift == "2":
                    fixed.append((95, 165, "MECH", "Spindle drive alarm 2041, tech called"))
                segments = {}                    # wo -> [job, made, scrap_target, run_minutes]
                stop_until, stop = -1, None
                for minute in range(SHIFT_PLANNED_MIN):
                    if minute < stop_until:
                        continue
                    f = next((x for x in fixed if x[0] <= minute), None)
                    if f:
                        fixed.remove(f)
                        stop_until = minute + f[1]
                        events.append([minute, stop_until, f[2], f[3], job["wo"]])
                        continue
                    if rng.random() < 1.6 / SHIFT_PLANNED_MIN:
                        code, text = rng.choice(SMALL_STOPS)
                        stop_until = min(SHIFT_PLANNED_MIN, minute + rng.randint(6, 38))
                        events.append([minute, stop_until, code, text, job["wo"]])
                        continue
                    seg = segments.setdefault(job["wo"], [job, 0.0, 0])
                    rate = rng.uniform(0.84, 0.95) / job["part"][6]
                    seg[1] += rate
                    job["done"] += rate
                    if job["done"] >= job["qty"]:
                        job["done"] = job["qty"]
                        job = new_job(asset_id, wc, d)
                        dur = int(job["part"][5] * 60) + rng.randint(-10, 15)
                        stop_until = min(SHIFT_PLANNED_MIN, minute + 1 + dur)
                        events.append([minute + 1, stop_until, "CHG", f"Changeover to {job['part'][0]}", job["wo"]])
                prod_day = d.isoformat()
                for wo, (j, made_f, _) in segments.items():
                    made = int(made_f)
                    j["booked"] += made  # the ERP completion is the MES count (good + scrap)
                    scrap = int(made * rng.choice([0.0, 0.005, 0.01, 0.015, 0.03]))
                    counters["lot"] += 1
                    planned = int(made_f / 0.89 * 0.95) if made else 0
                    plan_rows.append([prod_day, shift, asset_id, wo, j["part"][0], planned, "each"])
                    output_rows.append([prod_day, shift, asset_id, wo, j["part"][0], f"L{counters['lot']}",
                                        made - scrap, scrap, "each", iso(ts(d, 8 * 60 + 30, shift))])
                    sheet_rows.append([d, shift, asset_id, wo, j["part"][0], made - scrap, scrap, None, None])
                down = 0
                for st, en, code, text, ev_wo in events:
                    counters["ev"] += 1
                    down += en - st
                    downtime_rows.append([f"DT-{counters['ev']:05d}", asset_id, iso(ts(d, st, shift)),
                                          iso(ts(d, en, shift)), en - st, code, text, ev_wo, shift, prod_day])
                # The log keeps one row per machine per shift; downtime goes on the shift's last job row.
                rows_this_shift = [r for r in sheet_rows if r[0] == d and r[1] == shift and r[2] == asset_id]
                if not rows_this_shift:  # down the whole shift: still one log row
                    sheet_rows.append([d, shift, asset_id, job["wo"], job["part"][0], 0, 0, None, None])
                    rows_this_shift = sheet_rows[-1:]
                rows_this_shift[-1][7] = down
                rows_this_shift[-1][8] = "; ".join(f"{e[2]} {e[1] - e[0]}m" for e in events)
                for r in rows_this_shift[:-1]:
                    r[7] = 0

    # ---------- ERP exports (stand-in for NetSuite) ----------
    erp = ROOT / "erp"
    write_csv(erp / "items.csv",
              ["part_number", "description", "revision", "uom", "standard_cost", "item_type", "last_modified"],
              [[p[0], p[1], p[2], "each", f"{p[3]:.2f}", "Assembly/Bill of Materials", "2026-10-02T09:14-07:00"]
               for p in PARTS])
    write_csv(erp / "sales_orders.csv",
              ["sales_order", "customer", "part_number", "quantity", "uom", "promise_date", "status"],
              sales_orders)
    wos = []
    for j in wo_rows:
        done = j["booked"]
        status = "Closed" if j["done"] >= j["qty"] else ("In Process" if done else "Released")
        wos.append([j["wo"], j["part"][0], j["part"][2], j["qty"], done, "each", status, j["start"].isoformat(),
                    j["promise"].isoformat(), j["so"]])
    write_csv(erp / "production_work_orders.csv",
              ["production_wo", "part_number", "revision", "qty_ordered", "qty_completed", "uom", "status",
               "start_date", "due_date", "sales_order"], wos)
    write_csv(erp / "routings.csv",
              ["part_number", "revision", "op_seq", "operation", "work_center", "setup_hours", "run_min_per_unit",
               "effective_date"],
              [[p[0], p[2], 10, "Machine complete", p[4], p[5], p[6], "2026-01-05"] for p in PARTS])
    write_csv(erp / "inventory.csv",
              ["part_number", "location", "on_hand", "uom", "as_of"],
              [[p[0], "FG-Northfork", rng.randint(0, 400), "each", iso(EXPORTED_AT)] for p in PARTS])

    # ---------- MES, planning and CMMS exports (the export folder) ----------
    ex = ROOT / "exports"
    write_csv(ex / "mes_production_output_2026-11-09_to_2026-11-17.csv",
              ["production_day", "shift", "asset_id", "production_wo", "part_number", "lot", "good_qty",
               "scrap_qty", "uom", "reported_at"], output_rows)
    write_csv(ex / "mes_downtime_events_2026-11-09_to_2026-11-17.csv",
              ["event_id", "asset_id", "started_at", "ended_at", "duration_min", "reason_code", "reason_text",
               "production_wo", "shift", "production_day"], downtime_rows)
    write_csv(ex / "mes_reason_codes.csv",
              ["reason_code", "description", "category", "counts_as_downtime"],
              [[c, dsc, cat, "Y"] for c, dsc, cat in REASONS])
    write_csv(ex / "planning_schedule_2026-11-09_to_2026-11-17.csv",
              ["production_day", "shift", "asset_id", "production_wo", "part_number", "planned_qty", "uom"],
              plan_rows)
    write_csv(ex / "cmms_assets.csv",
              ["asset_id", "name", "make", "model", "cell", "work_center", "status"],
              [[a[0], a[1], a[2], a[3], a[4], a[5], "Active"] for a in ASSETS])
    write_csv(ex / "cmms_maintenance_work_orders.csv",
              ["maintenance_wo", "asset_id", "type", "opened_at", "closed_at", "description", "status"],
              [["MWO-3310", "MILL-VX2-02", "PM", "2026-11-12T06:32-07:00", "2026-11-12T07:32-07:00",
                "Quarterly PM: way lube, coolant concentration, filters", "Closed"],
               ["MWO-3324", "LATHE-TX2-02", "Corrective", "2026-11-16T16:05-07:00", "2026-11-16T18:50-07:00",
                "Spindle drive alarm 2041. Reseated drive connector, replaced cooling fan.", "Closed"],
               ["MWO-3329", "LATHE-CX1-01", "PM", "2026-11-17T22:00-07:00", None,
                "Monthly PM: hydraulic filter, chuck lube", "Open"]])

    # ---------- the supervisor's spreadsheet ----------
    wb = Workbook()
    ws = wb.active
    ws.title = "Shift Log"
    ws.append(["Date", "Shift", "Machine", "Job", "Part", "Good", "Scrap", "Downtime (min)", "Stops"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in sheet_rows:
        ws.append(r)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["I"].width = 48
    notes = wb.create_sheet("Notes")
    notes.append(["Northfork Precision: supervisor shift log"])
    notes.append(["Kept by the 1st shift lead. Filled in from the MES screen at end of shift."])
    notes.append(["Downtime (min) = all stops logged on the MES screen for the shift."])
    (ROOT / "spreadsheets").mkdir(parents=True, exist_ok=True)
    wb.save(ROOT / "spreadsheets" / "supervisor_shift_log_Nov.xlsx")

    print(f"Wrote {ROOT}")
    print(f"  {len(PARTS)} parts, {len(sales_orders)} sales orders, {len(wos)} production work orders")
    print(f"  {len(output_rows)} output rows, {len(downtime_rows)} downtime events, {len(sheet_rows)} log rows")


if __name__ == "__main__":
    build()
