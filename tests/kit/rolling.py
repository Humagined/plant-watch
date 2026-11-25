"""Rolling Northfork: plant exports that land day by day, with dated conflicts,
Excel faults, a weekly PDF report and a linked report. Everything is invented.

The plant folder is a pure function of the "through" date, so rebuilding with an
earlier date is the reset:

    python generator/rolling.py --through 2026-11-12            # state after day 4
    python generator/rolling.py --through 2026-11-25            # all 13 days
    python generator/rolling.py --through 2026-11-25 --out /tmp/plant

Writes <out>/ (what the tool reads) and, beside it, answer_key_<date>.json (what a
correct tool should find; never inside the plant folder) and excel_saved/ (the
supervisor workbook as Excel would save it, with cached error values).

The Phase 1 kit (build_shop.py + plant_conflict.py) is untouched.
"""

import argparse
import csv
import json
import os
import random
import re
import shutil
import time
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName

import pdfmini
from build_shop import ASSETS, CUSTOMERS, PARTS, REASONS, SHIFT_PLANNED_MIN, SMALL_STOPS, TZ, iso, ts, write_csv

SEED = 5521
DAYS = [date(2026, 11, d) for d in (9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 23, 24, 25)]
D = {f"day{i}": d for i, d in enumerate(DAYS, 1)}

# The dated scenario. Days 1-3 are clean on purpose: normal variation, no findings.
CHG_EXCLUDED_FROM = D["day4"]      # supervisor log stops counting changeovers
KEY_DRIFT_FROM = D["day5"]         # 2nd-shift job numbers lose WO-, one transposed
OVERWRITTEN = D["day6"]            # a typed number replaces a Summary formula
HIDDEN_ROW = D["day7"]             # a hidden row inside the summed range
DELETED_COLUMN = D["day8"]         # #REF! in Summary scrap formulas and a defined name
SCRAP_SPIKE = D["day9"]            # real anomaly: LATHE-TX2-01 scrap jumps
SUM_SHORT = D["day10"]             # grand total stops two rows short
PATTERN_BREAK = D["day11"]         # bad copy + hard-coded -65 in Shift Log col J
WRONG_PERIOD = D["day12"]          # Oct sheet, external link, circular reference
DIV_BLANK = D["day13"]             # divide by blank planned minutes, VLOOKUP past range
WEEK1_PUBLISHED = D["day6"]        # weekly pack for 9-13 Nov, agrees with the MES
WEEK2_PUBLISHED = D["day11"]       # weekly pack for 16-20 Nov, follows the log, new footnote
LINK_LAST_REFRESH = D["day10"]     # the linked daily summary stops updating after this


def rng_for(*parts):
    return random.Random("-".join(str(p) for p in (SEED,) + parts))


def simulate():
    """Simulate all 13 days. Day-outer, machine-inner, so numbers rise with time;
    each (day, machine) has its own random stream."""
    jobs, so_rows = [], []
    out, plan, down, log = [], [], [], []
    ctr = {"so": 51300, "wo": 88400, "ev": 0, "lot": 7300}
    current = {}

    def new_job(asset, wc, day, r):
        part = r.choice([p for p in PARTS if p[4] == wc])
        ctr["so"] += 1
        ctr["wo"] += 1
        qty = r.choice([500, 600, 800, 1000]) if part[6] < 3 else r.choice([150, 200, 250, 300])
        job = {"wo": f"WO-{ctr['wo']}", "so": f"SO-{ctr['so']}", "part": part, "qty": qty, "done": 0.0,
               "start": day, "promise": day + timedelta(days=r.choice([5, 7, 10, 14])), "made": {},
               "closed": None, "customer": r.choice(CUSTOMERS)}
        jobs.append(job)
        return job

    for asset, *_rest in ASSETS:
        wc = _rest[-1]
        r = rng_for(asset, "init")
        job = new_job(asset, wc, DAYS[0], r)
        job["done"] = job["qty"] * r.uniform(0.0, 0.4)
        job["pre"] = int(job["done"])
        current[asset] = job

    for d in DAYS:
        for asset, *_rest in ASSETS:
            wc = _rest[-1]
            r = rng_for(asset, d.isoformat())
            job = current[asset]
            for shift in ("1", "2"):
                events, fixed = [], []
                if asset == "MILL-VX2-02" and d == date(2026, 11, 12) and shift == "1":
                    fixed.append((30, 60, "PM", "Quarterly PM: way lube, coolant"))
                if asset == "LATHE-TX2-02" and d == date(2026, 11, 16) and shift == "2":
                    fixed.append((95, 70, "MECH", "Spindle drive alarm 2041, tech called"))
                segs, until = {}, -1
                for minute in range(SHIFT_PLANNED_MIN):
                    if minute < until:
                        continue
                    f = next((x for x in fixed if x[0] <= minute), None)
                    if f:
                        fixed.remove(f)
                        until = minute + f[1]
                        events.append([minute, until, f[2], f[3], job["wo"]])
                        continue
                    if r.random() < 1.6 / SHIFT_PLANNED_MIN:
                        code, text = r.choice(SMALL_STOPS)
                        until = min(SHIFT_PLANNED_MIN, minute + r.randint(6, 38))
                        events.append([minute, until, code, text, job["wo"]])
                        continue
                    seg = segs.setdefault(job["wo"], [job, 0.0])
                    rate = r.uniform(0.84, 0.95) / job["part"][6]
                    seg[1] += rate
                    job["done"] += rate
                    if job["done"] >= job["qty"]:
                        job["done"] = job["qty"]
                        job["closed"] = d
                        job = new_job(asset, wc, d, r)
                        dur = int(job["part"][5] * 60) + r.randint(-10, 15)
                        until = min(SHIFT_PLANNED_MIN, minute + 1 + dur)
                        events.append([minute + 1, until, "CHG", f"Changeover to {job['part'][0]}", job["wo"]])
                rows = []
                for wo, (j, made_f) in segs.items():
                    made = int(made_f)
                    pct = 0.12 if (asset == "LATHE-TX2-01" and d == SCRAP_SPIKE) else r.choice([0.0, 0.005, 0.01, 0.015, 0.03])
                    scrap = int(made * pct)
                    j["made"][d] = j["made"].get(d, 0) + made
                    ctr["lot"] += 1
                    plan.append([d, shift, asset, wo, j["part"][0], int(made_f / 0.89 * 0.95) if made else 0, "each"])
                    out.append([d, shift, asset, wo, j["part"][0], f"L{ctr['lot']}", made - scrap, scrap, "each",
                                iso(ts(d, 8 * 60 + 30, shift))])
                    rows.append({"d": d, "shift": shift, "asset": asset, "job": wo, "part": j["part"][0],
                                 "good": made - scrap, "scrap": scrap, "down": 0, "stops": ""})
                for st, en, code, text, ev_wo in events:
                    ctr["ev"] += 1
                    down.append([f"DT-{ctr['ev']:05d}", asset, iso(ts(d, st, shift)), iso(ts(d, en, shift)),
                                 en - st, code, text, ev_wo, shift, d])
                if not rows:
                    rows.append({"d": d, "shift": shift, "asset": asset, "job": job["wo"], "part": job["part"][0],
                                 "good": 0, "scrap": 0, "down": 0, "stops": ""})
                total = sum(e[1] - e[0] for e in events)
                chg = sum(e[1] - e[0] for e in events if e[2] == "CHG")
                last = rows[-1]
                last["down_all"] = total
                last["down_chg"] = chg
                last["stops_all"] = "; ".join(f"{e[2]} {e[1] - e[0]}m" for e in events)
                last["stops_no_chg"] = "; ".join(f"{e[2]} {e[1] - e[0]}m" for e in events if e[2] != "CHG")
                log.extend(rows)
            current[asset] = job
    return {"jobs": jobs, "out": out, "plan": plan, "down": down, "log": log}


# ---------------------------------------------------------------- the supervisor log as of a date

def log_rows_through(sim, through):
    """The supervisor's rows as typed, with the dated conflicts applied."""
    rows, key_drift_done = [], False
    for x in sim["log"]:
        if x["d"] > through:
            continue
        r = dict(x)
        if "down_all" in r:
            excl = r["d"] >= CHG_EXCLUDED_FROM
            r["down"] = r["down_all"] - (r["down_chg"] if excl else 0)
            r["stops"] = r["stops_no_chg"] if excl else r["stops_all"]
        if r["d"] >= KEY_DRIFT_FROM and r["shift"] == "2":
            num = r["job"][3:]
            if not key_drift_done:
                num = num[:-2] + num[-1] + num[-2]   # the transposed one
                key_drift_done = True
                r["orphan"] = True
            r["job"] = num
        rows.append(r)
    return rows


def build_log_workbook(rows, through, faults):
    """Shift Log + Summary + Notes, formulas included. Records the cells of each planted fault."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Shift Log"
    ws.append(["Date", "Shift", "Machine", "Job", "Part", "Good", "Scrap", "Downtime (min)", "Stops", "Total"])
    for c in ws[1]:
        c.font = Font(bold=True)
    first_row_of = {}
    for i, r in enumerate(rows, start=2):
        ws.append([r["d"], r["shift"], r["asset"], r["job"], r["part"], r["good"], r["scrap"], r["down"],
                   r["stops"], f"=F{i}+G{i}"])
        first_row_of.setdefault((r["d"], r["shift"]), i)
    last = len(rows) + 1

    if through >= PATTERN_BREAK:
        rows_day = [i for i, r in enumerate(rows, start=2) if r["d"] == PATTERN_BREAK]
        a, b = rows_day[3], rows_day[7]
        ws[f"J{a}"] = f"=F{a}+G{a - 1}"
        ws[f"J{b}"] = f"=F{b}+G{b}-65"
        faults.append({"onset": PATTERN_BREAK, "kind": "pattern_break", "sheet": "Shift Log", "cell": f"J{a}",
                       "formula": f"=F{a}+G{a - 1}"})
        faults.append({"onset": PATTERN_BREAK, "kind": "hardcoded_constant", "sheet": "Shift Log", "cell": f"J{b}",
                       "formula": f"=F{b}+G{b}-65"})
    if through >= HIDDEN_ROW:
        hr = first_row_of[(HIDDEN_ROW, "1")]
        ws.row_dimensions[hr].hidden = True
        faults.append({"onset": HIDDEN_ROW, "kind": "hidden_row_in_range", "sheet": "Shift Log", "cell": f"A{hr}",
                       "detail": f"row {hr} is hidden and inside the totals' range"})

    sm = wb.create_sheet("Summary")
    sm["A1"] = "Northfork shift log summary"
    sm["A1"].font = Font(bold=True)
    sm["A2"] = "Planned minutes per shift"
    sm["B2"] = SHIFT_PLANNED_MIN
    sm.append([])
    sm.append(["Date", "Good", "Scrap", "Downtime (min)", "Planned (min)", "Availability"])
    for c in sm[4]:
        c.font = Font(bold=True)
    days = sorted({r["d"] for r in rows})
    row_of = {}
    for i, d in enumerate(days, start=5):
        row_of[d] = i
        scrap_ref = "#REF!" if through >= DELETED_COLUMN else "'Shift Log'!$G:$G"
        sm.append([d, f"=SUMIFS('Shift Log'!$F:$F,'Shift Log'!$A:$A,$A{i})",
                   f"=SUMIFS({scrap_ref},'Shift Log'!$A:$A,$A{i})",
                   f"=SUMIFS('Shift Log'!$H:$H,'Shift Log'!$A:$A,$A{i})",
                   f"=12*PlannedMin", f"=(E{i}-D{i})/E{i}"])
    end = 5 + len(days)
    t = end + 1
    good_end = last - 2 if through >= SUM_SHORT else last
    sm[f"A{t}"] = "Total good (all log rows)"
    sm[f"B{t}"] = f"=SUM('Shift Log'!F2:F{good_end})"
    sm[f"A{t + 1}"] = "Total scrap (all log rows)"
    sm[f"B{t + 1}"] = "=SUM(#REF!)" if through >= DELETED_COLUMN else f"=SUM('Shift Log'!G2:G{last})"
    sm[f"A{t + 2}"] = "Total downtime (all log rows)"
    sm[f"B{t + 2}"] = f"=SUM('Shift Log'!H2:H{last})"
    wb.defined_names["PlannedMin"] = DefinedName("PlannedMin", attr_text="Summary!$B$2")
    wb.defined_names["ScrapCol"] = DefinedName(
        "ScrapCol", attr_text="'Shift Log'!#REF!" if through >= DELETED_COLUMN else "'Shift Log'!$G:$G")

    if through >= OVERWRITTEN:
        rr = row_of[OVERWRITTEN]
        typed = sum(r["good"] for r in rows if r["d"] == OVERWRITTEN)
        sm[f"B{rr}"] = typed
        faults.append({"onset": OVERWRITTEN, "kind": "overwritten_formula", "sheet": "Summary", "cell": f"B{rr}",
                       "detail": f"typed {typed} where the column has SUMIFS formulas"})
    if through >= DELETED_COLUMN:
        for d, rr in row_of.items():  # a day's row added after the deletion is broken from its own day
            faults.append({"onset": max(DELETED_COLUMN, d), "kind": "ref_error", "sheet": "Summary", "cell": f"C{rr}",
                           "formula": sm[f"C{rr}"].value})
        faults.append({"onset": DELETED_COLUMN, "kind": "ref_error", "sheet": "Summary", "cell": f"B{t + 1}",
                       "formula": "=SUM(#REF!)"})
        faults.append({"onset": DELETED_COLUMN, "kind": "broken_defined_name", "sheet": None, "cell": "ScrapCol",
                       "formula": "'Shift Log'!#REF!"})
    if through >= SUM_SHORT:
        faults.append({"onset": SUM_SHORT, "kind": "range_short", "sheet": "Summary", "cell": f"B{t}",
                       "formula": f"=SUM('Shift Log'!F2:F{good_end})", "detail": f"data runs to row {last}"})
    if through >= WRONG_PERIOD:
        oc = wb.create_sheet("Oct")
        oc.append(["Date", "Shift", "Machine", "Job", "Part", "Good", "Scrap"])
        rr = rng_for("oct")
        for k in range(30):
            oc.append([date(2026, 10, 1 + k % 30), "1", ASSETS[k % 6][0], f"WO-{88100 + k}", "PN-2003-A",
                       rr.randint(40, 140), rr.randint(0, 3)])
        x = t + 4
        sm[f"A{x}"] = "Month to date good"
        sm[f"B{x}"] = "=SUM(Oct!F2:F31)"
        sm[f"A{x + 1}"] = "Budget good (plan)"
        sm[f"B{x + 1}"] = "='[Budget_2026.xlsx]Targets'!B4"
        sm[f"A{x + 2}"] = "Adjusted downtime"
        sm[f"B{x + 2}"] = f"=B{x + 3}+D5"
        sm[f"A{x + 3}"] = "Downtime check"
        sm[f"B{x + 3}"] = f"=B{x + 2}-1"
        faults += [
            {"onset": WRONG_PERIOD, "kind": "wrong_period_sheet", "sheet": "Summary", "cell": f"B{x}",
             "formula": "=SUM(Oct!F2:F31)", "detail": "November summary built from the Oct sheet"},
            {"onset": WRONG_PERIOD, "kind": "external_link", "sheet": "Summary", "cell": f"B{x + 1}",
             "formula": "='[Budget_2026.xlsx]Targets'!B4"},
            {"onset": WRONG_PERIOD, "kind": "circular_reference", "sheet": "Summary", "cell": f"B{x + 2}",
             "detail": f"B{x + 2} -> B{x + 3} -> B{x + 2}"},
        ]
    if through >= DIV_BLANK:
        rr = row_of[DIV_BLANK]
        sm[f"E{rr}"] = None
        sm[f"G{rr}"] = f"=VLOOKUP(A{rr},Notes!A1:B6,4,FALSE)"
        sm["G4"] = "Shift lead"
        faults += [
            {"onset": DIV_BLANK, "kind": "div_by_blank", "sheet": "Summary", "cell": f"F{rr}",
             "formula": f"=(E{rr}-D{rr})/E{rr}", "will_show": "#DIV/0!"},
            {"onset": DIV_BLANK, "kind": "lookup_out_of_range", "sheet": "Summary", "cell": f"G{rr}",
             "formula": f"=VLOOKUP(A{rr},Notes!A1:B6,4,FALSE)", "will_show": "#REF!"},
        ]

    nt = wb.create_sheet("Notes")
    nt.append(["Northfork Precision: supervisor shift log"])
    nt.append(["Kept by the 1st shift lead. Filled in from the MES screen at end of shift."])
    nt.append(["Downtime (min) = all stops logged on the MES screen for the shift."])
    if through >= CHG_EXCLUDED_FROM:
        nt.append(["12 Nov: from today Downtime (min) leaves out changeovers. They're planned, not lost time."])
    if through >= KEY_DRIFT_FROM:
        nt.append(["13 Nov: 2nd shift enters job numbers without the WO- prefix (new tablet form)."])
    return wb


def save_excel_style(src, dst, errors):
    """Copy an openpyxl workbook and add the cached error values Excel would store.
    errors: {(sheet_xml_name, cell): '#DIV/0!'}"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            for (sheet_file, cell), err in errors.items():
                if item.filename == f"xl/worksheets/{sheet_file}":
                    text = data.decode("utf-8")
                    pat = re.compile(r'<c r="%s"([^>]*)>(<f>.*?</f>)(<v>.*?</v>|<v\s*/>)?</c>' % cell)
                    text = pat.sub(lambda m: '<c r="' + cell + '"' + re.sub(r' t="[^"]*"', "", m.group(1)) + ' t="e">'
                                   + m.group(2) + f"<v>{err}</v></c>", text, count=1)
                    data = text.encode("utf-8")
            zout.writestr(item, data)


# ---------------------------------------------------------------- other sources

def maintenance_tracker(path):
    """An unrelated workbook with its own errors, present from day 1."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Assets"
    ws.append(["Asset", "Run hours (YTD)", "Failures (YTD)", "MTBF (h)", "Vendor"])
    data = [("MILL-VX2-01", 3120, 4), ("MILL-VX2-02", 3010, 6), ("MILL-HX4-01", 2890, None),
            ("LATHE-TX2-01", 3305, 3), ("LATHE-TX2-02", 3180, 5), ("LATHE-CX1-01", 2975, 2)]
    for i, (a, h, f) in enumerate(data, start=2):
        ws.append([a, h, f, f"=B{i}/C{i}", f"=VLOOKUP(A{i},Vendors!A:B,2,FALSE)"])
    wb.save(path)
    return [{"onset": DAYS[0], "kind": "div_by_blank", "workbook": path.name, "sheet": "Assets", "cell": "D4",
             "formula": "=B4/C4", "will_show": "#DIV/0!"}] + [
            {"onset": DAYS[0], "kind": "missing_sheet_reference", "workbook": path.name, "sheet": "Assets",
             "cell": f"E{i}", "formula": f"=VLOOKUP(A{i},Vendors!A:B,2,FALSE)", "will_show": "#REF!"}
            for i in range(2, 8)]


def weekly_figures(sim, start, end, exclude_chg):
    figs = {}
    for a, *_ in ASSETS:
        good = sum(o[6] for o in sim["out"] if o[2] == a and start <= o[0] <= end)
        scrap = sum(o[7] for o in sim["out"] if o[2] == a and start <= o[0] <= end)
        dt = sum(e[4] for e in sim["down"] if e[1] == a and start <= e[9] <= end
                 and not (exclude_chg and e[5] == "CHG"))
        days = len([d for d in DAYS if start <= d <= end])
        planned = days * 2 * SHIFT_PLANNED_MIN
        figs[a] = {"good": good, "scrap": scrap, "downtime": dt, "planned": planned,
                   "availability": round(100 * (planned - dt) / planned, 1)}
    return figs


def weekly_pdf(path, sim, start, end, exclude_chg):
    figs = weekly_figures(sim, start, end, exclude_chg)
    p = pdfmini.Page()
    p.text(54, 740, "Northfork Precision", 16, bold=True)
    p.text(54, 720, f"Weekly OEE pack: {start.strftime('%-d %b')} to {end.strftime('%-d %b %Y')}", 12)
    p.text(54, 704, "Prepared by the plant office from the supervisor shift log and MES.", 9, gray=0.4)
    y = 670
    cols = [("Machine", 54, False), ("Good", 250, True), ("Scrap", 320, True), ("Downtime (min)", 430, True),
            ("Availability", 540, True)]
    p.rect(50, y - 6, 512, 20)
    for label, x, right in cols:
        p.text(x, y, label, 9, bold=True, right=right)
    y -= 22
    tot = {"good": 0, "scrap": 0, "downtime": 0, "planned": 0}
    for a, f in figs.items():
        p.text(54, y, a, 10)
        p.text(250, y, f"{f['good']:,}", 10, right=True)
        p.text(320, y, f"{f['scrap']:,}", 10, right=True)
        p.text(430, y, f"{f['downtime']:,}", 10, right=True)
        p.text(540, y, f"{f['availability']:.1f}%", 10, right=True)
        for k in tot:
            tot[k] += f[k]
        y -= 18
    p.line(50, y + 8, 562, y + 8)
    y -= 8
    avail = round(100 * (tot["planned"] - tot["downtime"]) / tot["planned"], 1)
    p.text(54, y, "Plant total", 10, bold=True)
    p.text(250, y, f"{tot['good']:,}", 10, bold=True, right=True)
    p.text(320, y, f"{tot['scrap']:,}", 10, bold=True, right=True)
    p.text(430, y, f"{tot['downtime']:,}", 10, bold=True, right=True)
    p.text(540, y, f"{avail:.1f}%", 10, bold=True, right=True)
    note = ("Downtime excludes changeovers, which are planned. Availability = (planned minutes - downtime) / planned "
            "minutes, 480 planned minutes per shift." if exclude_chg else
            "Downtime = all MES stops, changeovers and planned maintenance included. Availability = (planned minutes "
            "- downtime) / planned minutes, 480 planned minutes per shift.")
    p.text(54, y - 40, "Notes", 9, bold=True)
    words, line, yy = note.split(), "", y - 56
    for w in words:
        if pdfmini.width(line + " " + w, 9) > 500:
            p.text(54, yy, line.strip(), 9, gray=0.3)
            yy -= 12
            line = ""
        line += " " + w
    p.text(54, yy, line.strip(), 9, gray=0.3)
    path.write_bytes(pdfmini.pdf([p]))
    return {"week": f"{start}..{end}", "downtime_total": tot["downtime"], "availability": avail, "note": note}


def linked_summary(path, sim, through):
    """A values-only daily summary that someone publishes to SharePoint; it stops after LINK_LAST_REFRESH."""
    upto = min(through, LINK_LAST_REFRESH)
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily"
    ws.append(["Date", "Good", "Scrap", "Downtime (min)"])
    for d in [d for d in DAYS if d <= upto]:
        ws.append([d, sum(o[6] for o in sim["out"] if o[0] == d), sum(o[7] for o in sim["out"] if o[0] == d),
                   sum(e[4] for e in sim["down"] if e[9] == d)])
    ws["F1"] = "As of"
    ws["G1"] = upto
    wb.save(path)
    stamp = time.mktime(datetime(upto.year, upto.month, upto.day, 23, 30).timetuple())
    os.utime(path, (stamp, stamp))
    return upto


# ---------------------------------------------------------------- write the plant as of a date

def build(through, out_dir):
    if through not in DAYS:
        raise SystemExit(f"{through} isn't a production day. Days: {', '.join(map(str, DAYS))}")
    sim = simulate()
    if out_dir.exists():
        shutil.rmtree(out_dir)
    key = {"through": through.isoformat(), "day": DAYS.index(through) + 1, "findings": []}
    F = key["findings"]

    # ERP as of the date
    erp = out_dir / "erp"
    jobs = [j for j in sim["jobs"] if j["start"] <= through]
    write_csv(erp / "items.csv", ["part_number", "description", "revision", "uom", "standard_cost"],
              [[p[0], p[1], p[2], "each", f"{p[3]:.2f}"] for p in PARTS])
    write_csv(erp / "sales_orders.csv",
              ["sales_order", "customer", "part_number", "quantity", "uom", "promise_date", "status"],
              [[j["so"], j["customer"], j["part"][0], j["qty"], "each", j["promise"].isoformat(), "Open"] for j in jobs])
    wo = []
    for j in jobs:
        done = j.get("pre", 0) + sum(v for d, v in j["made"].items() if d <= through)
        closed = j["closed"] is not None and j["closed"] <= through
        wo.append([j["wo"], j["part"][0], j["part"][2], j["qty"], done, "each",
                   "Closed" if closed else ("In Process" if done else "Released"), j["start"].isoformat(),
                   j["promise"].isoformat(), j["so"]])
    write_csv(erp / "production_work_orders.csv",
              ["production_wo", "part_number", "revision", "qty_ordered", "qty_completed", "uom", "status",
               "start_date", "due_date", "sales_order"], wo)

    # MES, planning, CMMS: one file per day, as they land
    ex = out_dir / "exports"
    for d in [d for d in DAYS if d <= through]:
        write_csv(ex / f"mes_production_output_{d}.csv",
                  ["production_day", "shift", "asset_id", "production_wo", "part_number", "lot", "good_qty",
                   "scrap_qty", "uom", "reported_at"], [[o[0].isoformat()] + o[1:] for o in sim["out"] if o[0] == d])
        write_csv(ex / f"mes_downtime_events_{d}.csv",
                  ["event_id", "asset_id", "started_at", "ended_at", "duration_min", "reason_code", "reason_text",
                   "production_wo", "shift", "production_day"],
                  [e[:9] + [e[9].isoformat()] for e in sim["down"] if e[9] == d])
        write_csv(ex / f"planning_schedule_{d}.csv",
                  ["production_day", "shift", "asset_id", "production_wo", "part_number", "planned_qty", "uom"],
                  [[p[0].isoformat()] + p[1:] for p in sim["plan"] if p[0] == d])
    write_csv(ex / "mes_reason_codes.csv", ["reason_code", "description", "category", "counts_as_downtime"],
              [[c, dsc, cat, "Y"] for c, dsc, cat in REASONS])
    write_csv(ex / "cmms_assets.csv", ["asset_id", "name", "make", "model", "cell", "work_center", "status"],
              [[a[0], a[1], a[2], a[3], a[4], a[5], "Active"] for a in ASSETS])
    mwo = [["MWO-3410", "MILL-VX2-02", "PM", "2026-11-12T06:32-07:00", "2026-11-12T07:32-07:00",
            "Quarterly PM: way lube, coolant concentration, filters", "Closed"],
           ["MWO-3424", "LATHE-TX2-02", "Corrective", "2026-11-16T16:05-07:00", "2026-11-16T17:15-07:00",
            "Spindle drive alarm 2041. Reseated drive connector, replaced cooling fan.", "Closed"]]
    write_csv(ex / "cmms_maintenance_work_orders.csv",
              ["maintenance_wo", "asset_id", "type", "opened_at", "closed_at", "description", "status"],
              [m for m in mwo if m[3][:10] <= through.isoformat()])

    # The supervisor workbook
    sp = out_dir / "spreadsheets"
    sp.mkdir(parents=True)
    rows = log_rows_through(sim, through)
    faults = []
    wb = build_log_workbook(rows, through, faults)
    log_path = sp / "supervisor_shift_log_Nov.xlsx"
    wb.save(log_path)
    for f in faults:
        f["workbook"] = log_path.name
    F.extend({"category": "excel", **f} for f in faults)
    F.extend({"category": "excel", **f} for f in maintenance_tracker(sp / "maintenance_tracker.xlsx"))

    # Excel-saved copy of the supervisor workbook (outside the plant folder)
    saved_dir = out_dir.parent / "excel_saved"
    saved_dir.mkdir(exist_ok=True)
    sheet_file = {"Shift Log": "sheet1.xml", "Summary": "sheet2.xml"}
    cached = {(sheet_file[f["sheet"]], f["cell"]): f.get("will_show", "#REF!")
              for f in faults if f["kind"] in ("ref_error", "div_by_blank", "lookup_out_of_range")}
    save_excel_style(log_path, saved_dir / log_path.name, cached)

    # Definition conflicts and data findings (the answer key)
    if through >= CHG_EXCLUDED_FROM:
        chg = sum(e[4] for e in sim["down"] if e[5] == "CHG" and CHG_EXCLUDED_FROM <= e[9] <= through)
        F.append({"category": "definition", "kind": "definition_change", "onset": CHG_EXCLUDED_FROM,
                  "metric": "downtime", "sources": ["MES downtime events", "supervisor shift log"],
                  "detail": "log leaves out changeovers (CHG) from this day; Notes line added the same day",
                  "gap_min": chg, "stage": "prod::Downtime"})
    if through >= KEY_DRIFT_FROM:
        orphan = next(r["job"] for r in rows if r.get("orphan"))
        n = sum(1 for r in rows if r["d"] >= KEY_DRIFT_FROM and r["shift"] == "2")
        F.append({"category": "keys", "kind": "key_format_drift", "onset": KEY_DRIFT_FROM, "rows": n,
                  "detail": "2nd-shift job numbers without WO-", "stage": "prod::Work Order"})
        F.append({"category": "keys", "kind": "orphan_key", "onset": KEY_DRIFT_FROM, "value": orphan,
                  "detail": "transposed job number, matches no work order", "stage": "prod::Work Order"})
    if through >= SCRAP_SPIKE:
        F.append({"category": "anomaly", "kind": "spike", "onset": SCRAP_SPIKE, "metric": "scrap",
                  "asset": "LATHE-TX2-01", "detail": "scrap about 12% of output vs a normal 0-3%",
                  "stage": "prod::Quality"})

    # Reports: the weekly PDF pack and the linked summary
    rep = out_dir / "reports"
    rep.mkdir()
    if through >= WEEK1_PUBLISHED:
        weekly_pdf(rep / "Northfork_weekly_OEE_pack_2026-11-13.pdf", sim, D["day1"], D["day5"], False)
    if through >= WEEK2_PUBLISHED:
        w2 = weekly_pdf(rep / "Northfork_weekly_OEE_pack_2026-11-20.pdf", sim, D["day6"], D["day10"], True)
        mes_w2 = sum(e[4] for e in sim["down"] if D["day6"] <= e[9] <= D["day10"])
        F.append({"category": "report", "kind": "report_vs_source", "onset": WEEK2_PUBLISHED,
                  "report": "Northfork_weekly_OEE_pack_2026-11-20.pdf", "metric": "downtime",
                  "report_value": w2["downtime_total"], "mes_value": mes_w2,
                  "detail": "report follows the shift log (no changeovers); MES counts them", "stage": "prod::Downtime"})
        F.append({"category": "report", "kind": "report_definition_drift", "onset": WEEK2_PUBLISHED,
                  "report": "Northfork_weekly_OEE_pack", "detail": "footnote changed to 'excludes changeovers'",
                  "stage": "prod::Downtime"})
    linked = out_dir.parent / "linked"
    linked.mkdir(exist_ok=True)
    upto = linked_summary(linked / "daily_production_summary.xlsx", sim, through)
    if through > LINK_LAST_REFRESH:
        F.append({"category": "report", "kind": "stale_link", "onset": DAYS[DAYS.index(LINK_LAST_REFRESH) + 1],
                  "report": "daily_production_summary.xlsx", "detail": f"not refreshed since {upto}"})

    key_path = out_dir.parent / f"answer_key_{through}.json"
    key_path.write_text(json.dumps(key, indent=2, default=str))
    print(f"Wrote {out_dir} as of {through} (day {key['day']}): {len(F)} expected findings -> {key_path.name}")
    return key


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--through", default=DAYS[-1].isoformat())
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent.parent / "rolling" / "plant"))
    a = ap.parse_args()
    build(date.fromisoformat(a.through), Path(a.out))


if __name__ == "__main__":
    main()
