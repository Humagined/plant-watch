"""Detectors. Each takes plain data and returns findings (dicts); none writes anything.

Finding fields shared by every kind: kind, category, metric (or None), stage (or None),
onset (date the problem starts in the data, when known), title, detail, evidence, question.
"""

import json
import re
import statistics
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import pdftext
import sources as S

DATA = Path(__file__).resolve().parent.parent / "data"
METRICS = json.loads((DATA / "metrics.json").read_text(encoding="utf-8"))["metrics"]
TOL = 0.5


def stage_of(metric):
    return METRICS.get(metric, {}).get("stage")


def src_label(cfg, sid):
    s = next((x for x in cfg["sources"] if x["id"] == sid), {})
    return s.get("label") or sid


# ---------------------------------------------------------------- shared shapes

def cells_by_source(facts, metric):
    """{source: {(date, shift, entity): value}} and {source: {cell: {category: value}}}."""
    vals, cats = defaultdict(lambda: defaultdict(float)), defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for f in facts:
        v = f.get(metric)
        if v is None:
            continue
        cell = (f["date"], f["shift"], f["entity"])
        vals[f["source"]][cell] += v
        if f.get("category"):
            cats[f["source"]][cell][f["category"]] += v
    return vals, cats


def record_source(cfg, metric, dictionary, candidates):
    entry = (dictionary or {}).get(metric) or {}
    if entry.get("record") in candidates:
        return entry["record"]
    for s in cfg["sources"]:
        if s["id"] in candidates and metric in s.get("record_for", []):
            return s["id"]
    ev = [s["id"] for s in cfg["sources"] if s["id"] in candidates and s.get("kind") == "events"]
    return ev[0] if ev else sorted(candidates)[0]


def dated_notes(notes, metric, year):
    """Notes lines that start with a date and mention the metric: [(date, source, line)]."""
    syns = METRICS.get(metric, {}).get("synonyms", [metric])
    out = []
    for sid, lines in notes.items():
        for line in lines:
            m = re.match(r"\s*(\d{1,2})\s+([A-Za-z]{3})[a-z]*\.?:?\s", line)
            if not m or not any(s in line.lower() for s in syns):
                continue
            d = S.to_date(f"{m.group(1)} {m.group(2).title()} {year}")
            if d:
                out.append((d, sid, line.strip()))
    return out


# ---------------------------------------------------------------- definitions: source vs source vs dictionary

def definitions(cfg, facts, notes, dictionary):
    found = []
    metrics = sorted({m for f in facts for m in METRICS if f.get(m) is not None})
    year = max((f["date"].year for f in facts), default=date.today().year)
    for metric in metrics:
        vals, cats = cells_by_source(facts, metric)
        if len(vals) < 2:
            continue
        rec = record_source(cfg, metric, dictionary, set(vals))
        entry = (dictionary or {}).get(metric) or {}
        for other in sorted(set(vals) - {rec}):
            common_days = {c[0] for c in vals[rec]} & {c[0] for c in vals[other]}
            cells = sorted({c for c in list(vals[rec]) + list(vals[other]) if c[0] in common_days})
            explained, unexplained = defaultdict(list), []
            left_out = {}   # cell -> the record's category this source leaves out there ("" for none)
            for c in cells:
                a, b = vals[rec].get(c, 0.0), vals[other].get(c, 0.0)
                diff = a - b
                if abs(diff) < TOL:
                    left_out[c] = ""
                    continue
                code = next((k for k, v in cats[rec].get(c, {}).items() if abs(v - diff) < TOL), None)
                if code is None:
                    code = next((k for k, v in cats[other].get(c, {}).items() if abs(v + diff) < TOL), None)
                    if code:
                        code = "+" + code
                if code:
                    explained[code].append((c, diff))
                    left_out[c] = "" if code.startswith("+") else code
                else:
                    unexplained.append((c, a, b))
            confirmed = entry.get("status") == "confirmed"
            exclude = set(entry.get("exclude_categories", [])) if confirmed else set()
            # What the definition leaves out but this source still counts: wherever the source matches the
            # record (or differs only by another code), it includes every category the record has there.
            # Once a source starts leaving a code out, the days before are the change finding's story, not this one's.
            still_counts = {}
            for x in sorted(exclude):
                since = min((c[0] for c, _ in explained.get(x, [])), default=None)
                hits = [(c, cats[rec][c][x]) for c, k in sorted(left_out.items())
                        if k != x and cats[rec].get(c, {}).get(x, 0) >= TOL and (since is None or c[0] >= since)]
                if hits:
                    still_counts[x] = hits
            for code, hits in explained.items():
                onset = min(c[0] for c, _ in hits)
                gap = sum(d for _, d in hits)
                base = sum(v for c, v in vals[rec].items() if c[0] >= onset and c[0] in common_days)
                leaves_out = not code.startswith("+")
                cname = code.lstrip("+")
                changed = onset > min(common_days)
                agrees = False
                if confirmed:
                    counts_code = cname not in exclude
                    if leaves_out and counts_code:
                        verdict = (f"{src_label(cfg, other)} leaves out {cname} from {metric}, but your definition counts {cname}. "
                                   f"Its {metric} is understated by {gap:,.0f} {METRICS[metric]['unit']} since {S.day(onset)}.")
                    elif leaves_out and not counts_code:
                        agrees = True
                        verdict = (f"{src_label(cfg, other)} leaves out {cname}, as your definition does; {src_label(cfg, rec)} "
                                   f"counts {cname}, so any figure built straight from it needs {cname} taken out.")
                        also = [x for x in still_counts if x != cname]
                        if also:
                            verdict += (f" It doesn't fully match your definition: it still counts "
                                        f"{', '.join(also)} (a separate finding).")
                    else:
                        verdict = f"{src_label(cfg, other)} adds {cname} that {src_label(cfg, rec)} doesn't have."
                else:
                    verdict = (f"{src_label(cfg, rec)} counts {cname} in {metric}; {src_label(cfg, other)} "
                               f"{'leaves it out' if leaves_out else 'adds it'}. The gap is {gap:,.0f} {METRICS[metric]['unit']} "
                               f"since {S.day(onset)}.")
                ev_notes = [{"date": d, "source": sid, "line": ln} for d, sid, ln in dated_notes(notes, metric, year)
                            if sid == other and abs((d - onset).days) <= 3]
                found.append({
                    "kind": "definition_change" if changed else "definition_mismatch", "category": "definition",
                    "metric": metric, "stage": stage_of(metric), "onset": onset, "sources": [rec, other],
                    "code": cname, "record": rec, "gap": round(gap, 2), "base": round(base, 2),
                    "share": round(gap / base, 4) if base else None, "cells": len(hits),
                    "days": sorted({c[0] for c, _ in hits}),
                    "title": f"{METRICS[metric]['label']}: {src_label(cfg, other)} "
                             f"{'leaves out' if leaves_out else 'adds'} {cname}" + (f" from {S.day(onset)}" if changed else ""),
                    "detail": verdict, "evidence": {"notes": ev_notes, "examples": [
                        {"date": c[0], "shift": c[1], "entity": c[2], "gap": d} for c, d in hits[:5]]},
                    "question": METRICS[metric]["question"] if not confirmed else
                                (f"Which figures are built straight from {src_label(cfg, rec)}, and do they take {cname} out?"
                                 if agrees else
                                 f"Should {src_label(cfg, other)} go back to counting {cname}, or is the definition wrong?"
                                 if leaves_out else
                                 f"Where do {src_label(cfg, other)}'s extra {cname} {METRICS[metric]['unit']} come from?"),
                    "id_parts": [metric, rec, other, cname]})
            for x, hits in still_counts.items():
                onset = min(c[0] for c, _ in hits)
                over = sum(v for _, v in hits)
                found.append({
                    "kind": "definition_mismatch", "category": "definition", "metric": metric, "stage": stage_of(metric),
                    "onset": onset, "sources": [rec, other], "code": x, "record": rec, "gap": round(-over, 2),
                    "cells": len(hits), "days": sorted({c[0] for c, _ in hits}),
                    "title": f"{METRICS[metric]['label']}: {src_label(cfg, other)} counts {x}, which your definition leaves out",
                    "detail": (f"Your definition leaves out {x}, but {src_label(cfg, other)} counts it: its {metric} includes "
                               f"{over:,.0f} {METRICS[metric]['unit']} of {x} across {len(hits)} machine-shift(s) "
                               f"since {S.day(onset)}."),
                    "evidence": {"notes": [], "examples": [
                        {"date": c[0], "shift": c[1], "entity": c[2], "gap": -v} for c, v in hits[:5]]},
                    "question": f"Should {src_label(cfg, other)} stop counting {x}, or should your definition include it?",
                    "id_parts": [metric, rec, other, x, "counts_excluded"]})
            if unexplained:
                onset = min(c[0] for c, _, _ in unexplained)
                found.append({
                    "kind": "unexplained_gap", "category": "definition", "metric": metric, "stage": stage_of(metric),
                    "onset": onset, "sources": [rec, other], "record": rec, "cells": len(unexplained),
                    "gap": round(sum(a - b for _, a, b in unexplained), 2),
                    "title": f"{METRICS[metric]['label']}: {src_label(cfg, rec)} and {src_label(cfg, other)} disagree",
                    "detail": f"{len(unexplained)} machine-shifts differ and no single reason code explains the gap.",
                    "evidence": {"examples": [{"date": c[0], "shift": c[1], "entity": c[2], rec: a, other: b}
                                              for c, a, b in unexplained[:8]]},
                    "question": METRICS[metric]["question"], "id_parts": [metric, rec, other, "unexplained"]})
    return found


# ---------------------------------------------------------------- keys

def keys(cfg, facts, notes):
    found = []
    year = max((f["date"].year for f in facts), default=date.today().year)
    for src in cfg["sources"]:
        if not (src.get("key") and src.get("key_master")):
            continue
        master = S.master_keys(cfg, src["key_master"])
        prefixes = [p for p, n in _common_prefixes(master)]
        rows = [f for f in facts if f["source"] == src["id"] and f.get("key")]
        drift, orphans = [], []
        for f in rows:
            k = f["key"]
            if k in master:
                continue
            fixed = next((p + k.strip().upper() for p in prefixes if p + k.strip().upper() in master), None)
            if fixed:
                drift.append((f, fixed))
            else:
                orphans.append(f)
        stage = next((s.get("stage") for s in cfg["sources"] if s["id"] == src["key_master"]), None) or stage_of("job_key")
        if drift:
            onset = min(f["date"] for f, _ in drift)
            shifts = sorted({f["shift"] for f, _ in drift})
            span = [f for f in rows if f["date"] >= onset]
            ev = [{"date": d, "source": sid, "line": ln} for d, sid, ln in dated_notes(notes, "job_key", year)
                  if sid == src["id"]]
            found.append({"kind": "key_format_drift", "category": "keys", "metric": "job_key", "stage": stage,
                          "onset": onset, "sources": [src["id"], src["key_master"]], "rows": len(drift),
                          "share": round(len(drift) / len(span), 4) if span else None,
                          "title": f"{src_label(cfg, src['id'])}: job numbers typed without '{prefixes[0]}'"
                                   f" (shift {', '.join(shifts)}) from {S.day(onset)}",
                          "detail": f"{len(drift)} rows don't match a work order as typed but do once '{prefixes[0]}' is added.",
                          "evidence": {"notes": ev, "examples": [{"file": f["file"], "row": f["row"], "typed": f["key"],
                                                                  "matches": fx} for f, fx in drift[:5]]},
                          "question": METRICS["job_key"]["question"], "id_parts": [src["id"], "format"]})
        for f in orphans:
            hint = _transposition(f["key"], master, prefixes)
            found.append({"kind": "orphan_key", "category": "keys", "metric": "job_key", "stage": stage,
                          "onset": f["date"], "sources": [src["id"], src["key_master"]], "value": f["key"],
                          "title": f"{src_label(cfg, src['id'])}: job {f['key']} matches no work order",
                          "detail": f"{f['file']} row {f['row']} ({S.day(f['date'])}, shift {f['shift']}, {f['entity']})"
                                    + (f". Likely a typo of {hint}: two digits swapped." if hint else "."),
                          "evidence": {"file": f["file"], "row": f["row"], "likely": hint},
                          "question": "Which job was this? Ask the shift lead before anyone corrects it.",
                          "id_parts": [src["id"], "orphan", f["key"]]})
    return found


def _common_prefixes(keys):
    counts = defaultdict(int)
    for k in keys:
        m = re.match(r"([A-Za-z]+[-_ ]?)", k)
        if m:
            counts[m.group(1).upper()] += 1
    return sorted(counts.items(), key=lambda kv: -kv[1])


def _transposition(k, master, prefixes):
    digits = re.sub(r"\D", "", k)
    for i in range(len(digits) - 1):
        sw = digits[:i] + digits[i + 1] + digits[i] + digits[i + 2:]
        for p in prefixes:
            if p + sw in master:
                return p + sw
    return None


# ---------------------------------------------------------------- anomalies over time

def anomalies(cfg, facts, dictionary):
    found = []
    by = defaultdict(lambda: defaultdict(lambda: {"good": 0.0, "scrap": 0.0, "downtime": 0.0, "seen": set()}))
    sources_with = defaultdict(set)
    for f in facts:
        for m in ("good", "scrap", "downtime"):
            if f.get(m) is not None:
                sources_with[m].add(f["source"])
    rec_out = record_source(cfg, "good", dictionary, sources_with["good"]) if sources_with["good"] else None
    for f in facts:
        if f["source"] != rec_out:
            continue
        d = by[f["entity"]][f["date"]]
        d["good"] += f.get("good") or 0
        d["scrap"] += f.get("scrap") or 0
    for entity, days in by.items():
        order = sorted(days)
        for i, d in enumerate(order):
            hist = order[max(0, i - 10):i]
            if len(hist) < 3:
                continue
            tot = days[d]["good"] + days[d]["scrap"]
            rate = days[d]["scrap"] / tot if tot else 0
            hrates = [days[h]["scrap"] / (days[h]["good"] + days[h]["scrap"]) for h in hist
                      if days[h]["good"] + days[h]["scrap"]]
            if hrates:
                p90 = sorted(hrates)[int(0.9 * (len(hrates) - 1))]
                if rate >= 0.05 and rate >= max(3 * p90, p90 + 0.03):
                    found.append({"kind": "spike", "category": "anomaly", "metric": "scrap", "stage": stage_of("scrap"),
                                  "onset": d, "entity": entity, "sources": [rec_out],
                                  "title": f"Scrap spike on {entity}, {S.day(d)}: {rate:.0%} of output",
                                  "detail": f"Scrap was {days[d]['scrap']:.0f} of {tot:.0f} ({rate:.1%}); the previous "
                                            f"{len(hrates)} days ran at or below {p90:.1%}.",
                                  "evidence": {"rate": round(rate, 4), "typical_max": round(p90, 4)},
                                  "question": "Is this a real quality problem on that machine, or a recording error?",
                                  "id_parts": ["spike", "scrap", entity, d.isoformat()]})
            hgood = [days[h]["good"] for h in hist]
            if days[d]["good"] == 0 and statistics.median(hgood) > 0:
                found.append({"kind": "zero_output", "category": "anomaly", "metric": "good", "stage": stage_of("good"),
                              "onset": d, "entity": entity, "sources": [rec_out],
                              "title": f"No output recorded on {entity}, {S.day(d)}",
                              "detail": f"Usually about {statistics.median(hgood):.0f} a day.", "evidence": {},
                              "question": "Was the machine down all day, or is the day's output missing?",
                              "id_parts": ["zero", entity, d.isoformat()]})
    # Days present in one source and missing from another
    days_by = defaultdict(set)
    for f in facts:
        days_by[f["source"]].add(f["date"])
    if days_by:
        alld = set().union(*days_by.values())
        for sid, ds in days_by.items():
            for d in sorted(alld - ds):
                if d > min(ds):
                    found.append({"kind": "missing_day", "category": "anomaly", "metric": None, "stage": None,
                                  "onset": d, "sources": [sid], "title": f"{src_label(cfg, sid)} has no rows for {S.day(d)}",
                                  "detail": "Other sources have data for that day.", "evidence": {},
                                  "question": "Was that day's file missed, or did the source not run?",
                                  "id_parts": ["missing", sid, d.isoformat()]})
    return found


# ---------------------------------------------------------------- reports (PDF / linked) vs the data

def _fig_period(ex, fig):
    """A figure's own date or period (a daily summary has one row per day), else the report's period."""
    if fig.get("date"):
        d = S.to_date(fig["date"])
        return d, d
    p = fig.get("period") or ex.get("period") or {}
    return S.to_date(p.get("start")), S.to_date(p.get("end"))


def verify_figure(path, fig, _books={}):
    """True / False if the figure is / isn't where the reading says; None if that can't be checked.
    PDF: the quote is on the cited page. Spreadsheet: the cited sheet!cell holds the value, or the quote is a row."""
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix == ".pdf":
        return pdftext.verify(p, fig.get("page", 1), fig.get("quote", ""))
    if suffix in (".xlsx", ".xlsm"):
        import xlsx_audit as xa
        key = (str(p), p.stat().st_mtime if p.exists() else None)
        if key not in _books:
            try:
                _books[key] = xa.Workbook(p)
            except Exception:
                return None
        wb = _books[key]
        sheet = fig.get("sheet") or (wb.sheet_names[0] if wb.sheet_names else None)
        if sheet not in wb.sheet_names:
            return False
        if fig.get("cell"):
            m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", fig["cell"].upper())
            cell = wb.value(sheet, xa.split_cell(f"{m.group(1)}{m.group(2)}")) if m else None
            v = S.num(cell["value"]) if cell else None
            return v is not None and abs(v - float(fig["value"])) < TOL
        if fig.get("quote"):
            rows = defaultdict(list)
            for (c, r), cell in sorted(wb.cells.get(sheet, {}).items(), key=lambda kv: (kv[0][1], kv[0][0])):
                if cell.get("value") not in (None, ""):
                    rows[r].append(str(cell["value"]))
            q = pdftext.norm(fig["quote"])
            return any(q in pdftext.norm(" ".join(vs)) for vs in rows.values())
        return None
    if suffix == ".csv" and fig.get("quote"):
        return pdftext.norm(fig["quote"]) in pdftext.norm(p.read_text(errors="replace"))
    return None


def reports(cfg, facts, extractions, dictionary):
    """extractions: dicts written by the model and confirmed against a template:
    {file, path, series, period: {start, end}, published, footnote,
     figures: [{metric, entity ('total' for the plant), value, unit, page, quote}]}
    A spreadsheet reading cites {sheet, cell} instead of {page}, and a figure may carry its own
    date (or period) when the report has one row per day."""
    found = []
    series = defaultdict(list)
    for ex in extractions:
        if not ex.get("period"):
            ds = [d for fig in ex["figures"] for d in _fig_period({}, fig) if d]
            ex["period"] = {"start": min(ds).isoformat(), "end": max(ds).isoformat()} if ds else {}
        series[ex.get("series", "")].append(ex)
        start, end = S.to_date(ex["period"].get("start")), S.to_date(ex["period"].get("end"))
        is_pdf = str(ex.get("path") or ex.get("file", "")).lower().endswith(".pdf")
        unverified = []
        for fig in ex["figures"]:
            ok = verify_figure(ex["path"], fig) if ex.get("path") else None
            fig["verified"] = ok
            if ok is False:
                unverified.append(fig)
        if unverified:
            found.append({"kind": "extraction_unverified", "category": "report", "metric": None, "stage": None,
                          "onset": S.to_date(ex.get("published")) or end, "report": ex["file"], "sources": [ex["file"]],
                          "title": f"{ex['file']}: {len(unverified)} figure(s) not found "
                                   + ("on the cited page" if is_pdf else "in the cited cell"),
                          "detail": "The reading of this report couldn't be confirmed against the "
                                    + ("PDF" if is_pdf else "file") + ", so these figures "
                                    "weren't compared. Re-read the report or confirm them by eye.",
                          "evidence": {"figures": unverified[:5]}, "question": None,
                          "id_parts": [ex["file"], "unverified"]})
        by_metric = defaultdict(list)
        for fig in ex["figures"]:
            if fig.get("verified") is not False and fig["metric"] in METRICS:
                by_metric[fig["metric"]].append(fig)
        for metric, figs in by_metric.items():
            vals, cats = cells_by_source(facts, metric)
            if not vals:
                continue
            rec = record_source(cfg, metric, dictionary, set(vals))
            entry = (dictionary or {}).get(metric) or {}
            exclude = set(entry.get("exclude_categories", [])) if entry.get("status") == "confirmed" else set()
            diffs = []
            for fig in figs:
                ent = fig.get("entity", "total")
                fs, fe = _fig_period(ex, fig)
                if not (fs and fe):
                    continue
                pick = [c for c in vals[rec] if fs <= c[0] <= fe and (ent == "total" or c[2] == ent)]
                code_tot = defaultdict(float)
                for c in pick:
                    for k, v in cats[rec].get(c, {}).items():
                        code_tot[k] += v
                # The figure the definition gives: the record, less whatever the definition leaves out.
                src_val = sum(vals[rec][c] for c in pick) - sum(code_tot[x] for x in exclude)
                diff = src_val - fig["value"]
                if abs(diff) < TOL:
                    continue
                code = next((k for k, v in code_tot.items() if k not in exclude and abs(v - diff) < TOL), None)
                if code is None and diff < 0:   # the report counts what the definition leaves out
                    present = sorted(x for x in exclude if code_tot[x] >= TOL)
                    one = next((x for x in present if abs(code_tot[x] + diff) < TOL), None)
                    if one or (present and abs(sum(code_tot[x] for x in present) + diff) < TOL):
                        code = "+" + "+".join([one] if one else present)
                diffs.append({"entity": ent, "start": fs, "end": fe, "report": fig["value"], "data": src_val, "gap": diff,
                              "explained_by": code, "page": fig.get("page"), "sheet": fig.get("sheet"),
                              "cell": fig.get("cell"), "quote": fig.get("quote")})
            if diffs:
                codes = {d["explained_by"] for d in diffs}
                code = codes.pop() if len(codes) == 1 else None
                if code is None and all(c and c.startswith("+") for c in codes | {d["explained_by"] for d in diffs}):
                    # every gap is something the definition leaves out, if not the same thing each day
                    code = "+" + "+".join(sorted({x for d in diffs for x in d["explained_by"].lstrip("+").split("+")}))
                tot = next((d for d in diffs if d["entity"] == "total"), None)
                counts = " and ".join(code.lstrip("+").split("+")) if code and code.startswith("+") else None
                basis = src_label(cfg, rec) + (" under your definition" if exclude else "")
                if tot:
                    span = (f"{S.day(tot['start'])}" if tot["start"] == tot["end"]
                            else f"{S.day(tot['start'])} to {S.day(tot['end'])}")
                    detail = (f"For {span} the report shows {tot['report']:,.0f} where {basis} has {tot['data']:,.0f}."
                              + (f" {len(diffs) - 1} other figure(s) differ too." if len(diffs) > 1 else ""))
                else:
                    detail = f"{len(diffs)} figures differ from {basis}."
                if counts:
                    detail += f" The gap is exactly the {counts} the report still counts, which your definition leaves out."
                elif code:
                    detail += f" The gap equals the {code} {METRICS[metric]['unit']} exactly."
                found.append({
                    "kind": "report_vs_source", "category": "report", "metric": metric, "stage": stage_of(metric),
                    "onset": S.to_date(ex.get("published")) or (min(d["start"] for d in diffs)
                                                                 if any(f.get("date") or f.get("period") for f in figs) else end),
                    "report": ex["file"], "sources": [ex["file"], rec],
                    "code": code, "gap": tot["gap"] if tot else sum(d["gap"] for d in diffs),
                    "title": f"{ex['file']}: {METRICS[metric]['label'].lower()} doesn't match {src_label(cfg, rec)}"
                             + (f" (it counts {counts}, which your definition leaves out)" if counts else
                                f" (it leaves out {code})" if code else ""),
                    "detail": detail,
                    "evidence": {"figures": diffs[:8], "footnote": ex.get("footnote")},
                    "question": (f"Should this report stop counting {counts}, or should your definition include "
                                 f"{'them' if ' and ' in counts else 'it'}?"
                                 if counts else METRICS[metric]["question"]),
                    "id_parts": [ex["file"], metric, "vs"]})
    for name, exs in series.items():
        exs.sort(key=lambda e: S.to_date(e["period"].get("start")) or date.min)
        for prev, cur in zip(exs, exs[1:]):
            if pdftext.norm(prev.get("footnote") or "") != pdftext.norm(cur.get("footnote") or ""):
                found.append({"kind": "report_definition_drift", "category": "report", "metric": "downtime"
                              if "downtime" in (cur.get("footnote") or "").lower() else None,
                              "stage": stage_of("downtime") if "downtime" in (cur.get("footnote") or "").lower() else None,
                              "onset": S.to_date(cur.get("published")) or S.to_date(cur["period"].get("end")),
                              "report": cur["file"], "sources": [prev["file"], cur["file"]],
                              "title": f"{name or cur['file']}: the definition note changed",
                              "detail": f"Before: \"{prev.get('footnote')}\" Now: \"{cur.get('footnote')}\"",
                              "evidence": {"before": prev.get("footnote"), "after": cur.get("footnote")},
                              "question": "Was this definition change agreed, and should earlier weeks be restated?",
                              "id_parts": [name, "footnote", cur["file"]]})
    return found


def links(cfg, link_states, days):
    """link_states: {source_id: {"last_modified": iso, "as_of": iso or None, "error": str or None}}
    days: the production days in the data, so a stale report's onset is the first day it missed."""
    latest = max(days) if days else None
    found = []
    for src in cfg["sources"]:
        if src.get("kind") != "link":
            continue
        st = (link_states or {}).get(src["id"]) or {}
        title = src.get("label") or src["id"]
        if st.get("error"):
            found.append({"kind": "link_unreachable", "category": "report", "metric": None, "stage": src.get("stage"),
                          "onset": latest, "report": title, "sources": [src["id"]],
                          "title": f"{title}: link couldn't be read", "detail": st["error"], "evidence": {},
                          "question": "Has the report moved or has access changed?", "id_parts": [src["id"], "unreachable"]})
            continue
        stamp = S.to_date(st.get("as_of")) or S.to_date(st.get("last_modified"))
        if stamp and latest and stamp < latest:
            found.append({"kind": "stale_link", "category": "report", "metric": None, "stage": src.get("stage"),
                          "onset": min(d for d in days if d > stamp), "report": title, "sources": [src["id"]],
                          "title": f"{title}: not refreshed since {S.day(stamp)}",
                          "detail": f"The exports run to {S.day(latest)}, but this report stops at {S.day(stamp)}.",
                          "evidence": {"last_modified": st.get("last_modified"), "as_of": st.get("as_of")},
                          "question": "Who refreshes this report, and did the refresh stop?",
                          "id_parts": [src["id"], "stale"]})
    return found


# ---------------------------------------------------------------- Excel audit findings → stages

def place_excel(cfg, audit_result, registered):
    """Add stage + metric to each Excel finding from the header or row label the cell sits under."""
    import xlsx_audit as xa
    wb = xa.Workbook(audit_result["path"])
    is_source = audit_result["workbook"] in registered
    default_stage = registered.get(audit_result["workbook"])

    def metric_for_text(text):
        t = (text or "").lower()
        for m, info in METRICS.items():
            if m == "job_key":
                continue
            if any(re.search(rf"(?<![a-z]){re.escape(s)}(?![a-z])", t) for s in info["synonyms"]):
                return m
        return None

    def metric_for_cell(sheet, a1):
        pos = xa.split_cell(a1) if re.fullmatch(r"[A-Z]{1,3}\d+", a1 or "") else None
        if not pos or sheet not in wb.cells:
            return metric_for_text(a1)
        cells = wb.cells[sheet]
        label = (cells.get((1, pos[1])) or {}).get("value")
        m = metric_for_text(label) if isinstance(label, str) else None
        if m:
            return m
        for r in range(pos[1] - 1, 0, -1):
            v = (cells.get((pos[0], r)) or {}).get("value")
            if isinstance(v, str) and v.strip() and not re.fullmatch(r"-?[\d.]+", v):
                return metric_for_text(v)
        return None

    out = []
    for f in audit_result["findings"]:
        f = dict(f)
        f["category"] = "excel"
        if not is_source:
            f["metric"], f["stage"] = None, None
        else:
            m = metric_for_cell(f["sheet"], f["cell"]) if f["sheet"] else metric_for_text(f["cell"])
            if not m and f.get("used_by"):
                us, uc = f["used_by"][0].rsplit("!", 1)
                m = metric_for_cell(us, uc)
            f["metric"] = m
            f["stage"] = stage_of(m) if m else default_stage
        f["title"] = f"{f['workbook']} › {f['sheet'] or 'workbook'} › {f['cell']}: {f['kind'].replace('_', ' ')}"
        f["detail"] = f["problem"]
        f["question"] = None
        f["onset"] = None
        f["id_parts"] = [f["workbook"], f["sheet"] or "", f["kind"]] + _cell_identity(wb, f)
        out.append(f)
    return out


def _cell_identity(wb, f):
    """What the cell *is*, so a total that moves down as rows are added stays the same finding:
    column + the row's label (column A) + the formula's relative shape. Falls back to the address."""
    import xlsx_audit as xa
    cell = f.get("cell") or ""
    pos = xa.split_cell(cell) if re.fullmatch(r"[A-Z]{1,3}\d+", cell) else None
    if not pos or not f.get("sheet") or f["sheet"] not in wb.cells:
        return [cell]
    label = (wb.cells[f["sheet"]].get((1, pos[1])) or {}).get("value")
    if label in (None, ""):
        return [cell]
    if isinstance(label, str) and not re.fullmatch(r"-?[\d.]+", label.strip()):
        return [xa.col_letters(pos[0]), label.strip()]      # a labeled row ("Total good"): label is enough
    shape = xa.r1c1(f["formula"], pos) if f.get("formula") else ""
    return [xa.col_letters(pos[0]), str(label), shape]      # a data row (dated): add the formula's shape



# ---------------------------------------------------------------- definitions that live in formulas

def metric_for_label(label):
    """The known metric a label names, if any (by its synonyms)."""
    t = (label or "").lower()
    for m, info in METRICS.items():
        if m != "job_key" and any(re.search(rf"(?<![a-z]){re.escape(s)}(?![a-z])", t) for s in info["synonyms"]):
            return m
    return None


def formula_groups(paths):
    """Every label calculated in 2+ places across the workbooks, with its definitions (formula_defs.py)."""
    import formula_defs as fd
    lines = []
    for p in paths:
        try:
            lines += fd.calculated_lines(p)
        except Exception:      # an unreadable workbook is reported by the audit
            continue
    return fd.group(lines)


def formula_definitions(cfg, paths, formulas_dictionary, registered=None):
    """One finding per label that the plant's spreadsheets calculate in more than one way. With a
    confirmed definition, only the places that differ from it are reported."""
    registered = registered or {}
    found = []
    for g in formula_groups(paths):
        key = g["label"].lower()
        entry = (formulas_dictionary or {}).get(key) or {}
        variants = g["variants"]
        confirmed = entry.get("status") == "confirmed" and entry.get("definition")
        if confirmed:
            off = [v for v in variants if v["definition"] != entry["definition"]]
            if not off:
                continue
            n_off = sum(v["places"] for v in off)
            detail = (f"Your definition of {g['label']} is {entry['definition']}. {n_off} place(s) calculate it differently: "
                      + "; ".join(f"{v['definition']} in {v['places']} place(s), e.g. {v['examples'][0]}" for v in off) + ".")
            question = f"Should these {n_off} place(s) be changed to your definition of {g['label']}, or is the definition wrong?"
        elif len(variants) < 2:
            continue
        else:
            detail = (f"{g['label']} is calculated {len(variants)} different ways: "
                      + "; ".join(f"{v['definition']} in {v['places']} place(s) across {v['workbooks']} workbook(s), "
                                  f"e.g. {v['examples'][0]}" for v in variants) + ".")
            question = f"Which of these is how {g['label']} should be calculated? (It's fine to say 'not settled yet'.)"
        metric = metric_for_label(g["label"])
        books = Counter(ex.split(" › ")[0] for v in variants for ex in v["examples"])
        stage = stage_of(metric) or next((registered[b] for b, _ in books.most_common() if registered.get(b)), None)
        found.append({
            "kind": "formula_definition_conflict", "category": "definition", "metric": metric, "stage": stage,
            "onset": None, "sources": [], "label": g["label"], "places": g["places"], "workbooks": g["workbooks"],
            "copies": g["copies"], "confirmed_definition": entry.get("definition") if confirmed else None,
            "title": f"{g['label']}: {len(variants)} definitions across {g['places']} places"
                     + (f" in {g['workbooks']} workbooks" if g["workbooks"] > 1 else ""),
            "detail": detail, "question": question,
            "evidence": {"variants": [{"definition": v["definition"], "places": v["places"], "workbooks": v["workbooks"],
                                       "example": v["examples"][0], "written": v["written"][:2],
                                       "is_yours": bool(confirmed and v["definition"] == entry["definition"])}
                                      for v in variants]},
            "id_parts": ["formula", key]})
    return sorted(found, key=lambda f: (-f["workbooks"], -f["places"]))
