#!/usr/bin/env python3
"""plant-watch: find where a plant's metrics, reports and spreadsheets disagree. Standard library only.

    python3 watch.py scan       PLANT_FOLDER                     # files, headers, sheets: for drafting sources.json
    python3 watch.py inventory  --config sources.json
    python3 watch.py run        --config sources.json --state STATE [--as-of 2026-11-25]
    python3 watch.py report     --config sources.json --state STATE [--out report.html] [--pdf report.pdf]
    python3 watch.py export     --state STATE --out findings.csv|findings.json
    python3 watch.py check-dictionary   --state STATE                  # is the dictionary sound?
    python3 watch.py set-status --state STATE --id ID --status acknowledged --by NAME --reason TEXT
    python3 watch.py audit      BOOK.xlsx [...]                  # Excel audit on any workbook
    python3 watch.py pdf-text   REPORT.pdf                      # page text, to check a reading

STATE holds dictionary.json (confirmed metric meanings), findings.jsonl (the ledger),
snapshots/ (one per run), extractions/ (report readings, written at setup/run by the
skill and confirmed by the person) and links.json (last-modified per linked report).
Every command prints one JSON object. Nothing here writes to the plant's own files.
"""

import argparse
import csv
import glob
import io
import json
import os
import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import definitions  # noqa: E402
import detect  # noqa: E402
import ledger  # noqa: E402
import pdftext  # noqa: E402
import safeio  # noqa: E402
import sources as S  # noqa: E402
import xlsx_audit as xa  # noqa: E402


def out(obj, code=0):
    print(json.dumps(obj, indent=2, default=str))
    sys.exit(code)


def read_json(path, default):
    p = Path(path)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else default


def scan(folder):
    """What's in a plant folder, grouped by file pattern, so setup can draft sources.json.
    Shows headers, sheets and two sample rows per group; reads nothing else."""
    import re as _re
    groups = {}
    for p in sorted(Path(folder).rglob("*")):
        if not p.is_file() or p.name.startswith(("~$", ".")):
            continue
        rel = p.relative_to(folder)
        if safeio.OWN_FOLDER in rel.parts or not safeio.within(p, folder):
            continue
        pattern = (rel.parent / _re.sub(r"\d{4}-\d{2}-\d{2}", "*", p.name)).as_posix()
        g = groups.setdefault(pattern, {"pattern": pattern, "files": 0, "type": p.suffix.lower().lstrip(".")})
        g["files"] += 1
        if "headers" in g or "sheets" in g:
            continue
        try:
            if p.suffix.lower() == ".csv":
                h, rows, _ = S.read_table(p)
                g["headers"], g["sample"] = h, rows[:2]
            elif p.suffix.lower() in (".xlsx", ".xlsm"):
                wb = xa.Workbook(p)
                g["sheets"] = {}
                for sh in wb.sheet_names:
                    h, rows, _ = S.read_table(p, sh)
                    g["sheets"][sh] = {"headers": h, "sample": rows[:2], "rows": len(rows)}
            elif p.suffix.lower() == ".pdf":
                pg = pdftext.pages(p)
                g["pages"] = len(pg)
                g["first_lines"] = (pg[0] or [])[:6] if pg else []
        except Exception as e:
            g["error"] = str(e)
    return {"folder": str(folder), "groups": list(groups.values())}


def inventory(cfg):
    """Every metric found in two or more sources, with the evidence for how each counts it."""
    facts = S.facts(cfg)
    codes = S.codes(cfg)
    notes = S.notes(cfg)
    hdrs = S.headers(cfg)
    found = {}
    for metric in detect.METRICS:
        srcs = sorted({f["source"] for f in facts if f.get(metric) is not None})
        reports = [s["id"] for s in cfg["sources"] if s.get("kind") in ("report", "link")
                   and metric in s.get("measures", [metric])]
        if len(srcs) + len(reports) < 2:
            continue
        ev = {}
        for sid in srcs:
            src = next(s for s in cfg["sources"] if s["id"] == sid)
            e = {"label": src.get("label", sid), "column": src.get("measures", {}).get(metric), "notes": [], "flags": {}}
            e["notes"] = [l for l in notes.get(sid, []) if any(w in l.lower() for w in detect.METRICS[metric]["synonyms"])]
            if src.get("category"):
                cats = sorted({f["category"] for f in facts if f["source"] == sid and f.get(metric) is not None})
                e["categories"] = cats
                for cid, flags in codes.items():
                    e["flags"][cid] = {c: v for c, v in flags.items() if c in cats}
            ev[sid] = e
        vals, cats = detect.cells_by_source(facts, metric)
        if len(vals) >= 2:
            rec = detect.record_source(cfg, metric, None, set(vals))
            for sid in vals:
                if sid == rec:
                    continue
                tot_a = sum(vals[rec].values())
                tot_b = sum(v for c, v in vals[sid].items() if c[0] in {x[0] for x in vals[rec]})
                ev[sid]["vs_" + rec] = {"total_" + rec: round(tot_a, 1), "total_" + sid: round(tot_b, 1)}
        for d in detect.definitions(cfg, [f for f in facts if f.get(metric) is not None], notes, None):
            if d["metric"] == metric and d["sources"][1] in ev:
                ev[d["sources"][1]]["differs_from_" + d["sources"][0]] = {
                    "how": d["detail"], "since": d["onset"], "gap": d.get("gap"), "explained_by": d.get("code")}
        found[metric] = {"label": detect.METRICS[metric]["label"], "stage": detect.METRICS[metric]["stage"],
                         "sources": ev, "reports": reports, "question": detect.METRICS[metric]["question"]}
    formulas = [{"label": g["label"], "places": g["places"], "workbooks": g["workbooks"], "copies": g["copies"],
                 "metric": detect.metric_for_label(g["label"]),
                 "definitions": [{"definition": v["definition"], "places": v["places"], "workbooks": v["workbooks"],
                                  "example": v["examples"][0]} for v in g["variants"]]}
                for g in detect.formula_groups(S.workbooks(cfg)) if len(g["variants"]) > 1]
    formulas.sort(key=lambda g: (-g["workbooks"], -g["places"]))
    return {"metrics": found, "headers": hdrs, "formula_definitions": formulas}


def stated_as_of(path):
    """The date beside an "As of" cell on a workbook's first sheet, if there is one."""
    if Path(path).suffix.lower() not in (".xlsx", ".xlsm"):
        return None
    try:
        wb = xa.Workbook(path)
        cells = wb.cells[wb.sheet_names[0]]
    except Exception:
        return None
    for (c, r), cell in cells.items():
        if str(cell.get("value") or "").strip().lower().rstrip(":") == "as of":
            d = S.to_date((cells.get((c + 1, r)) or {}).get("value"))
            if d:
                return d.isoformat()
    return None


def link_states(cfg, state):
    states = read_json(Path(state) / "links.json", {})
    for src in cfg["sources"]:
        if src.get("kind") == "link" and src.get("path") and Path(src["path"]).exists():
            mt = datetime.fromtimestamp(os.path.getmtime(src["path"]))
            st = states.setdefault(src["id"], {})
            st["last_modified"] = mt.isoformat(timespec="minutes")
            # What the report says it covers beats its file time, which copying or downloading resets.
            stated = stated_as_of(src["path"])
            if stated:
                st["as_of"] = stated
    return states


def extractions(cfg, state):
    present = {}
    for src in cfg["sources"]:
        if src.get("kind") == "report":
            for p in S.files_for(cfg, src):
                present[p.name] = (p, src)
    # A linked report has one reading, refreshed in place: matched by its file name or its source id. A link
    # read only through the connector has no local file, so its figures are compared but can't be re-checked.
    links = {src["id"]: src for src in cfg["sources"] if src.get("kind") == "link"}
    for src in links.values():
        for p in S.files_for(cfg, src):
            present[p.name] = (p, src)
    out_ = []
    for f in sorted(glob.glob(str(Path(state) / "extractions" / "*.json"))):
        ex = json.loads(Path(f).read_text(encoding="utf-8"))
        if ex.get("file") in present:
            p, src = present[ex["file"]]
            ex["path"] = str(p)
        elif ex.get("source") in links:
            src = links[ex["source"]]
            ex["path"] = None
        else:
            continue
        ex.setdefault("series", src.get("series", src["id"]))
        ex.setdefault("file", src.get("label") or src["id"])
        out_.append(ex)
    return out_


def snapshot_drift(cfg, state, run_date, codes, hdrs):
    snaps = sorted(glob.glob(str(Path(state) / "snapshots" / "*.json")))
    prev = json.loads(Path(snaps[-1]).read_text(encoding="utf-8")) if snaps else None
    found = []
    if prev and prev.get("date") != run_date.isoformat():
        for cid, flags in codes.items():
            for code, fl in flags.items():
                old = prev.get("codes", {}).get(cid, {}).get(code)
                if old and old != fl:
                    found.append({"kind": "code_flag_changed", "category": "definition", "metric": "downtime",
                                  "stage": detect.stage_of("downtime"), "onset": run_date, "sources": [cid],
                                  "title": f"Reason code {code}: flags changed", "detail": f"Was {old}, now {fl}.",
                                  "evidence": {"before": old, "after": fl},
                                  "question": "Was this change agreed, and from which date does it apply?",
                                  "id_parts": [cid, code, json.dumps(fl, sort_keys=True)]})
        for sid, h in hdrs.items():
            old = prev.get("headers", {}).get(sid)
            if old and old != h:
                found.append({"kind": "header_changed", "category": "definition", "metric": None, "stage": None,
                              "onset": run_date, "sources": [sid], "title": f"{detect.src_label(cfg, sid)}: columns changed",
                              "detail": f"Added: {sorted(set(h) - set(old)) or 'none'}; removed: {sorted(set(old) - set(h)) or 'none'}.",
                              "evidence": {"before": old, "after": h},
                              "question": "Did a column change meaning, or was it renamed?",
                              "id_parts": [sid, "headers", "|".join(h)]})
    return found


def run(cfg, state, as_of=None):
    state = Path(state)
    safeio.check_place(state, cfg.get("folder"))
    state.mkdir(parents=True, exist_ok=True)
    safeio.remember(state, cfg)
    dictionary = read_json(state / "dictionary.json", {}).get("metrics", {})
    facts = S.facts(cfg, as_of)
    latest = max((f["date"] for f in facts), default=as_of)
    as_of = as_of or latest
    notes, codes, hdrs = S.notes(cfg), S.codes(cfg), S.headers(cfg)
    findings = []
    findings += detect.definitions(cfg, facts, notes, dictionary)
    findings += detect.keys(cfg, facts, notes)
    findings += detect.anomalies(cfg, facts, dictionary)
    findings += detect.reports(cfg, facts, extractions(cfg, state), dictionary)
    links = link_states(cfg, state)
    findings += detect.links(cfg, links, sorted({f['date'] for f in facts}))
    findings += snapshot_drift(cfg, state, as_of, codes, hdrs)
    registered = {}
    for src in cfg["sources"]:
        for p in S.files_for(cfg, src):
            if p.suffix.lower() in (".xlsx", ".xlsm"):
                registered[p.name] = src.get("stage") or detect.stage_of(next(iter(src.get("measures", {}) or [None]), None))
    audits = []
    copies = S.workbook_copies(cfg)

    def shown(p):
        try:
            return str(Path(p).relative_to(cfg["folder"]))
        except ValueError:
            return str(p)
    for wb in sorted(copies):
        try:
            a = xa.audit(wb)
        except Exception as e:  # a workbook we can't open is itself worth saying
            findings.append({"kind": "workbook_unreadable", "category": "excel", "metric": None, "stage": None,
                             "onset": None, "workbook": wb.name, "sheet": None, "cell": None,
                             "title": f"{wb.name}: couldn't be read", "detail": str(e), "question": None,
                             "id_parts": [wb.name, "unreadable"]})
            continue
        audits.append({"workbook": a["workbook"], "path": shown(wb), "copies": [shown(c) for c in copies[wb]],
                       "checks_run": a["checks_run"], "not_checked": a["not_checked"]})
        findings += detect.place_excel(cfg, a, registered)
    formulas = read_json(state / "dictionary.json", {}).get("formulas", {})
    findings += detect.formula_definitions(cfg, S.workbooks(cfg), formulas, registered)
    for f in findings:
        if f.get("onset") is None:
            f["onset"] = None
    delta = ledger.update(state, findings, as_of)
    snap = {"date": as_of.isoformat(), "codes": codes, "headers": hdrs, "notes": notes, "links": links,
            "audits": audits, "dictionary_version": read_json(state / "dictionary.json", {}).get("version")}
    safeio.save_state(state / "snapshots" / f"{as_of.isoformat()}.json", json.dumps(snap, indent=2, default=str))
    with (state / "runs.jsonl").open("a", encoding="utf-8") as f:
        f.write(json.dumps({"as_of": as_of.isoformat(), "at": datetime.now().isoformat(timespec="seconds"),
                            "new": len(delta["new"]), "changed": len(delta["changed"]),
                            "cleared": len(delta["cleared"]), "active": len(delta["active"])}) + "\n")
    return {"as_of": as_of, "latest_data": latest, "new": [e["id"] for e in delta["new"]],
            "changed": [e["id"] for e in delta["changed"]], "cleared": [e["id"] for e in delta["cleared"]],
            "active": len(delta["active"]), "audits": audits}


def export(state, path, sources=()):
    """The findings as CSV or JSON, and the confirmed definitions beside them. Both stay on this computer."""
    entries = list(ledger.load(state).values())
    p = Path(path)
    if p.suffix == ".json":
        text = json.dumps({"schema": "plant-watch.findings/v1", "findings": entries}, indent=2, default=str)
    else:
        cols = ["id", "kind", "category", "stage", "metric", "title", "detail", "question", "onset", "first_seen",
                "last_seen", "active", "cleared_on", "status", "status_by", "status_on", "status_reason",
                "workbook", "sheet", "cell", "formula", "fix", "report", "gap", "share"]
        buf = io.StringIO()
        w = csv.DictWriter(buf, cols, extrasaction="ignore", lineterminator="\r\n")
        w.writeheader()
        for e in entries:
            w.writerow({k: safeio.csv_safe(e.get(k)) for k in cols})
        text = "\ufeff" + buf.getvalue()         # so Excel reads the accents
    safeio.save_result(state, p, text, sources)
    res = {"ok": True, "path": str(p), "count": len(entries)}
    dictionary = read_json(Path(state) / "dictionary.json", {})
    if any(e.get("status") == "confirmed" for _, _, e in definitions.entries(dictionary)):
        doc = definitions.document(dictionary, safeio.settings(state).get("plant", ""))
        dp = p.with_name("plant-watch-definitions.json")
        safeio.save_result(state, dp, definitions.text(doc), sources)
        res.update({"definitions_path": str(dp), "definitions": len(doc["definitions"]),
                    "not_settled": doc["not_settled"], "problems": definitions.check(dictionary)})
    return res


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        command()
    except safeio.Refused as e:
        out({"ok": False, "error": str(e)}, 1)


def command():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["scan", "inventory", "run", "report", "export", "set-status", "audit", "pdf-text",
                                    "check-dictionary"])
    ap.add_argument("files", nargs="*")
    ap.add_argument("--config")
    ap.add_argument("--state")
    ap.add_argument("--as-of")
    ap.add_argument("--out")
    ap.add_argument("--pdf")
    ap.add_argument("--id")
    ap.add_argument("--status")
    ap.add_argument("--by")
    ap.add_argument("--reason")
    a = ap.parse_args()
    if a.cmd == "scan":
        out(scan(a.files[0]))
    if a.cmd == "audit":
        out([xa.audit(f) for f in a.files])
    if a.cmd == "pdf-text":
        out({f: pdftext.pages(f) for f in a.files})
    if a.cmd == "set-status":
        try:
            out({"ok": True, "finding": ledger.set_status(a.state, a.id, a.status, a.by, a.reason)})
        except (ValueError, KeyError) as e:
            out({"ok": False, "error": str(e)}, 1)
    if a.cmd == "export":
        out(export(a.state, a.out or "findings.csv"))
    if a.cmd == "check-dictionary":
        problems = definitions.check(read_json(Path(a.state) / "dictionary.json", {}))
        out({"ok": not problems, "problems": problems}, 0 if not problems else 1)
    cfg = S.load_config(a.config)
    if a.cmd == "inventory":
        out(inventory(cfg))
    if a.cmd == "run":
        out(run(cfg, a.state, date.fromisoformat(a.as_of) if a.as_of else None))
    if a.cmd == "report":
        import report
        out(report.render(cfg, a.state, a.out or str(Path(a.state) / "metric-mismatch-report.html"), a.pdf))


if __name__ == "__main__":
    main()
