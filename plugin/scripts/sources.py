"""Reading the plant's sources into one shape. Standard library only.

A source config (sources.json) is written once at setup (the model drafts it from the
file headers, the person confirms it) and then read by code on every run:

{
  "plant": "Northfork Precision",
  "folder": "/path/to/plant",              # sources' globs are relative to this
  "sources": [
    {"id": "mes_downtime", "label": "MES downtime events", "glob": "exports/mes_downtime_events_*.csv",
     "kind": "events", "date": "production_day", "shift": "shift", "entity": "asset_id",
     "measures": {"downtime": "duration_min"}, "category": "reason_code", "record_for": ["downtime"]},
    {"id": "shift_log", "label": "Supervisor shift log", "glob": "spreadsheets/supervisor_shift_log_*.xlsx",
     "sheet": "Shift Log", "kind": "rows", "date": "Date", "shift": "Shift", "entity": "Machine",
     "key": "Job", "key_master": "erp_wo", "measures": {"good": "Good", "scrap": "Scrap", "downtime": "Downtime (min)"},
     "notes_sheet": "Notes"},
    {"id": "reason_codes", "kind": "codes", "glob": "exports/mes_reason_codes.csv", "code": "reason_code",
     "flags": {"counts_as_downtime": "downtime"}},
    {"id": "erp_wo", "kind": "master", "glob": "erp/production_work_orders.csv", "key": "production_wo",
     "stage": "prod::Work Order"},
    {"id": "weekly_pack", "kind": "report", "glob": "reports/*.pdf", "series": "Weekly OEE pack"},
    {"id": "daily_summary", "kind": "link", "path": "/path/or/null", "uri": "file:///drive/item",
     "expect": "daily", "measures": ["good", "scrap", "downtime"], "stage": "prod::Production Reporting"}
  ]
}
"""

import csv
import glob
import io
import json
import os
import re
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

import safeio
import xlsx_audit as xa


def day(d):
    """Nov 9. Written out because the no-padding date code isn't portable."""
    return f"{d:%b} {d.day}"


def load_config(path):
    cfg = json.loads(Path(path).read_text(encoding="utf-8"))
    cfg["_dir"] = str(Path(path).resolve().parent)
    return cfg


def files_for(cfg, src):
    if src.get("glob"):
        pattern = src["glob"].replace("\\", "/")
        if pattern.startswith("/") or re.match(r"[A-Za-z]:", pattern) or ".." in pattern.split("/"):
            raise safeio.Refused(f"{src.get('id')}: '{src['glob']}' points outside the plant folder")
        return sorted(Path(p) for p in glob.glob(os.path.join(glob.escape(cfg["folder"]), src["glob"]), recursive=True)
                      if safeio.within(p, cfg["folder"]))     # a link out of the folder isn't followed
    if src.get("path"):
        p = Path(src["path"])
        if p.suffix.lower() not in (".xlsx", ".xlsm", ".csv", ".pdf"):
            raise safeio.Refused(f"{src.get('id')}: a linked report must be an Excel, CSV or PDF file")
        return [p] if p.exists() else []
    return []


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if re.fullmatch(r"\d+(\.\d+)?", s):          # Excel serial
        return date(1899, 12, 30) + timedelta(days=int(float(s)))
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        pass
    for f in ("%m/%d/%Y", "%m/%d/%y", "%d %b %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def read_table(path, sheet=None):
    """(header, rows as dicts, row numbers). CSV or the named xlsx sheet (values only)."""
    p = Path(path)
    safeio.check_size(p)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        wb = xa.Workbook(p)
        name = sheet or wb.sheet_names[0]
        cells = wb.cells.get(name, {})
        if not cells:
            return [], [], []
        max_col = max(c for c, _ in cells)
        max_row = max(r for _, r in cells)
        header = [str((cells.get((c, 1)) or {}).get("value") or "").strip() for c in range(1, max_col + 1)]
        rows, nums = [], []
        for r in range(2, max_row + 1):
            vals = [(cells.get((c, r)) or {}).get("value") for c in range(1, max_col + 1)]
            if any(v not in (None, "") for v in vals):
                rows.append(dict(zip(header, vals)))
                nums.append(r)
        return header, rows, nums
    data = p.read_bytes()
    try:
        text = data.decode("utf-16" if data[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252", "replace")       # what Excel on Windows saves a plain CSV as
    raw = list(csv.reader(io.StringIO(text, newline="")))
    if not raw:
        return [], [], []
    header = [h.strip() for h in raw[0]]
    rows = [dict(zip(header, r)) for r in raw[1:] if any(x.strip() for x in r)]
    return header, rows, list(range(2, len(rows) + 2))


UNREADABLE = (OSError, ValueError, KeyError, csv.Error, zipfile.BadZipFile, ET.ParseError)


def table(path, sheet=None):
    """read_table, with a file that can't be read as an empty one. data_files() reports it."""
    try:
        return read_table(path, sheet)
    except UNREADABLE:
        return [], [], []


def data_files(cfg):
    """[(source, path, problem or None)] for each export or workbook a source reads its rows from."""
    out = []
    for src in cfg["sources"]:
        if src.get("kind") not in ("rows", "events", "codes", "master"):
            continue
        for path in files_for(cfg, src):
            try:
                _, rows, _ = read_table(path, src.get("sheet"))
                out.append((src, path, None if rows else "it has no rows"))
            except UNREADABLE as e:
                out.append((src, path, f"it couldn't be read ({e})"))
    return out


def sheet_lines(path, sheet):
    """Free text of a sheet (a Notes tab), one line per non-empty row."""
    try:
        wb = xa.Workbook(path)
    except UNREADABLE:
        return []
    if sheet not in wb.sheet_names:
        return []
    cells = wb.cells[sheet]
    lines = []
    for r in sorted({r for _, r in cells}):
        parts = [str(cells[(c, rr)]["value"]) for (c, rr) in sorted(cells) if rr == r and cells[(c, rr)]["value"]]
        if parts:
            lines.append(" ".join(parts))
    return lines


def facts(cfg, as_of=None):
    """Normalized records from rows/events sources:
    {source, file, row, date, shift, entity, key, category, metric: value ...}"""
    out = []
    for src in cfg["sources"]:
        if src.get("kind") not in ("rows", "events"):
            continue
        for path in files_for(cfg, src):
            header, rows, nums = table(path, src.get("sheet"))
            for row, rn in zip(rows, nums):
                d = to_date(row.get(src["date"]))
                if d is None or (as_of and d > as_of):
                    continue
                rec = {"source": src["id"], "file": path.name, "row": rn, "date": d,
                       "shift": str(row.get(src.get("shift"), "") or "").strip().split(".")[0],
                       "entity": str(row.get(src.get("entity"), "") or "").strip(),
                       "key": str(row.get(src["key"], "") or "").strip() if src.get("key") else None,
                       "category": str(row.get(src["category"], "") or "").strip() if src.get("category") else None}
                for metric, col in src.get("measures", {}).items():
                    rec[metric] = num(row.get(col))
                out.append(rec)
    return out


def codes(cfg):
    """Reason-code flags: {source_id: {code: {flag: value}}}."""
    out = {}
    for src in cfg["sources"]:
        if src.get("kind") != "codes":
            continue
        for path in files_for(cfg, src):
            _, rows, _ = table(path)
            out[src["id"]] = {str(r[src["code"]]).strip(): {k: str(r.get(k, "")).strip() for k in src.get("flags", {})}
                              for r in rows}
    return out


def master_keys(cfg, src_id):
    src = next(s for s in cfg["sources"] if s["id"] == src_id)
    keys = set()
    for path in files_for(cfg, src):
        _, rows, _ = table(path)
        keys |= {str(r[src["key"]]).strip() for r in rows if r.get(src["key"])}
    return keys


def notes(cfg):
    """{source_id: [lines]} for sources with a notes sheet."""
    out = {}
    for src in cfg["sources"]:
        if src.get("notes_sheet"):
            for path in files_for(cfg, src):
                out[src["id"]] = sheet_lines(path, src["notes_sheet"])
    return out


def headers(cfg):
    out = {}
    for src in cfg["sources"]:
        if src.get("kind") in ("rows", "events", "master", "codes"):
            fs = files_for(cfg, src)
            if fs:
                out[src["id"]] = table(fs[-1], src.get("sheet"))[0]
    return out


def _all_workbooks(cfg):
    paths = {Path(p) for ext in ("*.xlsx", "*.xlsm") for p in glob.glob(os.path.join(glob.escape(cfg["folder"]), "**", ext), recursive=True)
             if safeio.within(p, cfg["folder"])}
    for src in cfg["sources"]:
        if src.get("kind") == "link" and src.get("path") and Path(src["path"]).suffix.lower() in (".xlsx", ".xlsm"):
            if Path(src["path"]).exists():
                paths.add(Path(src["path"]))
    return sorted(p for p in paths if not p.name.startswith("~$"))


def _formula_print(path):
    """Same sheets and same formulas in the same cells: the same workbook, whatever values it last saved."""
    import hashlib
    h = hashlib.sha256()
    try:
        with safeio.open_book(path) as z:
            for n in sorted(z.namelist()):
                if n == "xl/workbook.xml" or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n):
                    xml = safeio.book_part(z, n).decode("utf-8", "replace")
                    if n == "xl/workbook.xml":
                        h.update("|".join(re.findall(r'<sheet [^>]*name="([^"]*)"', xml)).encode())
                    else:
                        h.update(n.encode())
                        for cell, f in re.findall(r'<c r="([A-Z]+\d+)"[^>]*>\s*(<f[^>]*?(?:/>|>.*?</f>))', xml, re.S):
                            h.update(f"{cell}={f};".encode())
    except (OSError, zipfile.BadZipFile, KeyError, safeio.Refused):
        return None   # unreadable: never merged, so the audit reports it
    return h.hexdigest()


def workbook_copies(cfg):
    """{kept workbook: [other files that are the same workbook]}. A source's own file is the one kept,
    then the shortest path; a backup or a cached-values copy elsewhere in the folder isn't audited twice."""
    registered = {p for src in cfg["sources"] for p in files_for(cfg, src)}
    groups = {}
    for p in _all_workbooks(cfg):
        groups.setdefault(_formula_print(p) or str(p), []).append(p)
    out = {}
    for ps in groups.values():
        ps.sort(key=lambda p: (p not in registered, len(p.parts), str(p)))
        out[ps[0]] = ps[1:]
    return out


def workbooks(cfg):
    """Every distinct .xlsx/.xlsm in the plant folder (and any linked local file), registered or not."""
    return sorted(workbook_copies(cfg))
