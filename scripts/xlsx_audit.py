#!/usr/bin/env python3
"""Excel formula audit. Standard library only.

Reads .xlsx/.xlsm at the XML level (formulas as written, cached values, error cells,
defined names, external links, hidden rows and sheets, shared formulas) and reports
every problem with workbook, sheet, cell, formula, the problem in plain words and a fix.

    python3 xlsx_audit.py BOOK.xlsx [BOOK2.xlsx ...]      # prints JSON

Nothing is recalculated. Errors Excel would show are found statically ("will show
#DIV/0!"), and cached error values are read when the file has them.
"""

import json
from bisect import bisect_left, bisect_right
from functools import lru_cache
import posixpath
import re
import sys
import zipfile  # noqa: F401
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

import safeio

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
M = "{%s}" % NS["m"]

SEVERITY = {  # lower sorts first
    "cached_error": 1, "ref_error": 1, "broken_defined_name": 1, "circular_reference": 1,
    "div_by_blank": 2, "lookup_out_of_range": 2, "missing_sheet_reference": 2, "unknown_function": 2,
    "external_link": 3, "wrong_period_sheet": 3, "hidden_row_in_range": 3, "hidden_sheet_reference": 3,
    "empty_range": 3, "range_short": 4, "pattern_break": 5, "overwritten_formula": 5,
    "hardcoded_constant": 6, "text_in_numeric_column": 6, "duplicate_row": 6,
}
SEVERITY_LABEL = {1: "broken", 2: "will break", 3: "improper reference", 4: "range error",
                  5: "inconsistent", 6: "hidden assumption"}

MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
AGG = {"SUM", "SUMIF", "SUMIFS", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "COUNT", "COUNTA", "COUNTIF",
       "COUNTIFS", "MAX", "MIN", "SUBTOTAL", "AGGREGATE", "SUMPRODUCT"}
KNOWN_FUNCS = set("""
ABS ACOS ADDRESS AGGREGATE AND AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS CEILING CEILING.MATH CHAR CHOOSE CHOOSECOLS
CHOOSEROWS CLEAN CODE COLUMN COLUMNS CONCAT CONCATENATE CONVERT CORREL COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS
DATE DATEDIF DATEVALUE DAY DAYS DB DDB DEGREES DOLLAR DROP EDATE EOMONTH ERROR.TYPE EVEN EXACT EXP FALSE FILTER
FIND FIXED FLOOR FLOOR.MATH FORECAST FORECAST.LINEAR FV GROWTH HLOOKUP HOUR HSTACK HYPERLINK IF IFERROR IFNA IFS
INDEX INDIRECT INT INTERCEPT IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER
ISODD ISREF ISTEXT LAMBDA LARGE LEFT LEN LET LINEST LN LOG LOG10 LOOKUP LOWER MATCH MAX MAXA MAXIFS MEDIAN MID
MIN MINA MINIFS MINUTE MOD MODE MONTH MROUND N NA NETWORKDAYS NETWORKDAYS.INTL NOT NOW NPER NPV ODD OFFSET OR
PERCENTILE PERCENTILE.INC PERCENTRANK PI PMT POWER PPMT PRODUCT PROPER PV QUARTILE QUOTIENT RADIANS RAND
RANDBETWEEN RANK RANK.EQ RATE REPLACE REPT RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SECOND SEQUENCE SIGN SIN
SLOPE SMALL SORT SORTBY SQRT STDEV STDEV.P STDEV.S STDEVA SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SWITCH
T TAKE TAN TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TODAY TOCOL TOROW TRANSPOSE TREND TRIM
TRUE TRUNC TYPE UNIQUE UPPER VALUE VAR VAR.P VAR.S VLOOKUP VSTACK WEEKDAY WEEKNUM WORKDAY WORKDAY.INTL XIRR XLOOKUP
XMATCH XNPV XOR YEAR YEARFRAC ISOWEEKNUM GETPIVOTDATA CELL INFO SHEET SHEETS FORMULATEXT ARRAYTOTEXT VALUETOTEXT
""".split())

REF_RE = re.compile(r"""
    (?<![A-Za-z0-9_.\]])
    (?P<sheet>(?:'(?:[^']|'')+'|\[[^\]]+\][A-Za-z0-9_. ]*?|[A-Za-z_][A-Za-z0-9_.]*)!)?
    (?P<ref>
        \$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?
      | \$?[A-Z]{1,3}:\$?[A-Z]{1,3}
      | \$?\d+:\$?\d+
      | \#REF!
    )
    (?![A-Za-z0-9_(])
""", re.X)
FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:_xlfn\.|_xlws\.)?[A-Za-z][A-Za-z0-9.]*)\(")
NAME_RE = re.compile(r"(?<![A-Za-z0-9_.!$'\]])([A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_.(!])")
STR_RE = re.compile(r'"(?:[^"]|"")*"')
CONST_RE = re.compile(r"(?<![A-Za-z0-9_.$:])(\d+(?:\.\d+)?)(?![A-Za-z0-9_.:(])")


# ---------------------------------------------------------------- A1 helpers

def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_cell(a1):
    m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", a1)
    return (col_num(m.group(1)), int(m.group(2))) if m else None


def blank_strings(f):
    return STR_RE.sub(lambda m: " " * len(m.group(0)), f)


def sheet_of(prefix, host):
    """(sheet name, external book or None) from a REF_RE sheet prefix."""
    if not prefix:
        return host, None
    s = prefix[:-1]
    if s.startswith("'"):
        s = s[1:-1].replace("''", "'")
    m = re.match(r"\[([^\]]+)\](.*)", s)
    if m:
        return m.group(2), m.group(1)
    return s, None


# ---------------------------------------------------------------- reading the workbook

class Workbook:
    def __init__(self, path):
        self.path = Path(path)
        self.name = self.path.name
        with safeio.open_book(path) as z:       # closed before the audit, so the file is never held open
            self._read(z)

    def _read(self, z):
        names = set(z.namelist())
        rels = {r.get("Id"): r.get("Target") for r in
                ET.fromstring(safeio.book_part(z, "xl/_rels/workbook.xml.rels")).findall("rel:Relationship", NS)}
        wb = ET.fromstring(safeio.book_part(z, "xl/workbook.xml"))
        self.sheets = []        # (name, state, xml path)
        for s in wb.find("m:sheets", NS):
            target = rels[s.get("{%s}id" % NS["r"])]
            target = posixpath.normpath(posixpath.join("xl", target)) if not target.startswith("/") else target[1:]
            self.sheets.append((s.get("name"), s.get("state", "visible"), target))
        self.sheet_names = [s[0] for s in self.sheets]
        self.defined_names = {}
        dn = wb.find("m:definedNames", NS)
        for d in (dn if dn is not None else []):
            self.defined_names[d.get("name")] = (d.text or "").strip()
        self.external = []
        for n in sorted(names):
            if re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", n):
                relp = n.replace("externalLinks/", "externalLinks/_rels/") + ".rels"
                target = ""
                if relp in names:
                    r = ET.fromstring(safeio.book_part(z, relp)).find("rel:Relationship", NS)
                    target = r.get("Target", "") if r is not None else ""
                self.external.append(target)
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(safeio.book_part(z, "xl/sharedStrings.xml")).findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter(M + "t")))
        self.cells = {}         # sheet -> {(col,row): cell dict}
        self.hidden_rows = {}   # sheet -> set(rows)
        for name, _, target in self.sheets:
            cells, hidden, masters = {}, set(), {}
            root = ET.fromstring(safeio.book_part(z, target))
            for row in root.iter(M + "row"):
                if row.get("hidden") in ("1", "true"):
                    hidden.add(int(row.get("r")))
                for c in row.findall(M + "c"):
                    pos = split_cell(c.get("r"))
                    t = c.get("t")
                    v = c.find(M + "v")
                    f = c.find(M + "f")
                    val = v.text if v is not None else None
                    if t == "s" and val is not None:
                        val = shared[int(val)]
                    elif t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter(M + "t"))
                    cell = {"t": t, "value": val, "formula": None, "error": val if t == "e" else None}
                    if f is not None:
                        if f.get("t") == "shared":
                            si = f.get("si")
                            if f.text:
                                masters[si] = (pos, f.text)
                                cell["formula"] = "=" + f.text
                            else:
                                cell["shared_si"] = si
                        elif f.text:
                            cell["formula"] = "=" + f.text
                    cells[pos] = cell
            for pos, cell in cells.items():
                si = cell.pop("shared_si", None)
                if si is not None and si in masters:
                    (mc, mr), text = masters[si]
                    cell["formula"] = "=" + translate(text, pos[1] - mr, pos[0] - mc)
            self.cells[name] = cells
            self.hidden_rows[name] = hidden

    def value(self, sheet, pos):
        return self.cells.get(sheet, {}).get(pos)

    def is_blank(self, sheet, pos):
        c = self.value(sheet, pos)
        return c is None or (c["formula"] is None and (c["value"] is None or str(c["value"]).strip() == ""))

    def _index(self, sheet):
        """Per column, sorted rows: filled (a value or a formula), with a value, with a formula. Built once."""
        if not hasattr(self, "_idx"):
            self._idx = {}
        if sheet not in self._idx:
            filled, valued, formulas = defaultdict(list), defaultdict(list), defaultdict(list)
            for (c, r), cell in self.cells.get(sheet, {}).items():
                if cell["value"] not in (None, ""):
                    valued[c].append(r)
                if cell["value"] not in (None, "") or cell["formula"]:
                    filled[c].append(r)
                if cell["formula"]:
                    formulas[c].append(r)
            for d in (filled, valued, formulas):
                for rows in d.values():
                    rows.sort()
            mr = max((r for (_, r) in self.cells.get(sheet, {})), default=0)
            self._idx[sheet] = (filled, valued, formulas, mr)
        return self._idx[sheet]

    def last_row(self, sheet, col):
        rows = self._index(sheet)[0].get(col)
        return rows[-1] if rows else 0

    def last_value_row_before(self, sheet, col, row):
        """The last row above `row` in this column holding a value (0 if none)."""
        rows = self._index(sheet)[1].get(col, [])
        i = bisect_left(rows, row)
        return rows[i - 1] if i else 0

    def formula_cells_in(self, sheet, c1, c2, r1, r2):
        idx = self._index(sheet)[2]
        for c in range(c1, c2 + 1):
            rows = idx.get(c, [])
            for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)]:
                yield (c, r)

    def max_row(self, sheet):
        return self._index(sheet)[3]


def translate(formula, drow, dcol):
    """Shift relative references by (drow, dcol), as Excel does for a shared formula."""
    def shift(m):
        ref = m.group("ref")
        if ref == "#REF!":
            return m.group(0)
        def one(part):
            mm = re.fullmatch(r"(\$?)([A-Z]{1,3})?(\$?)(\d+)?", part)
            ca, cl, ra, rw = mm.groups()
            out = ""
            if cl:
                out += ca + (cl if ca else col_letters(col_num(cl) + dcol))
            if rw:
                out += ra + (rw if ra else str(int(rw) + drow))
            return out
        return (m.group("sheet") or "") + ":".join(one(p) for p in ref.split(":"))
    text = blank_strings(formula)
    out, last = [], 0
    for m in REF_RE.finditer(text):
        out.append(formula[last:m.start()])
        out.append(shift(m))
        last = m.end()
    out.append(formula[last:])
    return "".join(out)


@lru_cache(maxsize=None)
def r1c1(formula, pos):
    """Relative R1C1 form, used to compare formulas down a column or across a row."""
    col, row = pos
    def conv(m):
        ref = m.group("ref")
        if ref == "#REF!":
            return m.group(0)
        def one(part):
            mm = re.fullmatch(r"(\$?)([A-Z]{1,3})?(\$?)(\d+)?", part)
            ca, cl, ra, rw = mm.groups()
            out = ""
            if rw:
                out += f"R{rw}" if ra else f"R[{int(rw) - row}]"
            if cl:
                out += f"C{col_num(cl)}" if ca else f"C[{col_num(cl) - col}]"
            return out
        return (m.group("sheet") or "") + ":".join(one(p) for p in ref.split(":"))
    text = blank_strings(formula)
    out, last = [], 0
    for m in REF_RE.finditer(text):
        out.append(formula[last:m.start()])
        out.append(conv(m))
        last = m.end()
    out.append(formula[last:])
    return "".join(out)


@lru_cache(maxsize=None)
def refs_in(formula, host):
    """Every reference in a formula: dicts with sheet, external, kind, bounds."""
    out = []
    for m in REF_RE.finditer(blank_strings(formula)):
        sheet, ext = sheet_of(m.group("sheet"), host)
        ref = m.group("ref")
        r = {"text": m.group(0), "sheet": sheet, "external": ext, "ref": ref}
        if ref == "#REF!":
            r["kind"] = "broken"
        elif re.fullmatch(r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}", ref):
            a, b = ref.replace("$", "").split(":")
            r.update(kind="columns", c1=col_num(a), c2=col_num(b))
        elif re.fullmatch(r"\$?\d+:\$?\d+", ref):
            a, b = ref.replace("$", "").split(":")
            r.update(kind="rows", r1=int(a), r2=int(b))
        else:
            parts = ref.replace("$", "").split(":")
            (c1, r1), (c2, r2) = split_cell(parts[0]), split_cell(parts[-1])
            r.update(kind="cell" if len(parts) == 1 else "range", c1=min(c1, c2), c2=max(c1, c2),
                     r1=min(r1, r2), r2=max(r1, r2))
        out.append(r)
    return out


@lru_cache(maxsize=None)
def calls(formula):
    """Function calls with their top-level argument strings, nested calls included."""
    text = blank_strings(formula)
    found = []
    for m in FUNC_RE.finditer(text):
        name = m.group(1).upper().replace("_XLFN.", "").replace("_XLWS.", "")
        i, depth, args, start = m.end(), 1, [], m.end()
        while i < len(text) and depth:
            ch = text[i]
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    args.append(formula[start:i].strip())
            elif ch == "," and depth == 1:
                args.append(formula[start:i].strip())
                start = i + 1
            i += 1
        found.append((name, args))
    return found


# ---------------------------------------------------------------- the audit

USED_BY_KEEP = 20   # formulas listed per hidden-row finding (the count is always exact)


def audit(path, period_hint=None):
    wb = Workbook(path)
    F = []
    checks = ["cached errors", "#REF! in formulas", "defined names", "missing sheets", "external links",
              "wrong-period sheets", "hidden rows and sheets in ranges", "range ends", "pattern breaks",
              "overwritten formulas", "hard-coded constants", "circular references", "division by blank",
              "lookup ranges", "unknown functions", "empty ranges", "text in numeric columns", "duplicate rows"]
    not_checked = []

    flagged = set()

    def add(kind, sheet, cell, problem, fix, formula=None, cached=None, **extra):
        flagged.add((sheet, cell))
        F.append({"kind": kind, "severity": SEVERITY[kind], "severity_label": SEVERITY_LABEL[SEVERITY[kind]],
                  "workbook": wb.name, "sheet": sheet, "cell": cell, "formula": formula, "cached": cached,
                  "problem": problem, "fix": fix, **extra})

    period = period_hint or next((m for m in MONTHS if re.search(rf"(?<![a-z]){m}", wb.name.lower())), None)
    formula_cells = [(s, pos, c) for s in wb.sheet_names for pos, c in wb.cells[s].items() if c["formula"]]
    defined_upper = {n.upper() for n in wb.defined_names}

    # Defined names
    for name, target in wb.defined_names.items():
        if name.startswith("_xlnm."):
            continue
        if "#REF!" in target:
            add("broken_defined_name", None, name, f"The name {name} points to {target}: what it referred to was deleted.",
                f"Point {name} at the right range again (Formulas > Name Manager), or delete it if it isn't used.",
                formula=target)
        else:
            for r in refs_in(target, None):
                if r["sheet"] and not r["external"] and r["sheet"] not in wb.sheet_names:
                    add("broken_defined_name", None, name, f"The name {name} refers to a sheet '{r['sheet']}' that doesn't exist.",
                        f"Point {name} at an existing sheet (Formulas > Name Manager).", formula=target)

    for target in wb.external:
        add("external_link", None, "(workbook links)", f"This workbook links to another file: {target or 'unknown path'}.",
            "Check the linked file is current and reachable by everyone who opens this one, or paste the values in.")

    for sheet, pos, c in formula_cells:
        a1 = f"{col_letters(pos[0])}{pos[1]}"
        f = c["formula"]
        refs = refs_in(f, sheet)
        fixed_here = set()
        if "[" in blank_strings(f) and not any(r["external"] for r in refs):
            not_checked.append({"sheet": sheet, "cell": a1, "formula": f, "reason": "structured table reference"})
        if "#REF!" in blank_strings(f):
            add("ref_error", sheet, a1, "The formula refers to something that was deleted (#REF!), so it can't calculate.",
                "Rebuild the reference to the right cells; check what was deleted (a column, row or sheet).",
                formula=f, cached=c["error"])
            fixed_here.add("ref")
        for r in refs:
            if r["external"]:
                add("external_link", sheet, a1, f"Pulls a value from another workbook, {r['external']}, which may be missing or out of date.",
                    "Bring the value into this workbook or document where the linked file lives and who updates it.",
                    formula=f, cached=c["error"])
            elif r["sheet"] not in wb.sheet_names and r["kind"] != "broken":
                add("missing_sheet_reference", sheet, a1, f"Refers to a sheet '{r['sheet']}' that isn't in this workbook.",
                    f"Add the '{r['sheet']}' sheet or point the formula at the right one.", formula=f,
                    cached=c["error"], will_show="#REF!")
                fixed_here.add("missing")
            elif period and r["sheet"] != sheet:
                other = next((m for m in MONTHS if re.search(rf"(?<![a-z]){m}", r["sheet"].lower())), None)
                if other and other != period:
                    add("wrong_period_sheet", sheet, a1,
                        f"This {period.title()} workbook takes a figure from the '{r['sheet']}' sheet, a different month.",
                        f"Point it at the {period.title()} data, or label the cell clearly as {r['sheet']}.",
                        formula=f, cached=c["error"])
            if r["kind"] in ("cell", "range", "columns") and r["sheet"] in wb.sheet_names:
                state = next(s[1] for s in wb.sheets if s[0] == r["sheet"])
                if state != "visible" and r["sheet"] != sheet:
                    add("hidden_sheet_reference", sheet, a1, f"Uses data on a hidden sheet, '{r['sheet']}'.",
                        "Unhide and check it, or document why it's hidden.", formula=f, cached=c["error"])
        # Lookups
        for name, args in calls(f):
            if name not in KNOWN_FUNCS and name.upper() not in defined_upper:
                add("unknown_function", sheet, a1, f"Excel doesn't recognize {name}(), so the cell will show #NAME?.",
                    "Check the spelling, or that the add-in providing it is installed.", formula=f,
                    cached=c["error"], will_show="#NAME?")
            if name in ("VLOOKUP", "HLOOKUP") and len(args) >= 3 and re.fullmatch(r"\d+", args[2] or ""):
                tr = refs_in(args[1], sheet)
                idx = int(args[2])
                if tr and tr[0]["kind"] in ("range", "columns") and "missing" not in fixed_here:
                    width = (tr[0]["c2"] - tr[0]["c1"] + 1) if name == "VLOOKUP" else (tr[0].get("r2", 0) - tr[0].get("r1", 0) + 1)
                    if idx > width:
                        add("lookup_out_of_range", sheet, a1,
                            f"{name} asks for column {idx} of a range only {width} wide, so it will show #REF!.",
                            f"Widen the lookup range or fix the column number (it must be {width} or less).",
                            formula=f, cached=c["error"], will_show="#REF!")
            if name in AGG:
                for arg in args:
                    for r in refs_in(arg, sheet):
                        if r["kind"] == "range" and not r["external"] and r["sheet"] in wb.sheet_names:
                            if all(wb.is_blank(r["sheet"], (cc, rr)) for cc in range(r["c1"], r["c2"] + 1)
                                   for rr in range(r["r1"], r["r2"] + 1)):
                                add("empty_range", sheet, a1, f"{name} runs over {r['text']}, which is empty.",
                                    "Point it at the cells that hold the data.", formula=f, cached=c["error"])
        # Division by a blank or zero cell
        text = blank_strings(f)
        for m in re.finditer(r"/\s*(" + REF_RE.pattern + r")", text, re.X):
            r = refs_in(m.group(1), sheet)
            if r and r[0]["kind"] == "cell" and not r[0]["external"] and r[0]["sheet"] in wb.sheet_names:
                tpos = (r[0]["c1"], r[0]["r1"])
                tc = wb.value(r[0]["sheet"], tpos)
                zero = tc is not None and tc["formula"] is None and str(tc["value"]).strip() in ("0", "0.0")
                if wb.is_blank(r[0]["sheet"], tpos) or zero:
                    add("div_by_blank", sheet, a1,
                        f"Divides by {r[0]['text']}, which is {'zero' if zero else 'blank'}, so it will show #DIV/0!.",
                        f"Fill in {r[0]['text']}, or wrap the formula in IF/IFERROR to show a blank until it's filled.",
                        formula=f, cached=c["error"], will_show="#DIV/0!")
                    fixed_here.add("div")
        # A cached error with no static explanation
        if c["error"] and not fixed_here and (sheet, a1) not in flagged:
            add("cached_error", sheet, a1, f"Excel shows {c['error']} here.", "Open the cell and trace the input that fails.",
                formula=f, cached=c["error"])

    # Hidden rows inside ranges that formulas use. A whole-column reference covers every hidden row on
    # its sheet, so it is counted once per sheet, not once per row (101k whole-column formulas x 31k
    # hidden rows is 3 billion pairs). Each finding keeps a count and the first USED_BY_KEEP users.
    hidden_sorted = {sh: sorted(rows) for sh, rows in wb.hidden_rows.items() if rows}
    by_row = defaultdict(lambda: [0, []])              # (sheet, hidden row) -> [count, [(formula idx, cell)]]
    by_col = defaultdict(lambda: [0, []])              # sheet -> same, for formulas with a whole-column reference
    if hidden_sorted:
        for idx, (s2, pos, c) in enumerate(formula_cells):
            name = None
            refs = [r for r in refs_in(c["formula"], s2) if not r["external"] and r["sheet"] in hidden_sorted]
            whole = {r["sheet"] for r in refs if r["kind"] == "columns"}
            for sh in whole:
                name = name or f"{s2}!{col_letters(pos[0])}{pos[1]}"
                e = by_col[sh]
                e[0] += 1
                if len(e[1]) < USED_BY_KEEP:
                    e[1].append((idx, name))
            seen = set()
            for r in refs:
                if r["sheet"] in whole or r["kind"] not in ("range", "rows"):
                    continue
                hs = hidden_sorted[r["sheet"]]
                for hr in hs[bisect_left(hs, r["r1"]):bisect_right(hs, r["r2"])]:
                    if (r["sheet"], hr) not in seen:
                        seen.add((r["sheet"], hr))
                        name = name or f"{s2}!{col_letters(pos[0])}{pos[1]}"
                        e = by_row[(r["sheet"], hr)]
                        e[0] += 1
                        if len(e[1]) < USED_BY_KEEP:
                            e[1].append((idx, name))
    for sheet, rows in wb.hidden_rows.items():
        for hr in sorted(rows):
            rc, rx = by_row.get((sheet, hr), (0, []))
            cc, cx = by_col.get(sheet, (0, []))
            count = rc + cc
            if count:
                users = [u for _, u in sorted(rx + cx)[:USED_BY_KEEP]]
                add("hidden_row_in_range", sheet, f"A{hr}",
                    f"Row {hr} is hidden but still counted by {count} formula(s), e.g. {', '.join(users[:3])}.",
                    "Unhide it and check the values belong in the totals, or delete the row if it shouldn't count.",
                    used_by=users, used_by_count=count)

    # Range ends: a total that stops short of the data
    for sheet, pos, c in formula_cells:
        a1 = f"{col_letters(pos[0])}{pos[1]}"
        for name, args in calls(c["formula"]):
            if name not in AGG:
                continue
            for arg in args:
                for r in refs_in(arg, sheet):
                    if r["kind"] != "range" or r["c1"] != r["c2"] or r["external"] or r["sheet"] not in wb.sheet_names:
                        continue
                    last = wb.last_row(r["sheet"], r["c1"])
                    if r["sheet"] == sheet and pos[0] == r["c1"] and pos[1] > r["r2"]:
                        last = wb.last_value_row_before(sheet, r["c1"], pos[1])
                    if r["r1"] <= 2 and r["r2"] < last and all(not wb.is_blank(r["sheet"], (r["c1"], rr))
                                                                for rr in range(r["r2"] + 1, last + 1)):
                        add("range_short", sheet, a1,
                            f"{name} covers {r['text']} but the data in column {col_letters(r['c1'])} runs to row {last}, "
                            f"so {last - r['r2']} row(s) are left out.",
                            f"Extend the range to row {last}, or use the whole column.", formula=c["formula"],
                            cached=c["error"], missing_rows=last - r["r2"])

    # Pattern breaks, overwritten formulas and hard-coded constants, column by column
    for sheet in wb.sheet_names:
        by_col = defaultdict(dict)
        for (col, row), c in wb.cells[sheet].items():
            by_col[col][row] = c
        for col, cells in by_col.items():
            rows = sorted(cells)
            runs, cur = [], []
            for rr in rows:
                if cur and rr != cur[-1] + 1:
                    runs.append(cur)
                    cur = []
                if not wb.is_blank(sheet, (col, rr)):
                    cur.append(rr)
                elif cur:
                    runs.append(cur)
                    cur = []
            if cur:
                runs.append(cur)
            for run in runs:
                forms = {rr: r1c1(cells[rr]["formula"], (col, rr)) for rr in run if cells[rr]["formula"]}
                if len(forms) < 3:
                    continue
                common, n = Counter(forms.values()).most_common(1)[0]
                if n / len(run) < 0.6:
                    continue
                for rr in run:
                    a1 = f"{col_letters(col)}{rr}"
                    cell = cells[rr]
                    if not cell["formula"]:
                        if re.fullmatch(r"-?\d+(\.\d+)?", str(cell["value"] or "")):
                            add("overwritten_formula", sheet, a1,
                                f"A typed number ({cell['value']}) sits where the rest of column {col_letters(col)} has a formula, "
                                "so it won't update when the data changes.",
                                f"Restore the column's formula (copy it from {col_letters(col)}{rr - 1 if rr - 1 in forms else rr + 1}).",
                                cached=cell["value"])
                        continue
                    form = forms[rr]
                    if form == common:
                        continue
                    extra = form.replace(common.rstrip(), "") if form.startswith(common.rstrip()) else None
                    if extra and re.fullmatch(r"\s*[-+*/]\s*\d+(\.\d+)?\s*", extra):
                        add("hardcoded_constant", sheet, a1,
                            f"Same formula as the column plus a typed {extra.strip()}: an adjustment nobody can see.",
                            "Remove the constant, or put the adjustment in its own labeled cell and reference it.",
                            formula=cell["formula"], cached=cell["error"])
                    else:
                        add("pattern_break", sheet, a1,
                            f"This formula differs from the rest of column {col_letters(col)} (usually a bad copy or a manual edit).",
                            f"Copy the column's formula back in, unless the difference is intended.",
                            formula=cell["formula"], cached=cell["error"],
                            expected=translate(next(cells[x]["formula"] for x in run if forms.get(x) == common),
                                               rr - next(x for x in run if forms.get(x) == common), 0))
        # Hard-coded constants in one-off formulas (outside any consistent column)
    in_runs = {(x["sheet"], x["cell"]) for x in F}
    shapes = defaultdict(Counter)                      # (sheet, column) -> how often each formula shape appears
    for s2, p2, c2 in formula_cells:
        shapes[(s2, p2[0])][r1c1(c2["formula"], p2)] += 1
    for sheet, pos, c in formula_cells:
        a1 = f"{col_letters(pos[0])}{pos[1]}"
        if (sheet, a1) in in_runs:
            continue
        if shapes[(sheet, pos[0])][r1c1(c["formula"], pos)] >= 3:
            continue  # the same constant down a whole column is a design choice, not a hidden tweak
        text = blank_strings(c["formula"])
        arith = [m.group(1) for m in re.finditer(r"[-+*/]\s*(\d+(?:\.\d+)?)(?![A-Za-z0-9_.:(])", text[1:])]
        arith = [x for x in arith if x not in ("0", "1", "100")]
        if arith and not any(n in KNOWN_FUNCS for n, _ in calls(c["formula"]) if n in ("ROUND", "VLOOKUP", "INDEX")):
            add("hardcoded_constant", sheet, a1, f"The formula hides a typed value ({', '.join(arith)}).",
                "Put the value in its own labeled input cell and reference it.", formula=c["formula"], cached=c["error"])

    # Circular references
    graph = {}
    for sheet, pos, c in formula_cells:
        deps = set()
        for r in refs_in(c["formula"], sheet):
            if r["external"] or r["sheet"] not in wb.sheet_names or r["kind"] not in ("cell", "range", "columns"):
                continue
            r1, r2 = (1, wb.max_row(r["sheet"])) if r["kind"] == "columns" else (r["r1"], r["r2"])
            deps.update((r["sheet"], p) for p in wb.formula_cells_in(r["sheet"], r["c1"], r["c2"], r1, r2))
        graph[(sheet, pos)] = sorted(deps)          # a fixed order: set order changes with every process
    # One finding per tangle: a strongly connected set of cells (Tarjan, iterative so a tangle of
    # thousands of cells can't overflow the stack). Every loop through one tangle is the same
    # problem; listing each loop gave one real workbook 332,197 findings over 16,263 cells.
    order = {n: i for i, n in enumerate(graph)}
    succ_of = {n: [m for m in deps if m in graph] for n, deps in graph.items()}   # built once
    index, low, on, stack, comps, counter = {}, {}, set(), [], [], 0
    for root in graph:
        if root in index:
            continue
        work = [(root, 0)]
        while work:
            node, i = work.pop()
            if i == 0:
                index[node] = low[node] = counter
                counter += 1
                stack.append(node)
                on.add(node)
            succ = succ_of[node]
            for j in range(i, len(succ)):
                nxt = succ[j]
                if nxt not in index:
                    work.append((node, j + 1))
                    work.append((nxt, 0))
                    break
                if nxt in on:
                    low[node] = min(low[node], index[nxt])
            else:
                if low[node] == index[node]:
                    comp = []
                    while True:
                        n = stack.pop()
                        on.discard(n)
                        comp.append(n)
                        if n == node:
                            break
                    comps.append(comp)
                if work:
                    parent = work[-1][0]
                    low[parent] = min(low[parent], low[node])
    tangles = [sorted(c, key=order.get) for c in comps if len(c) > 1 or c[0] in graph[c[0]]]
    for comp in sorted(tangles, key=lambda c: order[c[0]]):
        rep, members = comp[0], set(comp)
        prev, queue = {rep: None}, [rep]            # shortest way back to rep, inside the tangle
        back = None
        while queue and back is None:
            nxt_q = []
            for n in queue:
                for m in graph[n]:
                    if m == rep:
                        back = n
                        break
                    if m in members and m not in prev:
                        prev[m] = n
                        nxt_q.append(m)
                if back is not None:
                    break
            queue = nxt_q
        path, n = [], back
        while n is not None:
            path.append(n)
            n = prev[n]
        loop = list(reversed(path)) + [rep]
        names = [f"{s}!{col_letters(p[0])}{p[1]}" for s, p in loop]
        s0, p0 = rep
        size = len(comp)
        add("circular_reference", s0, f"{col_letters(p0[0])}{p0[1]}",
            (f"Circular reference: {' -> '.join(names)}." if size <= len(loop) - 1 else
             f"Circular reference: {size:,} cells refer to each other in a loop, e.g. {' -> '.join(names[:8])}"
             + (" -> ..." if len(names) > 8 else "") + ".")
            + " Excel either refuses to calculate or iterates to a meaningless number.",
            "Break the loop: one of these cells must take its value from outside the chain."
            + (" Start with the formulas that read whole columns, which pull in their own cells." if size > 10 else ""),
            formula=wb.value(s0, p0)["formula"], loop=names, tangle_size=size,
            tangle=[f"{s}!{col_letters(p[0])}{p[1]}" for s, p in comp[:20]])

    # Tables: text in numeric columns, duplicate rows (only sheets with a header row of 3+ labels)
    for sheet in wb.sheet_names:
        cells = wb.cells[sheet]
        header = [c for (cc, rr), c in cells.items() if rr == 1 and c["value"] and not c["formula"]]
        if len(header) < 3:
            continue
        last = wb.max_row(sheet)
        cols = sorted({cc for (cc, rr) in cells if rr == 1})
        for cc in cols:
            vals = [(rr, cells.get((cc, rr))) for rr in range(2, last + 1)]
            nums = [rr for rr, c in vals if c and c["t"] not in ("s", "inlineStr", "str", "e") and c["value"] not in (None, "")]
            texts = [rr for rr, c in vals if c and c["t"] in ("s", "inlineStr") and str(c["value"]).strip()]
            if len(nums) >= 10 and 0 < len(texts) <= len(nums) * 0.1:
                for rr in texts:
                    add("text_in_numeric_column", sheet, f"{col_letters(cc)}{rr}",
                        f"Text ('{cells[(cc, rr)]['value']}') in a column of numbers; totals skip it silently.",
                        "Type the number, or move the note to a comments column.")
        seen_rows = {}
        for rr in range(2, last + 1):
            vals = tuple(str((cells.get((cc, rr)) or {}).get("value") or "") for cc in cols)
            if any(vals) and not any((cells.get((cc, rr)) or {}).get("formula") for cc in cols):
                if vals in seen_rows:
                    add("duplicate_row", sheet, f"A{rr}", f"Row {rr} repeats row {seen_rows[vals]} exactly.",
                        "Delete the duplicate if it was entered twice.")
                else:
                    seen_rows[vals] = rr

    F.sort(key=lambda x: (x["severity"], x["sheet"] or "", x["cell"]))
    return {"workbook": wb.name, "path": str(wb.path), "sheets": wb.sheet_names, "checks_run": checks,
            "not_checked": not_checked, "findings": F,
            "counts": dict(Counter(f["kind"] for f in F))}


if __name__ == "__main__":
    print(json.dumps([audit(p) for p in sys.argv[1:]], indent=2, default=str))
