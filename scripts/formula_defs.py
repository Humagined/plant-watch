#!/usr/bin/env python3
"""Definitions that live in formulas. Standard library only.

    python3 scripts/formula_defs.py WORKBOOK.xlsx [...]   [--json]

A spreadsheet plant defines its metrics in formulas: "Performance %" is whatever the formula in that
column computes. This finds every calculated line (3+ consecutive formulas down a column under a
header, or along a row after a row label), rewrites the formula in terms of the labels it reads
instead of cell addresses, and groups lines by label across sheets and workbooks. The same label
with a different expression is a definition question, not a broken cell:

    Hours/Day   = [Run Hours] / [Days in Month]           20 places
                = [Standard Hours] / [Days in Month]       2 places

Rewriting by label removes layout: =E5/F5 on one sheet and =G5/H5 on another are the same
definition when E and G are both "Run Hours". Set aside so they don't read as a different
definition: wrappers (IFERROR, IFNA, ROUND), quoted selectors ("Apr", a customer), which row of a
block is read, and a plain copy of the same label from another sheet. Period labels (Jan, Q3,
YTD) are columns of one table, not metrics, and are skipped.
"""

import json
import re
import sys
from bisect import bisect_left
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsx_audit as xa  # noqa: E402

MIN_RUN = 3
REACH = 40                 # rows up (or columns left) to look for a line's label
NUMERIC = re.compile(r"\s*[-+]?[\d.,]+%?\s*")
WRAPPERS = ("IFERROR", "IFNA", "_XLFN.IFNA", "ROUND", "ROUNDUP", "ROUNDDOWN")
MONTHS = "jan feb mar apr may jun jul aug sep sept oct nov dec january february march april june july august " \
         "september october november december".split()
PERIOD = re.compile(r"(?i)^(?:(?:%s)\b|q[1-4]\b|h[12]\b|fy\s?\d|ytd\b|mtd\b|qtd\b|wk\s?\d|week\s?\d|\d{4}\b|"
                    r"\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?\b)" % "|".join(MONTHS))


def norm_label(v):
    return re.sub(r"\s+", " ", str(v)).strip().rstrip(":").strip()


class Labels:
    """Text cells by column and by row, to name a line by the nearest label above it or to its left."""

    def __init__(self, wb):
        self.wb = wb
        self._c, self._r = {}, {}

    def _build(self, sheet):
        if sheet in self._c:
            return
        bc, br = defaultdict(list), defaultdict(list)
        for (c, r), cell in self.wb.cells.get(sheet, {}).items():
            v = cell["value"]
            if v not in (None, "") and not cell["formula"] and cell["t"] in ("s", "inlineStr", "str") \
                    and not NUMERIC.fullmatch(str(v)):
                bc[c].append(r)
                br[r].append(c)
        for d in (bc, br):
            for v in d.values():
                v.sort()
        self._c[sheet], self._r[sheet] = bc, br

    def _val(self, sheet, c, r):
        return norm_label(self.wb.cells[sheet][(c, r)]["value"])

    def above(self, sheet, col, row):
        self._build(sheet)
        rows = self._c[sheet].get(col, [])
        i = bisect_left(rows, row)
        return self._val(sheet, col, rows[i - 1]) if i and row - rows[i - 1] <= REACH else None

    def left(self, sheet, col, row):
        self._build(sheet)
        cols = self._r[sheet].get(row, [])
        i = bisect_left(cols, col)
        return self._val(sheet, cols[i - 1], row) if i and col - cols[i - 1] <= REACH else None

    def top(self, sheet, col):
        self._build(sheet)
        rows = self._c[sheet].get(col, [])
        return self._val(sheet, col, rows[0]) if rows else None


def refs(formula, host):
    """(match, sheet, external, ref-without-$, anchors) for each reference outside string literals."""
    for m in xa.REF_RE.finditer(xa.blank_strings(formula)):
        sheet, ext = xa.sheet_of(m.group("sheet"), host)
        yield m, sheet, ext, m.group("ref").replace("$", ""), m.group("ref").count("$")


def orientation(formula, host, pos):
    """'col' when the formula reads its own row (a column definition), 'row' when it reads its own
    column at other rows (a row definition, e.g. =C1/C2 under month headers), None if neither."""
    same_row = same_col = 0
    for m, sheet, ext, ref, anchors in refs(formula, host):
        if ext or sheet != host or ":" in ref or ref == "#REF!" or anchors == 2:
            continue
        c, r = xa.split_cell(ref)
        same_row += r == pos[1] and c != pos[0]
        same_col += c == pos[0] and r != pos[1]
    return "row" if same_col > same_row else "col" if same_row else None


def rewrite(formula, host, pos, labels, wb, mode):
    """The formula with every reference replaced by the label it reads. mode 'col': cells are named
    by the header above; 'row': by the row label to their left."""
    name_of = (lambda s, c, r: labels.above(s, c, r) or labels.top(s, c)) if mode == "col" else \
              (lambda s, c, r: labels.left(s, c, r))
    out, last = [], 0
    for m, sheet, ext, ref, anchors in refs(formula, host):
        out.append(formula[last:m.start()])
        last = m.end()
        where = "" if sheet == host else "@other"
        if ext or sheet not in wb.sheet_names:
            out.append("[external]" if ext else "[missing sheet]")
            continue
        if ref == "#REF!":
            out.append("[#REF!]")
            continue
        if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", ref):
            a, b = ref.split(":")
            n = [labels.top(sheet, xa.col_num(x)) or f"col {x}" for x in (a, b)]
            out.append(f"[{n[0]}]" if a == b else f"[{n[0]} .. {n[1]}]")
            out.append(where)
            continue
        if re.fullmatch(r"\d+:\d+", ref):
            out.append("[rows]" + where)
            continue
        parts = ref.split(":")
        (c1, r1), (c2, r2) = xa.split_cell(parts[0]), xa.split_cell(parts[-1])
        if len(parts) == 1 and anchors == 2:      # $B$2: one fixed input
            name = labels.above(sheet, c1, r1) or labels.left(sheet, c1, r1)
            out.append(f"{{{name or 'input'}}}" + where)
            continue
        if mode == "col":
            n1, n2 = name_of(sheet, c1, r1), name_of(sheet, c2, r1)
        else:
            n1, n2 = name_of(sheet, c1, r1), name_of(sheet, c1, r2)
        n1 = n1 or f"col {xa.col_letters(c1)}" if mode == "col" else n1 or f"row {r1}"
        n2 = n2 or n1
        out.append((f"[{n1}]" if n1 == n2 else f"[{n1} .. {n2}]") + where)
    out.append(formula[last:])
    s = "".join(out).lstrip("=")
    # spaces only matter inside labels and strings
    return re.sub(r'(\[[^\]]*\]|\{[^}]*\}|"(?:[^"]|"")*")|\s+', lambda t: t.group(1) or "", s)


# Arguments that choose WHICH item to read (a customer, a part, a month), not how it is computed.
SELECTOR_ARGS = {"SUMIF": {0, 1}, "COUNTIF": {1}, "AVERAGEIF": {0, 1}, "VLOOKUP": {0}, "HLOOKUP": {0}, "XLOOKUP": {0},
                 "_XLFN.XLOOKUP": {0}, "MATCH": {0}, "GETPIVOTDATA": "all-but-0",
                 "SUMIFS": "all-but-0", "COUNTIFS": "odd", "AVERAGEIFS": "all-but-0"}
GENERIC = {"total", "subtotal", "sub total", "grand total", "sum", "totals", "average", "avg", "count"}


def _split_args(body):
    args, depth, start = [], 0, 0
    for i, ch in enumerate(body):
        depth += {"(": 1, ")": -1, "{": 1, "}": -1}.get(ch, 0)
        if ch == "," and depth == 0:
            args.append(body[start:i])
            start = i + 1
    args.append(body[start:])
    return args


def blank_selectors(expr):
    """SUMIF([Customer],[Acme],[Total]) -> SUMIF([Customer],…,[Total]), recursively."""
    out, i = [], 0
    pat = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\(")
    while True:
        m = pat.search(expr, i)
        if not m:
            out.append(expr[i:])
            break
        name = m.group(1).upper()
        depth, j = 1, m.end()
        while j < len(expr) and depth:
            depth += {"(": 1, ")": -1}.get(expr[j], 0)
            j += 1
        body = expr[m.end():j - 1]
        args = [blank_selectors(a) for a in _split_args(body)]
        rule = SELECTOR_ARGS.get(name)
        if rule:
            for k in range(len(args)):
                if (rule == "all-but-0" and k > 0) or (rule == "even-from-2" and k >= 2 and k % 2 == 0) or \
                        (rule == "odd" and k % 2 == 1) or (isinstance(rule, set) and k in rule):
                    args[k] = "…"
        out.append(expr[i:m.start()] + m.group(1) + "(" + ",".join(args) + ")")
        i = j
    return "".join(out)


def core(expr):
    """The definition without what doesn't change meaning: IFERROR/ROUND wrappers, quoted selectors,
    and repeats of the same label (a SUM of one column's subtotal rows is one thing however many)."""
    changed = True
    while changed:
        changed = False
        for w in WRAPPERS:
            if expr.upper().startswith(w + "("):
                body, depth, i = expr[len(w) + 1:], 1, 0
                while i < len(body):
                    depth += {"(": 1, ")": -1}.get(body[i], 0)
                    if depth == 1 and body[i] == ",":
                        break
                    i += 1
                if i < len(body):
                    expr, changed = body[:i], True
                break
    expr = re.sub(r'"(?:[^"]|"")*"', '"…"', expr)
    expr = blank_selectors(expr)
    expr = re.sub(r"\((?:row|col)[+-]\d+\)", "", expr)
    # SUM([X],[X],[X]) and [X]+[X]+[X] -> one [X]
    expr = re.sub(r"(\[[^\]]+\](?:@other)?)(?:[,+]\1)+", r"\1…", expr)
    # a plain sum of terms is the same definition in any order
    terms = _split_top(expr, "+")
    if len(terms) > 1 and all(re.fullmatch(r"[\[{][^\]}]+[\]}](?:@other)?…?", t) for t in terms):
        expr = "+".join(sorted(terms))
    return expr


def _split_top(expr, op):
    parts, depth, start = [], 0, 0
    for i, ch in enumerate(expr):
        depth += {"(": 1, ")": -1, "[": 1, "]": -1, "{": 1, "}": -1}.get(ch, 0)
        if ch == op and depth == 0:
            parts.append(expr[start:i])
            start = i + 1
    parts.append(expr[start:])
    return parts


def runs_of(positions):
    """Consecutive runs in a sorted list of ints."""
    runs, cur = [], []
    for p in positions:
        if cur and p != cur[-1] + 1:
            runs.append(cur)
            cur = []
        cur.append(p)
    if cur:
        runs.append(cur)
    return runs


def calculated_lines(path):
    """Every run of 3+ formulas down a column (named by its header) or along a row (by its row label)."""
    wb = xa.Workbook(path)
    labels = Labels(wb)
    found = []
    for sheet in wb.sheet_names:
        forms = {p: c["formula"] for p, c in wb.cells[sheet].items() if c["formula"]}
        orient = {p: orientation(f, sheet, p) for p, f in forms.items()}
        by_col, by_row = defaultdict(list), defaultdict(list)
        for (c, r) in forms:
            by_col[c].append(r)
            by_row[r].append(c)
        lines = []
        for c, rows in by_col.items():
            for run in runs_of(sorted(rows)):
                cells = [(c, r) for r in run if orient[(c, r)] != "row"]
                if len(cells) >= MIN_RUN:
                    lines.append(("col", cells, labels.above(sheet, c, run[0])))
        for r, cols in by_row.items():
            for run in runs_of(sorted(cols)):
                cells = [(c, r) for c in run if orient[(c, r)] == "row"]
                if len(cells) >= MIN_RUN:
                    lines.append(("row", cells, labels.left(sheet, run[0], r)))
        for mode, cells, label in lines:
            if not label or PERIOD.match(label) or label.lower() in GENERIC:
                continue
            exprs = Counter(rewrite(forms[p], sheet, p, labels, wb, mode) for p in cells)
            expr, n = exprs.most_common(1)[0]
            c0, r0 = cells[0]
            found.append({"label": label, "expr": expr, "core": core(expr), "workbook": wb.name, "sheet": sheet,
                          "at": f"{xa.col_letters(c0)}{r0}", "along": "down" if mode == "col" else "across",
                          "cells": len(cells), "agree": round(n / len(cells), 3)})
    return found


def is_copy(label, cor):
    """A plain pull of the same label (or a sum of it) from elsewhere: not a definition of its own."""
    base = re.escape(label)
    return bool(re.fullmatch(r"(?:SUM\()?\[%s\](?:@other)?…?\)?" % base, cor, re.I))


def group(lines):
    by = defaultdict(list)
    for ln in lines:
        by[ln["label"].lower()].append(ln)
    out = []
    for key, lns in by.items():
        defs = [l for l in lns if not is_copy(l["label"], l["core"])]
        places = {(l["workbook"], l["sheet"]) for l in lns}
        if len(places) < 2 or not defs:
            continue
        cores = defaultdict(list)
        for l in defs:
            cores[l["core"]].append(l)
        variants = sorted(cores.items(), key=lambda kv: -len({(l["workbook"], l["sheet"]) for l in kv[1]}))
        out.append({"label": lns[0]["label"], "places": len(places), "workbooks": len({l["workbook"] for l in lns}),
                    "copies": len(lns) - len(defs),
                    "variants": [{"definition": k, "places": len({(l["workbook"], l["sheet"]) for l in v}),
                                  "workbooks": len({l["workbook"] for l in v}),
                                  "written": sorted({l["expr"] for l in v})[:3],
                                  "examples": [f'{l["workbook"]} › {l["sheet"]} › {l["at"]} ({l["along"]})' for l in v[:3]]}
                                 for k, v in variants]})
    return sorted(out, key=lambda g: (-(len(g["variants"]) > 1), -g["workbooks"], -g["places"]))


def main(argv):
    as_json = "--json" in argv
    files = [a for a in argv if not a.startswith("--")]
    lines = [l for f in files for l in calculated_lines(f)]
    groups = group(lines)
    differ = [g for g in groups if len(g["variants"]) > 1]
    if as_json:
        print(json.dumps({"calculated_lines": len(lines), "labels_in_2_places": len(groups), "groups": differ}, indent=2))
        return
    print(f"{len(lines)} calculated lines; {len(groups)} labels defined in 2+ places; {len(differ)} with more than "
          f"one definition ({sum(1 for g in differ if g['workbooks'] > 1)} spanning workbooks)\n")
    for g in differ:
        print(f"{g['label']}   ({g['places']} places, {g['workbooks']} workbooks"
              + (f", {g['copies']} copies from elsewhere" if g["copies"] else "") + ")")
        for v in g["variants"]:
            print(f"    = {v['definition']}    [{v['places']} place{'s' if v['places'] != 1 else ''}, "
                  f"{v['workbooks']} workbook{'s' if v['workbooks'] != 1 else ''}]")
            print(f"        e.g. {v['examples'][0]}")
        print()


if __name__ == "__main__":
    main(sys.argv[1:])
