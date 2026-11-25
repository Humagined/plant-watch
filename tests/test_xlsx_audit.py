import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsx_audit as xa  # noqa: E402
import xlsx_fixture as fx  # noqa: E402

KIT = Path(__file__).resolve().parent / "kit"
KIT_PY = Path(sys.executable)


def kinds(result):
    return {(f["workbook"], f["sheet"], f["cell"], f["kind"]) for f in result["findings"]}


def build_plant(tmp_path, through):
    out = tmp_path / "plant"
    subprocess.run([str(KIT_PY), str(KIT / "rolling.py"), "--through", through, "--out", str(out)],
                   check=True, capture_output=True)
    key = json.loads((tmp_path / f"answer_key_{through}.json").read_text())
    return out, key


@pytest.mark.parametrize("through", ["2026-11-11", "2026-11-16", "2026-11-18", "2026-11-23", "2026-11-25"])
def test_matches_answer_key_on_each_day(tmp_path, through):
    plant, key = build_plant(tmp_path, through)
    found = set()
    for wb in sorted((plant / "spreadsheets").glob("*.xlsx")):
        found |= kinds(xa.audit(wb))
    expected = {(f["workbook"], f["sheet"], f["cell"], f["kind"]) for f in key["findings"] if f["category"] == "excel"}
    assert found - expected == set(), "reported something not planted"
    assert expected - found == set(), "missed a planted fault"


def test_clean_days_have_no_shift_log_findings(tmp_path):
    plant, _ = build_plant(tmp_path, "2026-11-11")
    assert xa.audit(plant / "spreadsheets" / "supervisor_shift_log_Nov.xlsx")["findings"] == []


def test_excel_saved_copy_reports_same_cells_with_cached_values(tmp_path):
    plant, _ = build_plant(tmp_path, "2026-11-25")
    plain = xa.audit(plant / "spreadsheets" / "supervisor_shift_log_Nov.xlsx")
    saved = xa.audit(tmp_path / "excel_saved" / "supervisor_shift_log_Nov.xlsx")
    assert kinds(plain) == {(w, s, c, k) for (w, s, c, k) in kinds(saved)}
    cached = {(f["sheet"], f["cell"]): f["cached"] for f in saved["findings"] if f["cached"]}
    assert cached[("Summary", "F17")] == "#DIV/0!"
    assert cached[("Summary", "G17")] == "#REF!"


def test_shared_formulas_expand_and_pattern_break(tmp_path):
    cells = {"A1": "Qty", "B1": "Price", "C1": "Total", "D1": "Note"}
    for r in range(2, 8):
        cells[f"A{r}"] = r
        cells[f"B{r}"] = 10
        cells[f"C{r}"] = {"f": "A2*B2", "shared": ("0", "C2:C7")} if r == 2 else {"f": "", "shared": ("0", None)}
    cells["C5"] = {"f": "A5*B4"}  # bad copy among shared formulas
    p = fx.build(tmp_path / "shared.xlsx", [("Data", cells, None)])
    r = xa.audit(p)
    assert {(f["cell"], f["kind"]) for f in r["findings"]} == {("C5", "pattern_break")}
    wb = xa.Workbook(p)
    assert wb.value("Data", (3, 7))["formula"] == "=A7*B7"


def test_external_link_part_missing_sheet_and_unknown_function(tmp_path):
    p = fx.build(tmp_path / "links.xlsx", [("Main", {
        "A1": {"f": "[1]Plan!B4"}, "A2": {"f": "SUM(Gone!A1:A3)"}, "A3": {"f": "TOTALIZE(B1:B3)"},
        "B1": 1, "B2": 2, "B3": 3}, None)], external="file:///C:/Plans/Budget_2026.xlsx")
    k = {(f["cell"], f["kind"]) for f in xa.audit(p)["findings"]}
    assert ("(workbook links)", "external_link") in k
    assert ("A1", "external_link") in k
    assert ("A2", "missing_sheet_reference") in k
    assert ("A3", "unknown_function") in k


def test_circular_chain_across_sheets_reported_once(tmp_path):
    p = fx.build(tmp_path / "circ.xlsx", [
        ("One", {"A1": {"f": "Two!A1+1"}}, None),
        ("Two", {"A1": {"f": "Three!A1*2"}}, None),
        ("Three", {"A1": {"f": "One!A1-3"}}, None)])
    circ = [f for f in xa.audit(p)["findings"] if f["kind"] == "circular_reference"]
    assert len(circ) == 1 and len(circ[0]["loop"]) == 4


def test_cached_error_without_static_cause_and_structured_refs(tmp_path):
    p = fx.build(tmp_path / "cached.xlsx", [("S", {
        "A1": 5, "A2": {"f": "INDIRECT(\"Z\"&A1)", "e": "#REF!"}, "A3": {"f": "SUM(Table1[Qty])", "v": 0}}, None)])
    r = xa.audit(p)
    assert [(f["cell"], f["kind"], f["cached"]) for f in r["findings"]] == [("A2", "cached_error", "#REF!")]
    assert r["not_checked"][0]["cell"] == "A3"


def test_hidden_sheet_and_hidden_row(tmp_path):
    p = fx.build(tmp_path / "hidden.xlsx", [
        ("Report", {"A1": {"f": "SUM(Data!B2:B5)"}}, None),
        ("Data", {"A1": "Item", "B1": "Qty", "C1": "Unit", "B2": 1, "B3": 2, "B4": 3, "B5": 4}, {"hidden_rows": [3], "state": "hidden"})])
    k = {(f["sheet"], f["cell"], f["kind"]) for f in xa.audit(p)["findings"]}
    assert ("Report", "A1", "hidden_sheet_reference") in k
    assert ("Data", "A3", "hidden_row_in_range") in k


def test_division_by_zero_and_blank_and_defined_name_to_missing_sheet(tmp_path):
    p = fx.build(tmp_path / "div.xlsx", [("S", {"A1": 10, "B1": 0, "C1": {"f": "A1/B1"}, "C2": {"f": "A1/B2"},
                                                "C3": {"f": "A1/2"}}, None)], names={"Rate": "Old!$A$1"})
    k = {(f["cell"], f["kind"]) for f in xa.audit(p)["findings"]}
    assert ("C1", "div_by_blank") in k and ("C2", "div_by_blank") in k
    assert ("Rate", "broken_defined_name") in k
    assert ("C3", "div_by_blank") not in k


def test_text_in_numeric_column_and_duplicate_row(tmp_path):
    cells = {"A1": "Machine", "B1": "Good", "C1": "Scrap"}
    for r in range(2, 14):
        cells[f"A{r}"] = f"M{r}"
        cells[f"B{r}"] = 100 + r
        cells[f"C{r}"] = 1
    cells["B6"] = "n/a"
    cells["A13"], cells["B13"], cells["C13"] = "M12", 112, 1
    k = {(f["cell"], f["kind"]) for f in xa.audit(fx.build(tmp_path / "t.xlsx", [("Log", cells, None)]))["findings"]}
    assert ("B6", "text_in_numeric_column") in k
    assert ("A13", "duplicate_row") in k


def test_circular_references_are_the_same_in_every_process(tmp_path):
    """The cycle search used to walk sets, whose order changes with Python's per-process hash seed,
    so each morning run found different loops and the ledger showed them as new and cleared."""
    import subprocess
    import xlsx_fixture as fx
    cells = {}
    for i in range(1, 9):                       # a tangle: each cell reads the next two
        cells[f"A{i}"] = {"f": f"A{i % 8 + 1}+A{(i + 1) % 8 + 1}+B{i}", "v": 0}
        cells[f"B{i}"] = {"f": f"A{i}*2", "v": 0}
    book = fx.build(tmp_path / "loops.xlsx", [("Sheet One", cells, None)])
    code = ("import sys, json; sys.path.insert(0, sys.argv[1]); import xlsx_audit as xa; "
            "print(json.dumps(sorted((f['cell'], f['loop']) for f in xa.audit(sys.argv[2])['findings'] "
            "if f['kind'] == 'circular_reference')))")
    scripts = str(Path(__file__).resolve().parent.parent / "scripts")
    runs = {subprocess.run([sys.executable, "-c", code, scripts, str(book)], capture_output=True, text=True, check=True,
                           env={**__import__("os").environ, "PYTHONHASHSEED": seed}).stdout for seed in ("1", "2", "3", "4")}
    assert len(runs) == 1 and json.loads(runs.pop())


def test_one_finding_per_tangle_matches_brute_force(tmp_path):
    """A tangle is a set of cells that all reach each other. Every loop through it is one problem, so
    it is one finding; the sample loop must follow real references."""
    import random
    import xlsx_fixture as fx
    for trial in range(60):
        rnd = random.Random(trial)
        n = rnd.randint(2, 30)
        edges = {i: sorted(set(rnd.sample(range(1, n + 1), min(n, rnd.randint(0, 3))))) for i in range(1, n + 1)}
        cells = {f"A{i}": {"f": ("+".join(f"A{j}" for j in edges[i]) or "1") + "+0"} for i in range(1, n + 1)}
        book = fx.build(tmp_path / f"g{trial}.xlsx", [("S", cells, None)])
        reach = {}
        for i in edges:
            seen, todo = set(), list(edges[i])
            while todo:
                x = todo.pop()
                if x not in seen:
                    seen.add(x)
                    todo += edges[x]
            reach[i] = seen
        comps = {frozenset([i] + [j for j in edges if j in reach[i] and i in reach[j]]) for i in edges}
        want = sorted((min(c), len(c)) for c in comps if len(c) > 1 or min(c) in edges[min(c)])
        found = [f for f in xa.audit(book)["findings"] if f["kind"] == "circular_reference"]
        assert sorted((int(f["cell"][1:]), f["tangle_size"]) for f in found) == want, trial
        for f in found:
            hops = [int(c.split("!A")[1]) for c in f["loop"]]
            assert hops[0] == hops[-1] and all(b in edges[a] for a, b in zip(hops, hops[1:])), (trial, hops)
