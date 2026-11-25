"""Definitions that live in formulas: the same label calculated differently across a plant's spreadsheets."""
import json
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "plugin" / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import detect  # noqa: E402
import formula_defs as fd  # noqa: E402
import ledger  # noqa: E402
import report  # noqa: E402
import watch  # noqa: E402
import xlsx_fixture as fx  # noqa: E402


def rows(header, n, make):
    """{cell: spec} for a header row and n data rows. make(r) -> {col: spec} for row r."""
    cells = {f"{c}1": h for c, h in header.items()}
    for r in range(2, n + 2):
        cells.update({f"{c}{r}": v for c, v in make(r).items()})
    return cells


def plant(folder):
    folder.mkdir(parents=True, exist_ok=True)
    # Mill cell: Availability = Run / Planned. Hours/Day = Run Hours / Days.
    fx.build(folder / "mill.xlsx", [("Mill", rows(
        {"A": "Machine", "B": "Run Hours", "C": "Planned Hours", "D": "Availability", "E": "Days", "F": "Hours/Day"}, 4,
        lambda r: {"A": f"M{r}", "B": 300 + r, "C": 400, "D": {"f": f"B{r}/C{r}"}, "E": 20, "F": {"f": f"B{r}/E{r}"}}), None),
        # Same definitions on a summary sheet with another layout and an IFERROR: not a different definition.
        ("Mill summary", rows(
            {"A": "Machine", "B": "Notes", "C": "Planned Hours", "D": "Run Hours", "E": "Availability", "F": "Days",
             "G": "Hours/Day"}, 4,
            lambda r: {"A": f"M{r}", "B": "ok", "C": 400, "D": 300 + r, "E": {"f": f"IFERROR(D{r}/C{r},0)"}, "F": 20,
                       "G": {"f": f"D{r}/F{r}"}}), None),
        # Copies from another sheet and month columns: not definitions of their own.
        ("Rollup", rows({"A": "Machine", "B": "Availability", "C": "Jan", "D": "Feb"}, 4,
                        lambda r: {"A": f"M{r}", "B": {"f": f"Mill!D{r}"}, "C": {"f": f"Mill!B{r}*2"},
                                   "D": {"f": f"Mill!B{r}*3"}}), None)])
    # Turn cell: Availability takes planned maintenance out of planned hours. A different definition.
    fx.build(folder / "turn.xlsx", [("Turn", rows(
        {"A": "Machine", "B": "Run Hours", "C": "Planned Hours", "D": "PM Hours", "E": "Availability", "F": "Days",
         "G": "Hours/Day"}, 4,
        lambda r: {"A": f"T{r}", "B": 250 + r, "C": 400, "D": 16, "E": {"f": f"B{r}/(C{r}-D{r})"}, "F": 20,
                   "G": {"f": f"B{r}/F{r}"}}), None),
        # Customer breakdown: SUMIF picks a different customer per row; the definition is the same.
        ("By customer", rows({"A": "Customer", "B": "Run Hours", "C": "Hours by customer"}, 4,
                             lambda r: {"A": f"C{r}", "B": 10 * r, "C": {"f": f"SUMIF($A$2:$A$5,A{r},$B$2:$B$5)"}}), None),
        # A month column here too, calculated differently from the mill's: months are columns, not metrics.
        ("Turn rollup", rows({"A": "Machine", "B": "Jan"}, 4, lambda r: {"A": f"T{r}", "B": {"f": f"Turn!B{r}*5"}}), None)])
    return folder


def test_same_label_different_formula_is_one_finding_and_layout_is_not(tmp_path):
    folder = plant(tmp_path / "plant")
    groups = {g["label"]: g for g in detect.formula_groups(sorted(folder.glob("*.xlsx")))}
    av = groups["Availability"]
    assert sorted(v["definition"] for v in av["variants"]) == \
        ["[Run Hours]/([Planned Hours]-[PM Hours])", "[Run Hours]/[Planned Hours]"]
    assert {v["definition"]: v["places"] for v in av["variants"]}["[Run Hours]/[Planned Hours]"] == 2   # two layouts, one definition
    assert av["copies"] == 1                                                                        # the Rollup pull
    assert len(groups["Hours/Day"]["variants"]) == 1                                                # same everywhere
    assert "Jan" not in groups and "Feb" not in groups                                              # period columns
    assert "Hours by customer" not in groups or len(groups["Hours by customer"]["variants"]) == 1    # selectors

    found = detect.formula_definitions({"sources": []}, sorted(folder.glob("*.xlsx")), {})
    assert [f["label"] for f in found] == ["Availability"]
    f = found[0]
    assert f["kind"] == "formula_definition_conflict" and f["category"] == "definition"
    assert f["metric"] == "availability" and f["stage"] == detect.stage_of("availability")
    assert f["workbooks"] == 2 and "Which of these" in f["question"]


def test_a_confirmed_definition_reports_only_the_places_that_differ(tmp_path):
    folder = plant(tmp_path / "plant")
    mine = {"availability": {"status": "confirmed", "definition": "[Run Hours]/[Planned Hours]"}}
    f = detect.formula_definitions({"sources": []}, sorted(folder.glob("*.xlsx")), mine)[0]
    assert "1 place(s) calculate it differently" in f["detail"] and "PM Hours" in f["detail"]
    assert [v["is_yours"] for v in f["evidence"]["variants"]].count(True) == 1
    everywhere = {"availability": {"status": "confirmed", "definition": "[Run Hours]/([Planned Hours]-[PM Hours])"}}
    assert detect.formula_definitions({"sources": []}, sorted(folder.glob("*.xlsx")), everywhere)[0]["detail"].count("place(s)") >= 1
    # Once every place matches, there's nothing to report.
    (folder / "turn.xlsx").unlink()
    assert detect.formula_definitions({"sources": []}, sorted(folder.glob("*.xlsx")), mine) == []


def test_run_report_and_inventory_carry_it_with_a_stable_identity(tmp_path):
    folder = plant(tmp_path / "plant")
    cfg = {"plant": "Test plant", "folder": str(folder), "sources": []}
    state = tmp_path / "state"
    watch.run(cfg, state, date(2026, 11, 20))
    first = [e for e in ledger.load(state).values() if e["kind"] == "formula_definition_conflict"]
    watch.run(cfg, state, date(2026, 11, 21))
    second = [e for e in ledger.load(state).values() if e["kind"] == "formula_definition_conflict"]
    assert len(first) == 1 and [e["id"] for e in first] == [e["id"] for e in second] and second[0]["active"]

    out = tmp_path / "r.html"
    report.render(cfg, state, out)
    html = out.read_text(encoding="utf-8")
    assert html.index("Same name, different calculation") < html.index("Excel errors at a glance")
    assert "[Run Hours]/([Planned Hours]-[PM Hours])" in html

    inv = watch.inventory(cfg)
    assert [g["label"] for g in inv["formula_definitions"]] == ["Availability"]
    assert inv["formula_definitions"][0]["metric"] == "availability"


def test_rewrite_names_cells_by_the_labels_they_read():
    assert fd.core("IFERROR([A]/[B],0)") == "[A]/[B]"
    assert fd.core('SUMIF([Customer],[Acme],[Total])') == "SUMIF(…,…,[Total])"
    assert fd.core("[B]+[A]") == fd.core("[A]+[B]")
    assert fd.core("SUM([X](row-3),[X](row-9),[X](row-12))") == "SUM([X]…)"
    assert fd.PERIOD.match("Sep") and fd.PERIOD.match("Q3") and not fd.PERIOD.match("Scrap")


def test_labels_land_on_the_right_metric_and_every_metric_on_a_real_stage():
    cases = {
        "Orders QTD% to Budget": "budget_attainment",   # an attainment, not orders: the specific metric wins
        "Daily Order Target": "order_target",            # a target, not orders
        "Weekly Orders": "orders", "Shipped MTD": "shipments", "Budget": "budget", "Reforecast": "forecast",
        "Days in Month": "working_days", "Hours/Day": "capacity_hours", "Standard Hours": "capacity_hours",
        "Performance %": "performance",
        # shop-floor labels keep the metric they had before the new ones were added
        "Good": "good", "Scrap": "scrap", "Downtime (min)": "downtime", "Availability": "availability",
        "Work Order": None, "Machine": None,
    }
    assert {k: detect.metric_for_label(k) for k in cases} == cases
    streams = json.loads((Path(detect.DATA) / "value_streams.json").read_text(encoding="utf-8"))["streams"]
    stages = {st["key"] for s in streams for st in s["stages"]}
    for m, info in detect.METRICS.items():
        assert info["stage"] in stages, m
        assert info["question"] and info["synonyms"] and info["label"], m


def test_report_shows_a_fault_filled_down_a_column_once():
    def f(cell, kind="cached_error", formula="=B{r}/C{r}", cached="#DIV/0!", **kw):
        r = int(cell[1:])
        return {"workbook": "w.xlsx", "sheet": "S", "cell": cell, "kind": kind, "severity": 2, "problem": "p", "fix": "x",
                "formula": formula.format(r=r) if formula else None, "cached": cached, **kw}
    entries = ([f(f"D{r}") for r in range(2, 7)]                 # D2:D6, one run
               + [f("D9")]                                       # a gap: its own problem
               + [f("D7", formula="=B{r}/E{r}")]                 # a different formula: separate
               + [f("D8", cached="#N/A")]                        # shows something else: separate
               + [f(f"A{r}", kind="hidden_row_in_range", formula=None, cached=None, used_by=["S!F20"], used_by_count=1)
                  for r in (11, 12, 13)]
               + [f("A15", kind="hidden_row_in_range", formula=None, cached=None, used_by=["S!F20"], used_by_count=1)])
    gs = {g["where"]: g for g in report.excel_groups(entries)}
    assert set(gs) == {"D2:D6", "D9", "D7", "D8", "rows 11–13", "A15"}
    assert gs["D2:D6"]["count"] == 5 and gs["D9"]["count"] == 1 and gs["D9"]["problem"] == "p"
    assert "(3 rows) are hidden" in gs["rows 11–13"]["problem"]
    assert sum(g["count"] for g in gs.values()) == len(entries)       # every cell is in exactly one problem
