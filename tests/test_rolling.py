"""plant-watch against the rolling Northfork kit: first run, then day by day."""
import shutil
import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import northfork as nf  # noqa: E402
import ledger  # noqa: E402
import sources as S  # noqa: E402
import watch  # noqa: E402

DICT = Path(__file__).resolve().parent / "fixtures" / "northfork_dictionary.json"


def ident(e):
    """What makes a finding the same as an answer-key entry."""
    if e["category"] == "excel":
        return ("excel", e["workbook"], e["sheet"], e["cell"], e["kind"])
    return (e["category"], e["kind"])


def setup_state(base, through, dictionary=True):
    plant, key = nf.build(base, through)
    cfg = S.load_config(nf.config(base, plant))
    state = base / "state"
    state.mkdir(exist_ok=True)
    if dictionary:
        shutil.copy(DICT, state / "dictionary.json")
    nf.read_reports(plant, state)
    return cfg, state, key


def test_first_run_on_day_13_finds_exactly_the_answer_key(tmp_path):
    cfg, state, key = setup_state(tmp_path, "2026-11-25")
    watch.run(cfg, state, date(2026, 11, 25))
    got = {ident(e): e for e in ledger.load(state).values() if e["active"]}
    want = {}
    for f in key["findings"]:
        k = ("excel", f["workbook"], f["sheet"], f["cell"], f["kind"]) if f["category"] == "excel" else (f["category"], f["kind"])
        want[k] = f
    assert set(got) - set(want) == set(), "reported something not planted"
    assert set(want) - set(got) == set(), "missed a planted problem"
    for k, f in want.items():
        if f["category"] != "excel":
            assert got[k]["onset"] == f["onset"], k
    dc = got[("definition", "definition_change")]
    assert dc["code"] == "CHG" and dc["gap"] == key_gap(key)
    assert dc["evidence"]["notes"], "the dated Notes line should be tied to the change"


def key_gap(key):
    return next(f["gap_min"] for f in key["findings"] if f["kind"] == "definition_change")


def test_day_by_day_each_problem_first_seen_on_its_onset_day(tmp_path):
    state = tmp_path / "state"
    for day in nf.DAYS:
        base = tmp_path / day
        base.mkdir()
        cfg, _, key = setup_state(base, day)
        state.mkdir(exist_ok=True)
        shutil.copy(DICT, state / "dictionary.json")
        nf.read_reports(base / "plant", state)
        watch.run(cfg, state, date.fromisoformat(day))
        active = {ident(e): e for e in ledger.load(state).values() if e["active"]}
        want = {(("excel", f["workbook"], f["sheet"], f["cell"], f["kind"]) if f["category"] == "excel"
                 else (f["category"], f["kind"])): f for f in key["findings"]}
        assert set(active) == set(want), f"{day}: extra {set(active) - set(want)}, missing {set(want) - set(active)}"
        for k, f in want.items():
            assert active[k]["first_seen"] == f["onset"], f"{day}: {k} first seen {active[k]['first_seen']}, onset {f['onset']}"


def test_noise_days_raise_nothing_from_the_shop_data(tmp_path):
    cfg, state, _ = setup_state(tmp_path, "2026-11-11")
    watch.run(cfg, state, date(2026, 11, 11))
    shop = [e for e in ledger.load(state).values() if e["active"] and e.get("workbook") != "maintenance_tracker.xlsx"]
    assert shop == []


def test_findings_clear_when_the_problem_goes_and_status_is_human_only(tmp_path):
    cfg, state, _ = setup_state(tmp_path / "late", "2026-11-25") if (tmp_path / "late").mkdir() is None else None
    watch.run(cfg, state, date(2026, 11, 25))
    any_id = next(e["id"] for e in ledger.load(state).values() if e["kind"] == "circular_reference")
    with pytest.raises(ValueError):
        ledger.set_status(state, any_id, "acknowledged", "", "")
    ledger.set_status(state, any_id, "acknowledged", "Pat Lee", "Plant office fixing the summary sheet")
    (tmp_path / "early").mkdir()
    plant, _ = nf.build(tmp_path / "early", "2026-11-12")
    cfg2 = S.load_config(nf.config(tmp_path / "early", plant))
    r = watch.run(cfg2, state, date(2026, 11, 26))
    e = ledger.load(state)[any_id]
    assert any_id in r["cleared"] and e["active"] is False and e["cleared_on"] == "2026-11-26"
    assert e["status"] == "acknowledged" and e["status_by"] == "Pat Lee"


def test_not_settled_dictionary_phrases_a_disagreement_with_a_question(tmp_path):
    cfg, state, _ = setup_state(tmp_path, "2026-11-25", dictionary=False)
    watch.run(cfg, state, date(2026, 11, 25))
    dc = next(e for e in ledger.load(state).values() if e["kind"] == "definition_change")
    assert "your definition" not in dc["detail"]
    assert "changeovers" in dc["question"].lower()


def test_tampered_reading_is_not_compared(tmp_path):
    import json
    cfg, state, _ = setup_state(tmp_path, "2026-11-25")
    p = state / "extractions" / "Northfork_weekly_OEE_pack_2026-11-20.json"
    ex = json.loads(p.read_text(encoding="utf-8"))
    for f in ex["figures"]:
        if f["entity"] == "total" and f["metric"] == "downtime":
            f["quote"] = "Plant total 7,107 95 2,934 91.0%"   # not what the PDF says
    p.write_text(json.dumps(ex), encoding="utf-8")
    watch.run(cfg, state, date(2026, 11, 25))
    kinds = {e["kind"] for e in ledger.load(state).values() if e["active"]}
    assert "extraction_unverified" in kinds


def test_report_ranks_use_cases_and_places_findings_on_stages(tmp_path):
    import report
    import pdftext
    cfg, state, _ = setup_state(tmp_path, "2026-11-25")
    watch.run(cfg, state, date(2026, 11, 25))
    r = report.render(cfg, state, tmp_path / "r.html", str(tmp_path / "r.pdf"))
    assert [u["title"] for u in r["use_cases"]][0] == "Trusted OEE and downtime reporting"
    assert r["excel_total"] == 32
    by = {e["kind"]: e for e in ledger.load(state).values() if e["category"] != "excel"}
    assert by["definition_change"]["stage"] == "prod::Downtime"
    assert by["spike"]["stage"] == "prod::Quality"
    assert by["key_format_drift"]["stage"] == "prod::Work Order"
    excel = [e for e in ledger.load(state).values() if e["category"] == "excel"]
    assert all(e["stage"] is None for e in excel if e["workbook"] == "maintenance_tracker.xlsx")
    assert {e["stage"] for e in excel if e["workbook"] != "maintenance_tracker.xlsx"} <= {
        "prod::Quality", "prod::Production Reporting", "prod::Downtime"}
    scrap_cells = [e for e in excel if e["sheet"] == "Summary" and e["cell"].startswith("C")]
    assert scrap_cells and all(e["stage"] == "prod::Quality" for e in scrap_cells)
    html_ = (tmp_path / "r.html").read_text(encoding="utf-8")
    assert "Excel errors at a glance" in html_ and "maintenance_tracker.xlsx" in html_
    # The 13 #REF! cells left by the deleted column are one problem in the report, 13 cells in the ledger.
    assert len([e for e in scrap_cells if e["kind"] == "ref_error"]) == 13 and r["excel_problems"] < r["excel_total"]
    assert "Summary › C5:C17" in html_ and "13 adjacent cells" in html_
    text = " ".join(" ".join(p) for p in pdftext.pages(tmp_path / "r.pdf") if p)
    assert "Trusted OEE and downtime reporting" in text and "Summary!F17" in text
